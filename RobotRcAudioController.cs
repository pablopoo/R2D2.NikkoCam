using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Threading;
using NAudio.CoreAudioApi;
using NAudio.Wave;

namespace R2D2.NikkoCam;

// Owns the RC audio path. The Nikko receiver exposes a USB sound card, and the
// robot interprets DTMF tones sent to that output as movement/camera commands.
internal sealed class RobotRcAudioController : IDisposable
{
    internal const ushort PreferredVendorId = 0x04D9;
    internal const ushort PreferredProductId = 0x2821;
    internal const string PreferredDeviceNameToken = "NIKKO R/C";
    internal const string DeviceOverrideEnvironmentVariable = "R2D2_RC_AUDIO_DEVICE";

    private readonly object _sync = new();
    private readonly DtmfWaveProvider _provider = new();
    private static readonly TimeSpan RecoverableRetryDelay = TimeSpan.FromMilliseconds(180);
    private const int RecoverableAttemptCount = 3;

    private WasapiOut? _output;
    private MMDevice? _device;
    private bool _initialized;
    private int _activeDigitCode = -1;

    internal bool IsToneActive => Volatile.Read(ref _activeDigitCode) >= 0;
    internal bool IsConnected
    {
        get
        {
            lock (_sync)
            {
                return _initialized && _device is not null && _output is not null;
            }
        }
    }

    internal bool RefreshConnectionState()
    {
        lock (_sync)
        {
            return RefreshConnectionStateCore();
        }
    }

    internal string Connect()
    {
        lock (_sync)
        {
            EnsureInitialized();
            return _device!.FriendlyName;
        }
    }

    internal void StartTone(char digit)
    {
        lock (_sync)
        {
            StartToneCore(digit, allowReinitialize: true);
        }
    }

    internal void Disconnect()
    {
        lock (_sync)
        {
            ResetAudioDevice();
        }
    }

    internal void StopTone()
    {
        lock (_sync)
        {
            _provider.ClearDigit();
            Volatile.Write(ref _activeDigitCode, -1);
        }
    }

    public void Dispose()
    {
        lock (_sync)
        {
            _provider.ClearDigit();
            _output?.Stop();
            _output?.Dispose();
            _output = null;
            _device?.Dispose();
            _device = null;
            _initialized = false;
            Volatile.Write(ref _activeDigitCode, -1);
        }
    }

    // Select the receiver's audio output and open a shared WASAPI stream on it.
    // If the receiver has been unplugged/replugged, this recreates the session.
    private void EnsureInitialized()
    {
        using var enumerator = new MMDeviceEnumerator();
        var preferredDevice = SelectPreferredDevice(
            enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToArray());

        if (preferredDevice is null)
        {
            ResetAudioDevice();
            throw new InvalidOperationException(
                $"RC audio device not found. Expected an active render device tied to USB VID_{PreferredVendorId:X4}&PID_{PreferredProductId:X4}.");
        }

        if (CanReuseCurrentDevice(preferredDevice))
        {
            preferredDevice.Dispose();
            return;
        }

        ResetAudioDevice();
        _device = preferredDevice;
        _output = new WasapiOut(preferredDevice, AudioClientShareMode.Shared, useEventSync: true, latency: 80);
        _output.Init(_provider);
        _output.Play();
        _provider.ClearDigit();
        _initialized = true;
    }

    // Start one DTMF tone, retrying on the common COM errors Windows returns while
    // an endpoint is being re-enumerated after unplug/replug.
    private void StartToneCore(char digit, bool allowReinitialize)
    {
        var attemptsRemaining = allowReinitialize ? RecoverableAttemptCount : 1;
        Exception? lastRecoverable = null;

        while (attemptsRemaining-- > 0)
        {
            try
            {
                EnsureInitialized();
                _provider.SetDigit(digit);
                Volatile.Write(ref _activeDigitCode, digit);
                _output!.Play();
                return;
            }
            catch (Exception ex) when (IsRecoverableAudioFailure(ex))
            {
                lastRecoverable = ex;
                ResetAudioDevice();
                if (attemptsRemaining > 0)
                {
                    Thread.Sleep(RecoverableRetryDelay);
                }
            }
        }

        if (lastRecoverable is not null)
        {
            throw lastRecoverable;
        }

        throw new InvalidOperationException("RC audio device initialization failed.");
    }

    // Rank active render endpoints so the RC path follows the Nikko receiver and
    // not another unrelated speaker/headphone device on the machine.
    private static MMDevice? SelectPreferredDevice(IReadOnlyList<MMDevice> devices)
    {
        var overrideSelector = Environment.GetEnvironmentVariable(DeviceOverrideEnvironmentVariable)?.Trim();
        if (!string.IsNullOrWhiteSpace(overrideSelector))
        {
            var overridden = devices.FirstOrDefault(candidate => MatchesOverride(candidate, overrideSelector));
            if (overridden is not null)
            {
                return overridden;
            }
        }

        return devices
            .Where(candidate => candidate.State == DeviceState.Active)
            .Select(candidate => new { Device = candidate, Score = ComputeMatchScore(candidate) })
            .Where(match => match.Score > 0)
            .OrderByDescending(match => match.Score)
            .Select(match => match.Device)
            .FirstOrDefault();
    }

    private static bool MatchesOverride(MMDevice candidate, string selector)
    {
        return string.Equals(candidate.ID, selector, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(NormalizeDeviceId(candidate.ID), NormalizeDeviceId(selector), StringComparison.OrdinalIgnoreCase) ||
            candidate.FriendlyName.Contains(selector, StringComparison.OrdinalIgnoreCase);
    }

    // The safest match is the underlying USB VID/PID. Friendly name is only a
    // secondary hint because localized Windows installs can rename the endpoint.
    private static int ComputeMatchScore(MMDevice candidate)
    {
        var preferredUsbId = $"VID_{PreferredVendorId:X4}&PID_{PreferredProductId:X4}";
        var score = 0;

        if (DevicePropertyContains(candidate, preferredUsbId))
        {
            score += 1000;
        }

        if (DevicePropertyContains(candidate, $"USB\\{preferredUsbId}&MI_00"))
        {
            score += 100;
        }

        var friendlyName = candidate.FriendlyName;
        if (friendlyName.Contains(PreferredDeviceNameToken, StringComparison.OrdinalIgnoreCase))
        {
            score += 50;
            if (friendlyName.Contains("(2-", StringComparison.OrdinalIgnoreCase))
            {
                score += 10;
            }
        }

        return score;
    }

    // Walk the MMDevice property store because the USB identity can appear on
    // different properties depending on the Windows version and audio stack.
    private static bool DevicePropertyContains(MMDevice candidate, string needle)
    {
        if (candidate.ID.Contains(needle, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        try
        {
            var properties = candidate.Properties;
            for (var index = 0; index < properties.Count; index++)
            {
                object? value;
                try
                {
                    value = properties.GetValue(index).Value;
                }
                catch
                {
                    continue;
                }

                if (value is string text &&
                    text.Contains(needle, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        catch
        {
            return false;
        }

        return false;
    }

    private void ResetAudioDevice()
    {
        _provider.ClearDigit();
        _output?.Stop();
        _output?.Dispose();
        _output = null;
        _device?.Dispose();
        _device = null;
        _initialized = false;
        Volatile.Write(ref _activeDigitCode, -1);
    }

    private bool CanReuseCurrentDevice(MMDevice preferredDevice)
    {
        if (!_initialized || _device is null || _output is null)
        {
            return false;
        }

        if (!string.Equals(_device.ID, preferredDevice.ID, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        try
        {
            return _device.State == DeviceState.Active &&
                   _output.PlaybackState != PlaybackState.Stopped;
        }
        catch
        {
            return false;
        }
    }

    // Used by the UI timer to detect that the RC endpoint disappeared or changed
    // underneath us, so the Connect/Disconnect state stays honest.
    private bool RefreshConnectionStateCore()
    {
        if (!_initialized)
        {
            return false;
        }

        if (_device is null || _output is null)
        {
            ResetAudioDevice();
            return false;
        }

        try
        {
            if (_device.State != DeviceState.Active)
            {
                ResetAudioDevice();
                return false;
            }

            using var enumerator = new MMDeviceEnumerator();
            using var preferredDevice = SelectPreferredDevice(
                enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active).ToArray());

            if (preferredDevice is null ||
                !string.Equals(_device.ID, preferredDevice.ID, StringComparison.OrdinalIgnoreCase))
            {
                ResetAudioDevice();
                return false;
            }

            return true;
        }
        catch
        {
            ResetAudioDevice();
            return false;
        }
    }

    private static bool IsRecoverableAudioFailure(Exception ex)
    {
        if (ex is COMException comEx)
        {
            return comEx.HResult == unchecked((int)0x88890004) ||
                   comEx.HResult == unchecked((int)0x88890008) ||
                   comEx.HResult == unchecked((int)0x8889000A);
        }

        return false;
    }

    private static string NormalizeDeviceId(string deviceId)
    {
        const string prefix = @"SWD\MMDEVAPI\";
        return deviceId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            ? deviceId[prefix.Length..]
            : deviceId;
    }

    // Generates telephone-style DTMF in software: one low-group sine and one
    // high-group sine, plus a short envelope to avoid speaker clicks.
    private sealed class DtmfWaveProvider : IWaveProvider
    {
        private const int SampleRate = 48000;
        private const short PeakAmplitude = 12000;
        private const double LowGroupGain = 0.95;
        private const double HighGroupGain = 0.75;
        private const int FadeMilliseconds = 8;

        private readonly WaveFormat _waveFormat = new(SampleRate, 16, 1);

        private double _lowPhase;
        private double _highPhase;
        private int _digitCode = -1;
        private int _fadeSamplesRemaining;
        private bool _fadingIn;
        private bool _fadingOut;

        public WaveFormat WaveFormat => _waveFormat;

        public int Read(byte[] buffer, int offset, int count)
        {
            var bytesWritten = 0;
            var sampleCount = count / 2;
            var digitCode = Volatile.Read(ref _digitCode);
            var (lowFrequency, highFrequency) = ResolveFrequencies(digitCode);

            for (var sampleIndex = 0; sampleIndex < sampleCount; sampleIndex++)
            {
                short sampleValue;
                if (lowFrequency > 0 && highFrequency > 0)
                {
                    var envelope = GetEnvelopeGain();
                    var lowSample = Math.Sin(_lowPhase) * LowGroupGain;
                    var highSample = Math.Sin(_highPhase) * HighGroupGain;
                    var mixedSample = (lowSample + highSample) * 0.5 * envelope;
                    sampleValue = (short)Math.Clamp(mixedSample * PeakAmplitude, short.MinValue, short.MaxValue);

                    _lowPhase = AdvancePhase(_lowPhase, lowFrequency);
                    _highPhase = AdvancePhase(_highPhase, highFrequency);
                }
                else
                {
                    sampleValue = 0;
                }

                Unsafe.WriteUnaligned(ref buffer[offset + bytesWritten], sampleValue);
                bytesWritten += 2;
            }

            return sampleCount * 2;
        }

        internal void SetDigit(char digit)
        {
            Volatile.Write(ref _digitCode, digit);
            _fadingIn = true;
            _fadingOut = false;
            _fadeSamplesRemaining = FadeMilliseconds * SampleRate / 1000;
        }

        internal void ClearDigit()
        {
            if (Volatile.Read(ref _digitCode) < 0)
            {
                return;
            }

            _fadingOut = true;
            _fadingIn = false;
            _fadeSamplesRemaining = FadeMilliseconds * SampleRate / 1000;
        }

        private static double AdvancePhase(double phase, int frequency)
        {
            phase += 2 * Math.PI * frequency / SampleRate;
            if (phase >= 2 * Math.PI)
            {
                phase -= 2 * Math.PI;
            }

            return phase;
        }

        // Small fade in/out so button presses do not create hard clicks.
        private double GetEnvelopeGain()
        {
            if (_fadingIn)
            {
                if (_fadeSamplesRemaining <= 0)
                {
                    _fadingIn = false;
                    return 1.0;
                }

                var totalSamples = FadeMilliseconds * SampleRate / 1000;
                var position = totalSamples - _fadeSamplesRemaining;
                _fadeSamplesRemaining--;
                return Math.Clamp((double)position / totalSamples, 0.0, 1.0);
            }

            if (_fadingOut)
            {
                if (_fadeSamplesRemaining <= 0)
                {
                    _fadingOut = false;
                    Volatile.Write(ref _digitCode, -1);
                    return 0.0;
                }

                var totalSamples = FadeMilliseconds * SampleRate / 1000;
                var position = totalSamples - _fadeSamplesRemaining;
                _fadeSamplesRemaining--;
                return Math.Clamp(1.0 - (double)position / totalSamples, 0.0, 1.0);
            }

            return 1.0;
        }

        // Standard DTMF keypad map used by the robot receiver.
        private static (int LowFrequency, int HighFrequency) ResolveFrequencies(int digitCode) => digitCode switch
        {
            '1' => (697, 1209),
            '2' => (697, 1336),
            '3' => (697, 1477),
            '4' => (770, 1209),
            '5' => (770, 1336),
            '6' => (770, 1477),
            '7' => (852, 1209),
            '8' => (852, 1336),
            '9' => (852, 1477),
            '0' => (941, 1336),
            '*' => (941, 1209),
            '#' => (941, 1477),
            _ => (0, 0),
        };
    }
}
