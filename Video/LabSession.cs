namespace R2D2.NikkoCam;

// It is responsible for the device startup sequence,
// channel writes, the isoch preview loop, and periodic keepalive/status polling.
internal sealed class LabSession : IDisposable
{
    private const int XpTraceTickMs = 16;
    private const int FixedWidth = 352;
    private const int FixedHeight = 240;
    private const int PacketCount = 64;
    private const uint PipeTimeoutMs = 2500;
    private const byte PipeId = 0x81;
    private const byte InitialAlt = 0;
    private const byte TargetAlt = 1;
    private const int StatusPollIntervalMs = 250;

    private readonly SemaphoreSlim _ioGate = new(1, 1);
    private readonly RollingPreviewAssembler _previewAssembler = new();
    private WinUsbDevice? _device;
    private CancellationTokenSource? _previewCts;
    private Task? _previewTask;
    private DateTime _nextStatusPollUtc = DateTime.MinValue;
    private bool _initialized;

    // Bring the receiver into the same USB state the original app expected:
    // alt 0, fixed startup control sequence, then alt 1 for isoch video.
    internal async Task InitializeAsync(string devicePath, CancellationToken cancellationToken)
    {
        await ReleaseDeviceAsync(cancellationToken);

        await _ioGate.WaitAsync(cancellationToken);
        try
        {
            _device = WinUsbDevice.Open(devicePath);
            _initialized = false;
            _nextStatusPollUtc = DateTime.MinValue;

            _previewAssembler.Reset();
            _previewAssembler.SetPreferredField(-1);

            _device.SetCurrentAlternateSetting(InitialAlt);
            RunSequence(RecoveredDriverSequences.FixedStartupSequence, continueOnError: true);
            Thread.Sleep(XpTraceTickMs);
            _device.SetCurrentAlternateSetting(TargetAlt);

            _initialized = true;
        }
        finally
        {
            _ioGate.Release();
        }
    }

    // Fully disconnect the receiver from this process: stop preview, try to move
    // the interface back to the idle alternate setting, and release the WinUSB
    // handle so the next connect starts from a clean state.
    internal Task DisconnectAsync() => ReleaseDeviceAsync(CancellationToken.None);

    // Run the fixed live preview loop. Each iteration reads one batch of isoch
    // packets, updates the rolling raster, publishes a bitmap, and issues the
    // periodic 0x0300 keepalive/status read used by the original stack.
    internal async Task StartPreviewAsync(Action<PreviewFrameResult> onFrame, CancellationToken cancellationToken)
    {
        if (_device is null || !_initialized)
        {
            throw new InvalidOperationException("Initialize the device before starting preview.");
        }

        await StopPreviewAsync();
        _previewCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var previewToken = _previewCts.Token;

        _previewTask = Task.Run(async () =>
        {
            var firstFrame = true;
            _nextStatusPollUtc = DateTime.MinValue;

            while (!previewToken.IsCancellationRequested)
            {
                await _ioGate.WaitAsync(previewToken);
                try
                {
                    if (_device is null)
                    {
                        return;
                    }

                    var isInitialStart = firstFrame;
                    var result = _device.ReadIsochronousPipeAsap(
                        PipeId,
                        PacketCount,
                        PipeTimeoutMs,
                        continueStream: false,
                        afterSubmit: firstFrame
                            ? () =>
                            {
                                if (isInitialStart)
                                {
                                    Thread.Sleep(XpTraceTickMs);
                                }

                                RunSequence(RecoveredDriverSequences.PreviewStartSequence, continueOnError: true);
                            }
                            : null);
                    firstFrame = false;

                    var packetDataCount = result.Packets.Count(static packet => packet.Status == 0 && packet.Length > 0);
                    if (packetDataCount == 0)
                    {
                        PollStatusIfDue();
                        continue;
                    }

                    var frame = _previewAssembler.PushAndBuild(
                        result,
                        FixedWidth,
                        FixedHeight);
                    if (frame is null)
                    {
                        PollStatusIfDue();
                        continue;
                    }

                    _previewAssembler.SetPreferredField(frame.Field);
                    onFrame(frame);
                    PollStatusIfDue();
                }
                finally
                {
                    _ioGate.Release();
                }

                await Task.Delay(30, previewToken);
            }
        }, previewToken);
    }

    internal async Task StopPreviewAsync()
    {
        if (_previewCts is null)
        {
            return;
        }

        _previewCts.Cancel();
        try
        {
            if (_previewTask is not null)
            {
                await _previewTask;
            }
        }
        catch (OperationCanceledException)
        {
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("zero-bandwidth", StringComparison.OrdinalIgnoreCase))
        {
        }
        finally
        {
            _previewTask = null;
            _previewCts.Dispose();
            _previewCts = null;
            _previewAssembler.Reset();
        }
    }

    // The original channel menu reduced to two vendor-control bits:
    // 0102 selects the high bit, 0103 selects the low bit.
    internal async Task SetToyChannelAsync(int channelNumber, CancellationToken cancellationToken)
    {
        if (channelNumber is < 1 or > 4)
        {
            throw new ArgumentOutOfRangeException(nameof(channelNumber), channelNumber, "Channel must be CH1..CH4.");
        }

        var channelIndex = channelNumber - 1;
        var command0102 = (ushort)((channelIndex >> 1) & 0x1);
        var command0103 = (ushort)(channelIndex & 0x1);

        await _ioGate.WaitAsync(cancellationToken);
        try
        {
            if (_device is null)
            {
                throw new InvalidOperationException("No device is open.");
            }

            _ = _device!.ControlTransferOut(0x40, 0x03, 0x0102, command0102, []);
            await Task.Delay(10, cancellationToken);
            _ = _device.ControlTransferOut(0x40, 0x03, 0x0103, command0103, []);
            await Task.Delay(200, cancellationToken);
        }
        finally
        {
            _ioGate.Release();
        }
    }

    public void Dispose()
    {
        ReleaseDeviceAsync(CancellationToken.None).GetAwaiter().GetResult();
        _ioGate.Dispose();
    }

    // Lightweight keepalive/status poll. The UI does not surface the value, but the
    // receiver behaves better when we keep issuing the same read cadence as XP.
    private void PollStatusIfDue()
    {
        if (_device is null)
        {
            return;
        }

        var now = DateTime.UtcNow;
        if (_nextStatusPollUtc != DateTime.MinValue && now < _nextStatusPollUtc)
        {
            return;
        }

        _ = _device.ControlTransferIn(0xC0, 0x03, 0x0300, 0x0000, 1);
        _nextStatusPollUtc = now.AddMilliseconds(StatusPollIntervalMs);
    }

    // Execute the recovered XP-derived startup writes. The receiver tolerates
    // some failures, so the startup path intentionally continues when requested.
    private void RunSequence(IReadOnlyList<RecoveredDriverSequences.Step> sequence, bool continueOnError)
    {
        for (var i = 0; i < sequence.Count; i++)
        {
            var step = sequence[i];
            try
            {
                _ = _device!.ControlTransferOut(0x40, step.Request, step.Value, step.Index, []);
            }
            catch when (continueOnError)
            {
            }

            if (step.DelayAfterMs > 0)
            {
                Thread.Sleep(step.DelayAfterMs);
            }
        }
    }

    private async Task ReleaseDeviceAsync(CancellationToken cancellationToken)
    {
        await StopPreviewAsync();

        await _ioGate.WaitAsync(cancellationToken);
        try
        {
            if (_device is not null)
            {
                try
                {
                    _device.SetCurrentAlternateSetting(InitialAlt);
                }
                catch
                {
                }

                _device.Dispose();
                _device = null;
                _initialized = false;
            }

            _nextStatusPollUtc = DateTime.MinValue;
            _previewAssembler.Reset();
        }
        finally
        {
            _ioGate.Release();
        }

        // Give Windows a brief moment to settle the interface before a new open.
        await Task.Delay(75, cancellationToken);
    }
}
