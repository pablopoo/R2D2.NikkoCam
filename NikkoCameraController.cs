using System.ComponentModel;

namespace R2D2.NikkoCam;

// High-level camera entrypoint for the app. It discovers the WinUSB receiver,
// picks the streaming-capable interface, and delegates the actual USB/video work
// to LabSession.
internal sealed class NikkoCameraController : IDisposable
{
    private const ushort NikkoVid = 0x6000;
    private const ushort NikkoPid = 0x0001;
    private const byte VideoPipeId = 0x81;
    private const byte PreferredVideoAlt = 1;
    private static readonly Guid PreferredInterfaceGuid = new("E91C5AE1-8C81-48E1-AF86-1626C8E4703A");

    private readonly LabSession _session = new();

    // Expose only interface paths that look like the real video streaming endpoint.
    internal IReadOnlyList<string> DiscoverDevicePaths()
    {
        var preferred = UsbDeviceDiscovery.FindDevicePaths(NikkoVid, NikkoPid, PreferredInterfaceGuid)
            .Where(SupportsExpectedVideoInterface)
            .ToArray();

        if (preferred.Length > 0)
        {
            return preferred;
        }

        return EnumerateCandidatePaths()
            .Where(SupportsExpectedVideoInterface)
            .ToArray();
    }

    // Start the fixed Nikko camera path. The receiver can expose multiple device
    // interfaces, so we try the selected one first and then fall back to other
    // candidates that match the same hardware.
    internal async Task StartAsync(string devicePath, Action<PreviewFrameResult> onFrame, CancellationToken cancellationToken)
    {
        var startupCandidates = EnumerateStartupCandidates(devicePath).ToArray();
        if (!startupCandidates.Any(IsDeviceControlResponsive))
        {
            throw new InvalidOperationException(
                "The Nikko receiver is not responding to USB control requests. Unplug and reconnect the video receiver, then try again.");
        }

        Exception? lastError = null;
        foreach (var candidatePath in startupCandidates)
        {
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    await _session.InitializeAsync(candidatePath, cancellationToken);
                    await _session.StartPreviewAsync(onFrame, cancellationToken);
                    return;
                }
                catch (Exception ex) when (IsAlternateSettingFailure(ex))
                {
                    lastError = ex;
                    if (attempt == 1)
                    {
                        break;
                    }

                    await Task.Delay(120, cancellationToken);
                }
            }
        }

        if (lastError is not null)
        {
            throw lastError;
        }

        throw new InvalidOperationException("No compatible Nikko video interface was found.");
    }

    internal Task SetChannelAsync(int channelNumber, CancellationToken cancellationToken) =>
        _session.SetToyChannelAsync(channelNumber, cancellationToken);

    internal Task StopAsync() => _session.DisconnectAsync();

    public void Dispose()
    {
        _session.Dispose();
    }

    private static bool IsAlternateSettingFailure(Exception ex)
    {
        if (ex is Win32Exception win32Ex &&
            win32Ex.Message.Contains("WinUsb_SetCurrentAlternateSetting failed", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return ex.Message.Contains("WinUsb_SetCurrentAlternateSetting failed", StringComparison.OrdinalIgnoreCase);
    }

    // The receiver may show up both on the preferred WinUSB interface and on the
    // generic USB device interface. We keep both internally for startup fallback,
    // but the UI only shows the preferred streaming interface when available.
    private IReadOnlyList<string> EnumerateCandidatePaths()
    {
        var result = new List<string>();
        AddDistinct(result, UsbDeviceDiscovery.FindDevicePaths(NikkoVid, NikkoPid, PreferredInterfaceGuid));
        AddDistinct(result, UsbDeviceDiscovery.FindDevicePaths(NikkoVid, NikkoPid, UsbDeviceDiscovery.GenericUsbDeviceInterfaceGuid));
        return result;
    }

    // Retry startup on equivalent interfaces for the same physical receiver. This
    // helps when Windows exposes a stale or non-streaming path first.
    private IEnumerable<string> EnumerateStartupCandidates(string selectedPath)
    {
        var all = EnumerateCandidatePaths();
        if (!string.IsNullOrWhiteSpace(selectedPath))
        {
            yield return selectedPath;
        }

        foreach (var path in all.Where(SupportsExpectedVideoInterface))
        {
            if (!string.Equals(path, selectedPath, StringComparison.OrdinalIgnoreCase))
            {
                yield return path;
            }
        }

        foreach (var path in all)
        {
            if (!string.Equals(path, selectedPath, StringComparison.OrdinalIgnoreCase))
            {
                yield return path;
            }
        }
    }

    private static void AddDistinct(List<string> target, IReadOnlyList<string> source)
    {
        foreach (var path in source)
        {
            if (!target.Contains(path, StringComparer.OrdinalIgnoreCase))
            {
                target.Add(path);
            }
        }
    }

    // A valid video path must expose the isoch streaming pipe used by the receiver.
    private static bool SupportsExpectedVideoInterface(string devicePath)
    {
        try
        {
            using var device = WinUsbDevice.Open(devicePath);
            return device.GetAlternateSettings().Any(setting =>
                setting.Descriptor.AlternateSetting == PreferredVideoAlt &&
                setting.Pipes.Any(pipe =>
                    pipe.PipeId == VideoPipeId &&
                    pipe.PipeType == NativeMethods.UsbdPipeType.Isochronous &&
                    ((pipe.MaximumBytesPerInterval ?? 0) > 0 || pipe.MaximumPacketSize > 0)));
        }
        catch
        {
            return false;
        }
    }

    // Before doing the full startup sequence, verify that EP0 control transfers work.
    // If this fails, the receiver usually needs a physical unplug/replug.
    private static bool IsDeviceControlResponsive(string devicePath)
    {
        try
        {
            using var device = WinUsbDevice.Open(devicePath);
            _ = device.ControlTransferIn(0x80, 0x08, 0, 0, 1);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
