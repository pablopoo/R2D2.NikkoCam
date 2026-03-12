using Microsoft.Win32.SafeHandles;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace R2D2.NikkoCam;

internal sealed class WinUsbDevice : IDisposable
{
    private readonly SafeFileHandle _deviceHandle;
    private IntPtr _winUsbHandle;

    private WinUsbDevice(SafeFileHandle deviceHandle, IntPtr winUsbHandle)
    {
        _deviceHandle = deviceHandle;
        _winUsbHandle = winUsbHandle;
    }

    internal static WinUsbDevice Open(string devicePath)
    {
        var handle = NativeMethods.CreateFile(
            devicePath,
            NativeMethods.GenericRead | NativeMethods.GenericWrite,
            NativeMethods.FileShareRead | NativeMethods.FileShareWrite,
            IntPtr.Zero,
            NativeMethods.OpenExisting,
            NativeMethods.FileAttributeNormal | NativeMethods.FileFlagOverlapped,
            IntPtr.Zero);

        if (handle.IsInvalid)
        {
            throw CreateWin32Exception("CreateFile failed");
        }

        if (!NativeMethods.WinUsb_Initialize(handle, out var winUsbHandle))
        {
            handle.Dispose();
            throw CreateWin32Exception("WinUsb_Initialize failed");
        }

        return new WinUsbDevice(handle, winUsbHandle);
    }

    private byte GetCurrentAlternateSetting()
    {
        if (!NativeMethods.WinUsb_GetCurrentAlternateSetting(_winUsbHandle, out var alternateSetting))
        {
            throw CreateWin32Exception("WinUsb_GetCurrentAlternateSetting failed");
        }

        return alternateSetting;
    }

    internal void SetCurrentAlternateSetting(byte alternateSetting)
    {
        if (!NativeMethods.WinUsb_SetCurrentAlternateSetting(_winUsbHandle, alternateSetting))
        {
            throw CreateWin32Exception("WinUsb_SetCurrentAlternateSetting failed");
        }
    }

    internal IReadOnlyList<AlternateSettingInfo> GetAlternateSettings()
    {
        var result = new List<AlternateSettingInfo>();

        for (byte alternateSetting = 0; alternateSetting < byte.MaxValue; alternateSetting++)
        {
            if (!NativeMethods.WinUsb_QueryInterfaceSettings(_winUsbHandle, alternateSetting, out var descriptor))
            {
                var error = Marshal.GetLastWin32Error();
                if (error == NativeMethods.ErrorNoMoreItems)
                {
                    break;
                }

                throw new InvalidOperationException($"WinUsb_QueryInterfaceSettings failed for alt {alternateSetting}: {error}");
            }

            var pipes = new List<PipeInfo>();
            for (byte pipeIndex = 0; pipeIndex < descriptor.NumEndpoints; pipeIndex++)
            {
                if (!NativeMethods.WinUsb_QueryPipe(_winUsbHandle, alternateSetting, pipeIndex, out var pipeInformation))
                {
                    throw CreateWin32Exception($"WinUsb_QueryPipe failed for alt {alternateSetting}, pipe index {pipeIndex}");
                }

                uint? maximumBytesPerInterval = null;
                if (NativeMethods.WinUsb_QueryPipeEx(_winUsbHandle, alternateSetting, pipeIndex, out var pipeInformationEx))
                {
                    maximumBytesPerInterval = pipeInformationEx.MaximumBytesPerInterval;
                }

                pipes.Add(new PipeInfo(
                    pipeInformation.PipeType,
                    pipeInformation.PipeId,
                    pipeInformation.MaximumPacketSize,
                    pipeInformation.Interval,
                    maximumBytesPerInterval));
            }

            result.Add(new AlternateSettingInfo(descriptor, pipes));
        }

        return result;
    }

    internal IsochReadResult ReadIsochronousPipeAsap(byte pipeId, int packetCount, uint timeoutMs, bool continueStream = false, Action? afterSubmit = null)
    {
        if (packetCount <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(packetCount), "Packet count must be positive.");
        }

        var currentAlternateSetting = GetCurrentAlternateSetting();
        var pipe = GetAlternateSettings()
            .FirstOrDefault(setting => setting.Descriptor.AlternateSetting == currentAlternateSetting)?
            .Pipes
            .FirstOrDefault(candidate => candidate.PipeId == pipeId);

        if (pipe is null)
        {
            throw new InvalidOperationException($"Pipe 0x{pipeId:X2} was not found on alternate setting {currentAlternateSetting}.");
        }

        if (pipe.PipeType != NativeMethods.UsbdPipeType.Isochronous)
        {
            throw new InvalidOperationException($"Pipe 0x{pipeId:X2} is not isochronous.");
        }

        var bytesPerInterval = GetBytesPerInterval(pipe.MaximumPacketSize);
        if (bytesPerInterval <= 0)
        {
            throw new InvalidOperationException($"Pipe 0x{pipeId:X2} is currently zero-bandwidth. Select an active alternate setting first.");
        }

        var requestedLength = checked(bytesPerInterval * packetCount);
        var buffer = new byte[requestedLength];
        var pinnedBuffer = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        var registeredBufferHandle = IntPtr.Zero;
        var descriptorArraySize = Marshal.SizeOf<NativeMethods.UsbdIsoPacketDescriptor>() * packetCount;
        var descriptorsPtr = IntPtr.Zero;
        var overlappedPtr = IntPtr.Zero;

        try
        {
            if (!NativeMethods.WinUsb_RegisterIsochBuffer(_winUsbHandle, pipeId, pinnedBuffer.AddrOfPinnedObject(), (uint)buffer.Length, out registeredBufferHandle))
            {
                throw CreateWin32Exception("WinUsb_RegisterIsochBuffer failed");
            }

            using var completionEvent = NativeMethods.CreateEvent(IntPtr.Zero, false, false, null);
            if (completionEvent.IsInvalid)
            {
                throw CreateWin32Exception("CreateEvent failed");
            }

            descriptorsPtr = Marshal.AllocHGlobal(descriptorArraySize);
            for (var i = 0; i < descriptorArraySize; i++)
            {
                Marshal.WriteByte(descriptorsPtr, i, 0);
            }

            overlappedPtr = Marshal.AllocHGlobal(Marshal.SizeOf<NativeMethods.OverlappedNative>());
            var overlapped = new NativeMethods.OverlappedNative
            {
                EventHandle = completionEvent.DangerousGetHandle()
            };
            Marshal.StructureToPtr(overlapped, overlappedPtr, false);

            if (!NativeMethods.WinUsb_ReadIsochPipeAsapNative(
                registeredBufferHandle,
                0,
                (uint)buffer.Length,
                continueStream,
                (uint)packetCount,
                descriptorsPtr,
                overlappedPtr))
            {
                var error = Marshal.GetLastWin32Error();
                if (error != NativeMethods.ErrorIoPending)
                {
                    throw new Win32Exception(error, $"WinUsb_ReadIsochPipeAsap failed: {error}");
                }
            }

            afterSubmit?.Invoke();

            var waitResult = NativeMethods.WaitForSingleObject(completionEvent, timeoutMs == 0 ? NativeMethods.Infinite : timeoutMs);
            if (waitResult == NativeMethods.WaitTimeout)
            {
                _ = NativeMethods.CancelIoEx(_deviceHandle, overlappedPtr);
                throw new TimeoutException($"Isochronous read timed out after {timeoutMs} ms.");
            }

            if (waitResult != NativeMethods.WaitObject0)
            {
                throw new Win32Exception((int)waitResult, $"WaitForSingleObject failed: {waitResult}");
            }

            uint transferred = 0;
            int? completionError = null;
            if (!NativeMethods.WinUsb_GetOverlappedResultNative(_winUsbHandle, overlappedPtr, out transferred, false))
            {
                completionError = Marshal.GetLastWin32Error();
            }

            var packets = new IsochPacketInfo[packetCount];
            var descriptorSize = Marshal.SizeOf<NativeMethods.UsbdIsoPacketDescriptor>();
            for (var i = 0; i < packetCount; i++)
            {
                var descriptorPtr = IntPtr.Add(descriptorsPtr, i * descriptorSize);
                var descriptor = Marshal.PtrToStructure<NativeMethods.UsbdIsoPacketDescriptor>(descriptorPtr);
                packets[i] = new IsochPacketInfo(i, descriptor.Offset, descriptor.Length, descriptor.Status);
            }

            return new IsochReadResult(bytesPerInterval, requestedLength, (int)transferred, buffer, packets, completionError);
        }
        finally
        {
            if (overlappedPtr != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(overlappedPtr);
            }

            if (descriptorsPtr != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(descriptorsPtr);
            }

            if (registeredBufferHandle != IntPtr.Zero)
            {
                _ = NativeMethods.WinUsb_UnregisterIsochBuffer(registeredBufferHandle);
            }

            if (pinnedBuffer.IsAllocated)
            {
                pinnedBuffer.Free();
            }
        }
    }

    internal byte[] ControlTransferIn(byte requestType, byte request, ushort value, ushort index, int length)
    {
        var buffer = new byte[length];
        var setupPacket = CreateSetupPacket(requestType, request, value, index, (ushort)length);

        if (!NativeMethods.WinUsb_ControlTransfer(_winUsbHandle, setupPacket, buffer, (uint)buffer.Length, out var transferred, IntPtr.Zero))
        {
            throw CreateWin32Exception("WinUsb_ControlTransfer failed");
        }

        return buffer.AsSpan(0, (int)transferred).ToArray();
    }

    internal int ControlTransferOut(byte requestType, byte request, ushort value, ushort index, byte[] payload)
    {
        var setupPacket = CreateSetupPacket(requestType, request, value, index, (ushort)payload.Length);

        if (!NativeMethods.WinUsb_ControlTransfer(_winUsbHandle, setupPacket, payload, (uint)payload.Length, out var transferred, IntPtr.Zero))
        {
            throw CreateWin32Exception("WinUsb_ControlTransfer failed");
        }

        return (int)transferred;
    }

    public void Dispose()
    {
        if (_winUsbHandle != IntPtr.Zero)
        {
            _ = NativeMethods.WinUsb_Free(_winUsbHandle);
            _winUsbHandle = IntPtr.Zero;
        }

        _deviceHandle.Dispose();
    }

    private static NativeMethods.WinUsbSetupPacket CreateSetupPacket(byte requestType, byte request, ushort value, ushort index, ushort length)
    {
        return new NativeMethods.WinUsbSetupPacket
        {
            RequestType = requestType,
            Request = request,
            Value = value,
            Index = index,
            Length = length
        };
    }

    private static Win32Exception CreateWin32Exception(string prefix)
    {
        var error = Marshal.GetLastWin32Error();
        return new Win32Exception(error, $"{prefix}: {error}");
    }

    private static int GetBytesPerInterval(ushort maximumPacketSize)
    {
        var payloadBytes = maximumPacketSize & 0x07FF;
        var additionalTransactions = (maximumPacketSize >> 11) & 0x0003;
        return payloadBytes * (additionalTransactions + 1);
    }
}

internal sealed record AlternateSettingInfo(
    NativeMethods.UsbInterfaceDescriptor Descriptor,
    IReadOnlyList<PipeInfo> Pipes);

internal sealed record PipeInfo(
    NativeMethods.UsbdPipeType PipeType,
    byte PipeId,
    ushort MaximumPacketSize,
    byte Interval,
    uint? MaximumBytesPerInterval);

internal sealed record IsochReadResult(
    int BytesPerInterval,
    int RequestedLength,
    int CompletedBytes,
    byte[] Buffer,
    IReadOnlyList<IsochPacketInfo> Packets,
    int? CompletionError);

internal sealed record IsochPacketInfo(
    int Index,
    uint Offset,
    uint Length,
    uint Status);
