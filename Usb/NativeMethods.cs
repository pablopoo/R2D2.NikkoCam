using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;

namespace R2D2.NikkoCam;

internal static class NativeMethods
{
    public const uint GenericRead = 0x80000000;
    public const uint GenericWrite = 0x40000000;
    public const uint FileShareRead = 0x00000001;
    public const uint FileShareWrite = 0x00000002;
    public const uint OpenExisting = 3;
    public const uint FileAttributeNormal = 0x00000080;
    public const uint FileFlagOverlapped = 0x40000000;

    public const uint DigcfPresent = 0x00000002;
    public const uint DigcfDeviceInterface = 0x00000010;

    public const int ErrorNoMoreItems = 259;
    public const int ErrorIoPending = 997;
    public const uint WaitObject0 = 0x00000000;
    public const uint WaitTimeout = 0x00000102;
    public const uint Infinite = 0xFFFFFFFF;

    public static readonly IntPtr InvalidHandleValue = new(-1);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern SafeFileHandle CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern SafeWaitHandle CreateEvent(
        IntPtr lpEventAttributes,
        [MarshalAs(UnmanagedType.Bool)] bool bManualReset,
        [MarshalAs(UnmanagedType.Bool)] bool bInitialState,
        string? lpName);

    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint WaitForSingleObject(
        SafeWaitHandle hHandle,
        uint dwMilliseconds);

    [DllImport("kernel32.dll", SetLastError = true, EntryPoint = "CancelIoEx")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool CancelIoEx(
        SafeFileHandle hFile,
        IntPtr lpOverlapped);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr SetupDiGetClassDevs(
        ref Guid classGuid,
        string? enumerator,
        IntPtr hwndParent,
        uint flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetupDiEnumDeviceInterfaces(
        IntPtr deviceInfoSet,
        IntPtr deviceInfoData,
        ref Guid interfaceClassGuid,
        uint memberIndex,
        ref SpDeviceInterfaceData deviceInterfaceData);

    [DllImport("setupapi.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetupDiGetDeviceInterfaceDetail(
        IntPtr deviceInfoSet,
        ref SpDeviceInterfaceData deviceInterfaceData,
        IntPtr deviceInterfaceDetailData,
        uint deviceInterfaceDetailDataSize,
        out uint requiredSize,
        IntPtr deviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetupDiDestroyDeviceInfoList(IntPtr deviceInfoSet);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_Initialize(SafeFileHandle deviceHandle, out IntPtr interfaceHandle);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_Free(IntPtr interfaceHandle);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_QueryInterfaceSettings(
        IntPtr interfaceHandle,
        byte alternateSettingNumber,
        out UsbInterfaceDescriptor usbAltInterfaceDescriptor);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_QueryPipe(
        IntPtr interfaceHandle,
        byte alternateInterfaceNumber,
        byte pipeIndex,
        out WinUsbPipeInformation pipeInformation);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_QueryPipeEx(
        IntPtr interfaceHandle,
        byte alternateInterfaceNumber,
        byte pipeIndex,
        out WinUsbPipeInformationEx pipeInformationEx);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_GetCurrentAlternateSetting(
        IntPtr interfaceHandle,
        out byte settingValue);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_SetCurrentAlternateSetting(
        IntPtr interfaceHandle,
        byte settingValue);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_ControlTransfer(
        IntPtr interfaceHandle,
        WinUsbSetupPacket setupPacket,
        [Out] byte[] buffer,
        uint bufferLength,
        out uint lengthTransferred,
        IntPtr overlapped);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_RegisterIsochBuffer(
        IntPtr interfaceHandle,
        byte pipeId,
        IntPtr buffer,
        uint bufferLength,
        out IntPtr bufferHandle);

    [DllImport("winusb.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_UnregisterIsochBuffer(IntPtr bufferHandle);

    [DllImport("winusb.dll", SetLastError = true, EntryPoint = "WinUsb_ReadIsochPipeAsap")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_ReadIsochPipeAsapNative(
        IntPtr bufferHandle,
        uint offset,
        uint length,
        [MarshalAs(UnmanagedType.Bool)] bool continueStream,
        uint numberOfPackets,
        IntPtr isoPacketDescriptors,
        IntPtr overlapped);

    [DllImport("winusb.dll", SetLastError = true, EntryPoint = "WinUsb_GetOverlappedResult")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool WinUsb_GetOverlappedResultNative(
        IntPtr interfaceHandle,
        IntPtr overlapped,
        out uint numberOfBytesTransferred,
        [MarshalAs(UnmanagedType.Bool)] bool wait);

    public static SpDeviceInterfaceData CreateDeviceInterfaceData()
    {
        return new SpDeviceInterfaceData
        {
            Size = Marshal.SizeOf<SpDeviceInterfaceData>()
        };
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SpDeviceInterfaceData
    {
        public int Size;
        public Guid InterfaceClassGuid;
        public int Flags;
        public IntPtr Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct UsbInterfaceDescriptor
    {
        public byte Length;
        public byte DescriptorType;
        public byte InterfaceNumber;
        public byte AlternateSetting;
        public byte NumEndpoints;
        public byte InterfaceClass;
        public byte InterfaceSubClass;
        public byte InterfaceProtocol;
        public byte Interface;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WinUsbPipeInformation
    {
        public UsbdPipeType PipeType;
        public byte PipeId;
        public ushort MaximumPacketSize;
        public byte Interval;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WinUsbPipeInformationEx
    {
        public UsbdPipeType PipeType;
        public byte PipeId;
        public ushort MaximumPacketSize;
        public byte Interval;
        public uint MaximumBytesPerInterval;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WinUsbSetupPacket
    {
        public byte RequestType;
        public byte Request;
        public ushort Value;
        public ushort Index;
        public ushort Length;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct UsbdIsoPacketDescriptor
    {
        public uint Offset;
        public uint Length;
        public uint Status;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct OverlappedNative
    {
        public IntPtr Internal;
        public IntPtr InternalHigh;
        public uint Offset;
        public uint OffsetHigh;
        public IntPtr EventHandle;
    }

    public enum UsbdPipeType
    {
        Control = 0,
        Isochronous = 1,
        Bulk = 2,
        Interrupt = 3
    }
}
