using System.Runtime.InteropServices;

namespace R2D2.NikkoCam;

internal static class UsbDeviceDiscovery
{
    internal static readonly Guid GenericUsbDeviceInterfaceGuid = new("A5DCBF10-6530-11D2-901F-00C04FB951ED");

    internal static IReadOnlyList<string> FindDevicePaths(ushort vendorId, ushort productId, Guid interfaceGuid)
    {
        return EnumerateDevicePaths(interfaceGuid)
            .Where(path => PathMatchesVidPid(path, vendorId, productId))
            .ToArray();
    }

    private static IReadOnlyList<string> EnumerateDevicePaths(Guid interfaceGuid)
    {
        var result = new List<string>();
        var deviceInfoSet = NativeMethods.SetupDiGetClassDevs(
            ref interfaceGuid,
            null,
            IntPtr.Zero,
            NativeMethods.DigcfPresent | NativeMethods.DigcfDeviceInterface);

        if (deviceInfoSet == NativeMethods.InvalidHandleValue)
        {
            throw new InvalidOperationException($"SetupDiGetClassDevs failed: {Marshal.GetLastWin32Error()}");
        }

        try
        {
            for (uint memberIndex = 0; ; memberIndex++)
            {
                var interfaceData = NativeMethods.CreateDeviceInterfaceData();
                if (!NativeMethods.SetupDiEnumDeviceInterfaces(deviceInfoSet, IntPtr.Zero, ref interfaceGuid, memberIndex, ref interfaceData))
                {
                    var error = Marshal.GetLastWin32Error();
                    if (error == NativeMethods.ErrorNoMoreItems)
                    {
                        break;
                    }

                    throw new InvalidOperationException($"SetupDiEnumDeviceInterfaces failed: {error}");
                }

                _ = NativeMethods.SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref interfaceData, IntPtr.Zero, 0, out var requiredSize, IntPtr.Zero);
                var detailBuffer = Marshal.AllocHGlobal((int)requiredSize);

                try
                {
                    Marshal.WriteInt32(detailBuffer, IntPtr.Size == 8 ? 8 : 6);
                    if (!NativeMethods.SetupDiGetDeviceInterfaceDetail(deviceInfoSet, ref interfaceData, detailBuffer, requiredSize, out _, IntPtr.Zero))
                    {
                        throw new InvalidOperationException($"SetupDiGetDeviceInterfaceDetail failed: {Marshal.GetLastWin32Error()}");
                    }

                    var pathPtr = IntPtr.Add(detailBuffer, 4);
                    var devicePath = Marshal.PtrToStringUni(pathPtr);
                    if (!string.IsNullOrWhiteSpace(devicePath))
                    {
                        result.Add(devicePath);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(detailBuffer);
                }
            }
        }
        finally
        {
            _ = NativeMethods.SetupDiDestroyDeviceInfoList(deviceInfoSet);
        }

        return result;
    }

    private static bool PathMatchesVidPid(string devicePath, ushort vendorId, ushort productId)
    {
        var normalized = devicePath.ToUpperInvariant();
        return normalized.Contains($"VID_{vendorId:X4}", StringComparison.Ordinal)
            && normalized.Contains($"PID_{productId:X4}", StringComparison.Ordinal);
    }
}
