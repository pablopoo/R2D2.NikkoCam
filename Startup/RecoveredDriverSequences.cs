namespace R2D2.NikkoCam;

// Minimal subset of the recovered XP startup transfers that the fixed Nikko app
// still uses. The old generic sequence catalog was removed because this project
// only boots one NTSC 352x240 path and one preview-start handshake.
internal static class RecoveredDriverSequences
{
    internal sealed record Step(byte Request, ushort Value, ushort Index, int DelayAfterMs = 0);

    private static readonly (ushort Value, ushort Index)[] Xp19a0fTable =
    [
        (0x00DF, 0x001F), (0x00FF, 0x0008), (0x00FF, 0x0000), (0x00D5, 0x004F), (0x00DA, 0x0023),
        (0x00DB, 0x0008), (0x00E2, 0x0000), (0x00E3, 0x0010), (0x00E5, 0x0000), (0x00E8, 0x0000),
        (0x00EB, 0x0064), (0x00EE, 0x00C2), (0x003F, 0x0001), (0x0000, 0x0000), (0x0001, 0x0007),
        (0x0002, 0x005F), (0x0003, 0x0000), (0x0005, 0x0064), (0x0007, 0x0001), (0x0008, 0x0082),
        (0x0009, 0x0036), (0x000A, 0x0050), (0x000C, 0x006A), (0x0011, 0x00C9), (0x0012, 0x0007),
        (0x0013, 0x003B), (0x0014, 0x0047), (0x0015, 0x006F), (0x0017, 0x00CD), (0x0018, 0x001E),
        (0x0019, 0x008B), (0x001A, 0x00A2), (0x001B, 0x00E9), (0x001C, 0x001C), (0x001D, 0x00CC),
        (0x001E, 0x00CC), (0x001F, 0x00CD), (0x0020, 0x003C), (0x0021, 0x003C), (0x002D, 0x0048),
        (0x002E, 0x0088), (0x0030, 0x0022), (0x0031, 0x0061), (0x0032, 0x0074), (0x0033, 0x001C),
        (0x0034, 0x0074), (0x0035, 0x001C), (0x0036, 0x007A), (0x0037, 0x0026), (0x0038, 0x0040),
        (0x0039, 0x000A), (0x0042, 0x0055), (0x0051, 0x0011), (0x0055, 0x0001), (0x0057, 0x0002),
        (0x0058, 0x0035), (0x0059, 0x00A0), (0x0080, 0x0015), (0x0082, 0x0042), (0x00C1, 0x00D0),
        (0x00C3, 0x0088), (0x003F, 0x0000),
    ];

    private static readonly Step[] Xp19a0fCommonNoRead =
    [
        ..BuildRegisterTable(Xp19a0fTable),
        BuildVendorWrite(0x05, 0x0018, 0x0000),
        BuildVendorWrite(0x03, 0x0102, 0x0000, 10),
        BuildVendorWrite(0x03, 0x0102, 0x0001),
        BuildVendorWrite(0x03, 0x0104, 0x0001),
    ];

    private static readonly Step[] XpProcAmpDefaultsNtsc =
    [
        BuildVendorWrite(0x07, 0x0009, 0x0036),
        BuildVendorWrite(0x07, 0x000B, 0x0004),
        BuildVendorWrite(0x07, 0x0008, 0x0077),
        BuildVendorWrite(0x07, 0x000A, 0x0060),
    ];

    private static readonly Step[] XpObservedRouteInput1 =
    [
        BuildVendorWrite(0x07, 0x003F, 0x0001),
        BuildVendorWrite(0x07, 0x00C0, 0x0020),
        BuildVendorWrite(0x07, 0x003F, 0x0000),
        BuildVendorWrite(0x07, 0x00FF, 0x0008),
        BuildVendorWrite(0x07, 0x00FF, 0x0000),
        BuildVendorWrite(0x07, 0x00C0, 0x0020),
        BuildVendorWrite(0x07, 0x00C1, 0x00D0),
        BuildVendorWrite(0x07, 0x00C3, 0x0088),
        BuildVendorWrite(0x07, 0x00DA, 0x0023),
        BuildVendorWrite(0x07, 0x00D1, 0x00C0),
        BuildVendorWrite(0x07, 0x00D2, 0x00D8),
        BuildVendorWrite(0x07, 0x00D6, 0x0006),
        BuildVendorWrite(0x07, 0x00DF, 0x001F),
        BuildVendorWrite(0x07, 0x003F, 0x0000),
    ];

    private static readonly Step[] XpObservedEbHandshake =
    [
        BuildVendorWrite(0x07, 0x00EB, 0x0064),
        BuildVendorWrite(0x07, 0x00EB, 0x0064),
    ];

    private static readonly Step[] XpObservedPostEbNtsc =
    [
        BuildVendorWrite(0x07, 0x0009, 0x0047),
        BuildVendorWrite(0x07, 0x0080, 0x0029),
        BuildVendorWrite(0x07, 0x0008, 0x006B),
        BuildVendorWrite(0x07, 0x000A, 0x0060),
    ];

    internal static IReadOnlyList<Step> FixedStartupSequence { get; } = Combine(
        Xp19a0fCommonNoRead,
        [BuildVendorWrite(0x04, 0x0002, 0x0000)],
        [BuildVendorWrite(0x03, 0x0104, 0x0001)],
        XpProcAmpDefaultsNtsc,
        [BuildVendorWrite(0x04, 0x0002, 0x0001)],
        XpObservedRouteInput1,
        XpObservedEbHandshake,
        XpObservedPostEbNtsc);

    internal static IReadOnlyList<Step> PreviewStartSequence { get; } =
    [
        BuildVendorWrite(0x03, 0x0102, 0x0000),
        BuildVendorWrite(0x03, 0x0103, 0x0000),
    ];

    private static Step BuildVendorWrite(byte request, ushort value, ushort index, int delayAfterMs = 0)
        => new(request, value, index, delayAfterMs);

    private static IReadOnlyList<Step> BuildRegisterTable((ushort Value, ushort Index)[] entries)
        => entries.Select(static entry => BuildVendorWrite(0x07, entry.Value, entry.Index)).ToArray();

    private static IReadOnlyList<Step> Combine(params IReadOnlyList<Step>[] parts)
        => parts.SelectMany(static part => part).ToArray();
}
