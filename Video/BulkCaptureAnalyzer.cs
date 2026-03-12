namespace R2D2.NikkoCam;

// TM6000 record helpers used by the live Nikko preview path. The original file
// also contained offline capture-analysis utilities, but the app only needs:
// - record slices emitted by the packet parser
// - decoded field/line/block headers
// - a small command histogram for preview bookkeeping
internal static class BulkCaptureAnalyzer
{
    internal const int Tm6000UrbPayloadBytes = 180;

    internal sealed record RecordSlice(uint MarkerValue, byte[] Bytes);

    internal sealed record DecodedHeader(
        int PayloadBytes,
        int Block,
        int Field,
        int Line,
        int Command,
        bool LooksLikeVideo);

    internal static DecodedHeader DecodeHeader(uint markerValue)
    {
        var payloadBytes = (int)((markerValue & 0x7E) << 1);
        if (payloadBytes > 0)
        {
            payloadBytes -= 4;
        }

        var block = (int)((markerValue >> 7) & 0x0F);
        var field = (int)((markerValue >> 11) & 0x01);
        var line = (int)((markerValue >> 12) & 0x01FF);
        var command = (int)((markerValue >> 21) & 0x07);
        var looksLikeVideo = command == 1 && payloadBytes >= 0 && payloadBytes <= Tm6000UrbPayloadBytes;
        return new DecodedHeader(payloadBytes, block, field, line, command, looksLikeVideo);
    }

    internal static IReadOnlyDictionary<int, int> BuildCommandHistogram(IReadOnlyList<RecordSlice> records)
    {
        var histogram = new Dictionary<int, int>();
        foreach (var record in records)
        {
            var header = DecodeHeader(record.MarkerValue);
            histogram[header.Command] = histogram.GetValueOrDefault(header.Command) + 1;
        }

        return histogram;
    }
}
