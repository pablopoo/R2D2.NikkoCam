namespace R2D2.NikkoCam;

// Maintains the current in-flight video raster. The TM6000 transport sends small
// blocks tagged with field/line/block coordinates; this class merges them into one
// logical frame before handing them to the bitmap builder.
internal sealed class RollingPreviewAssembler
{
    private const int HeaderBytes = 4;
    private const int ExpectedFieldLineMaximum = 240;

    private readonly Tm6000IsoPacketParser _parser = new();
    private readonly Dictionary<RecordKey, BulkCaptureAnalyzer.RecordSlice> _buildingRecords = new();
    private readonly PreviewPersistentPlaneState _persistentPlane = new();

    private int _lastObservedField = -1;
    private int _preferredField = -1;

    internal void Reset()
    {
        _parser.Reset();
        _buildingRecords.Clear();
        _persistentPlane.Reset();
        _lastObservedField = -1;
        _preferredField = -1;
    }

    internal void SetPreferredField(int value)
    {
        _preferredField = value is 0 or 1 ? value : -1;
    }

    internal PreviewFrameResult? PushAndBuild(
        IsochReadResult result,
        int targetDisplayWidth,
        int targetDisplayHeight)
    {
        if (result.Buffer.Length == 0 || result.BytesPerInterval <= 0)
        {
            return null;
        }

        var records = _parser.Push(result);
        if (records.Count == 0)
        {
            return null;
        }

        foreach (var record in records)
        {
            var header = BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue);
            if (!header.LooksLikeVideo || header.Line <= 0 || header.Line > ExpectedFieldLineMaximum + 16)
            {
                continue;
            }

            // Field 1 after field 0 is the simplest reliable frame boundary the
            // receiver gives us in this fixed mode, so start a fresh block map there
            // while keeping the persistent raster for missing-block carry-over.
            if (_lastObservedField != -1 && _lastObservedField != header.Field && header.Field == 1)
            {
                _buildingRecords.Clear();
            }

            _lastObservedField = header.Field;

            var payloadBytes = Math.Min(header.PayloadBytes, record.Bytes.Length - HeaderBytes);
            if (payloadBytes <= 0)
            {
                continue;
            }

            var normalized = new byte[HeaderBytes + payloadBytes];
            Array.Copy(record.Bytes, 0, normalized, 0, HeaderBytes + payloadBytes);
            _buildingRecords[new RecordKey(header.Field, header.Line, header.Block)] =
                new BulkCaptureAnalyzer.RecordSlice(record.MarkerValue, normalized);
        }

        if (_buildingRecords.Count == 0)
        {
            return null;
        }

        var aggregated = _buildingRecords.Values
            .OrderBy(static record => BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue).Field)
            .ThenBy(static record => BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue).Line)
            .ThenBy(static record => BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue).Block)
            .ToArray();

        return PreviewFrameBuilder.TryBuildFixedNikko(
            aggregated,
            targetDisplayWidth,
            targetDisplayHeight,
            _persistentPlane,
            _preferredField,
            HeaderBytes);
    }

    private readonly record struct RecordKey(int Field, int Line, int Block);
}
