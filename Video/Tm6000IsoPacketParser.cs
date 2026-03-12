namespace R2D2.NikkoCam;

// Converts raw isoch packet payload into TM6000 record slices. The receiver sends
// a stream of fixed-size 184-byte records that can span USB packets, so we keep a
// small carry buffer and resynchronize on valid record headers.
internal sealed class Tm6000IsoPacketParser
{
    private readonly List<byte> _carry = new();
    internal void Reset()
    {
        _carry.Clear();
    }

    internal IReadOnlyList<BulkCaptureAnalyzer.RecordSlice> Push(IsochReadResult result)
    {
        var emitted = new List<BulkCaptureAnalyzer.RecordSlice>();

        foreach (var packet in result.Packets.OrderBy(static packet => packet.Index))
        {
            if (packet.Status != 0 || packet.Length == 0)
            {
                HandlePacketLoss();
                continue;
            }

            var offset = checked((int)packet.Offset);
            var length = checked((int)packet.Length);
            if (offset < 0 || length <= 0 || offset + length > result.Buffer.Length)
            {
                HandlePacketLoss();
                continue;
            }

            _carry.AddRange(result.Buffer.AsSpan(offset, length).ToArray());
            ParseCarry(emitted);
        }

        return emitted;
    }

    // Consume as many complete TM6000 records as possible from the carry buffer.
    private void ParseCarry(List<BulkCaptureAnalyzer.RecordSlice> emitted)
    {
        while (true)
        {
            if (_carry.Count < 4)
            {
                return;
            }

            var headerOffset = FindNextHeaderOffset();
            if (headerOffset < 0)
            {
                KeepTrailingBytes(3);
                return;
            }

            if (headerOffset > 0)
            {
                _carry.RemoveRange(0, headerOffset);
            }

            if (_carry.Count < 4)
            {
                return;
            }

            var markerValue = ReadUInt32LittleEndian(_carry, 0);
            var header = BulkCaptureAnalyzer.DecodeHeader(markerValue);
            var recordLength = 4 + header.PayloadBytes;
            if (!LooksLikeRecordHeader(header, recordLength))
            {
                _carry.RemoveAt(0);
                continue;
            }

            if (_carry.Count < recordLength)
            {
                return;
            }

            var bytes = _carry.Take(recordLength).ToArray();
            emitted.Add(new BulkCaptureAnalyzer.RecordSlice(markerValue, bytes));

            _carry.RemoveRange(0, recordLength);
        }
    }

    // Scan for the next byte position that plausibly starts a TM6000 record.
    private int FindNextHeaderOffset()
    {
        for (var offset = 0; offset + 4 <= _carry.Count; offset++)
        {
            var markerValue = ReadUInt32LittleEndian(_carry, offset);
            var header = BulkCaptureAnalyzer.DecodeHeader(markerValue);
            var recordLength = 4 + header.PayloadBytes;
            if (!LooksLikeRecordHeader(header, recordLength))
            {
                continue;
            }

            if (_carry.Count >= offset + recordLength + 4)
            {
                var nextMarkerValue = ReadUInt32LittleEndian(_carry, offset + recordLength);
                var nextHeader = BulkCaptureAnalyzer.DecodeHeader(nextMarkerValue);
                var nextRecordLength = 4 + nextHeader.PayloadBytes;
                if (!LooksLikeRecordHeader(nextHeader, nextRecordLength))
                {
                    continue;
                }
            }

            return offset;
        }

        return -1;
    }

    private static bool LooksLikeRecordHeader(BulkCaptureAnalyzer.DecodedHeader header, int recordLength)
    {
        if (header.Command < 1 || header.Command > 4)
        {
            return false;
        }

        if (header.PayloadBytes <= 0 || header.PayloadBytes > BulkCaptureAnalyzer.Tm6000UrbPayloadBytes)
        {
            return false;
        }

        if (recordLength != 184)
        {
            return false;
        }

        return true;
    }

    private void KeepTrailingBytes(int count)
    {
        if (_carry.Count <= count)
        {
            return;
        }

        _carry.RemoveRange(0, _carry.Count - count);
    }

    private void HandlePacketLoss()
    {
        _carry.Clear();
    }

    private static uint ReadUInt32LittleEndian(List<byte> bytes, int offset)
    {
        return (uint)(
            bytes[offset] |
            (bytes[offset + 1] << 8) |
            (bytes[offset + 2] << 16) |
            (bytes[offset + 3] << 24));
    }
}
