using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace R2D2.NikkoCam;

// Bitmap payload returned to the UI for one live preview update.
internal sealed record PreviewFrameResult(
    Bitmap Bitmap,
    int Field);

// Persistent raster storage. Missing blocks simply keep their previous content
// instead of turning black, which matches the more analog-looking behavior we want.
internal sealed class PreviewPersistentPlaneState
{
    private int _width;
    private int _height;
    private byte[] _packedPixels = [];
    private bool[] _writtenPixels = [];

    internal (byte[] PackedPixels, bool[] WrittenPixels) Acquire(int width, int height)
    {
        var packedLength = Math.Max(width * height * 2, 1);
        var writtenLength = Math.Max(width * height, 1);
        if (_width != width ||
            _height != height ||
            _packedPixels.Length != packedLength ||
            _writtenPixels.Length != writtenLength)
        {
            _width = width;
            _height = height;
            _packedPixels = new byte[packedLength];
            _writtenPixels = new bool[writtenLength];
        }

        return (_packedPixels, _writtenPixels);
    }

    internal void Reset()
    {
        _width = 0;
        _height = 0;
        _packedPixels = [];
        _writtenPixels = [];
    }
}

internal static class PreviewFrameBuilder
{
    private const int ExpectedBlocksPerLine = 8;
    private const int PayloadBytesPerRecord = BulkCaptureAnalyzer.Tm6000UrbPayloadBytes;

    // Fixed Nikko preview pipeline:
    // records -> interlaced UYVY raster -> XP-style 352 transform -> preview deinterlace -> bitmap.
    internal static PreviewFrameResult? TryBuildFixedNikko(
        IReadOnlyList<BulkCaptureAnalyzer.RecordSlice> records,
        int targetDisplayWidth,
        int targetDisplayHeight,
        PreviewPersistentPlaneState persistentPlaneState,
        int preferredField = -1,
        int headerBytes = 4)
    {
        if (records.Count == 0)
        {
            return null;
        }

        var commandHistogram = BulkCaptureAnalyzer.BuildCommandHistogram(records);
        if (!commandHistogram.TryGetValue(1, out var videoRecordCount) || videoRecordCount == 0)
        {
            return null;
        }

        var selectedField = SelectBestField(records, preferredField);
        var plane = BuildInterlacedPlane(
            records,
            headerBytes,
            targetDisplayHeight,
            persistentPlaneState);
        if (plane is null)
        {
            return null;
        }

        var displayPlane = ApplyOriginal352DeinterlaceTransform(
            plane,
            targetDisplayWidth,
            selectedField);
        var lumaRange = ComputeLumaRange(
            displayPlane.PackedPixels,
            displayPlane.WrittenPixels,
            displayPlane.Width,
            displayPlane.Height);
        var bitmap = BuildColorBitmap(
            displayPlane.Width,
            displayPlane.Height,
            displayPlane.PackedPixels,
            displayPlane.WrittenPixels,
            lumaRange);

        return new PreviewFrameResult(bitmap, selectedField);
    }

    // Place each TM6000 block at its exact field/line/block position in a full
    // interlaced raster. The persistent plane keeps previous pixels for blocks that
    // did not arrive in this update.
    private static Plane? BuildInterlacedPlane(
        IReadOnlyList<BulkCaptureAnalyzer.RecordSlice> records,
        int headerBytes,
        int targetDisplayHeight,
        PreviewPersistentPlaneState persistentPlaneState)
    {
        var videoRecords = records
            .Select(static record => (Record: record, Header: BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue)))
            .Where(static item => item.Header.LooksLikeVideo)
            .ToArray();
        if (videoRecords.Length == 0)
        {
            return null;
        }

        var blocksPerLine = Math.Max(videoRecords.Max(static item => item.Header.Block) + 1, ExpectedBlocksPerLine);
        var width = (blocksPerLine * PayloadBytesPerRecord) / 2;
        var height = Math.Max(targetDisplayHeight, 1);
        var (packedPixels, writtenPixels) = persistentPlaneState.Acquire(width, height);

        foreach (var item in videoRecords)
        {
            if (item.Header.Line <= 0)
            {
                continue;
            }

            var row = (item.Header.Line << 1) - item.Header.Field - 1;
            if (row < 0 || row >= height)
            {
                continue;
            }

            var payloadLength = Math.Min(PayloadBytesPerRecord, item.Record.Bytes.Length - headerBytes);
            if (payloadLength <= 0)
            {
                continue;
            }

            var payload = item.Record.Bytes.AsSpan(headerBytes, payloadLength);
            CopyPackedRow(payload, packedPixels, writtenPixels, width, row, item.Header.Block * (PayloadBytesPerRecord / 2));
        }

        return new Plane(width, height, packedPixels, writtenPixels);
    }

    // The preview shown in this app is deliberately friendlier than the raw field
    // stream: preserve one field parity and synthesize the opposite rows by copying
    // or blending nearby rows after the original 352 transform.
    private static Plane ApplyOriginal352DeinterlaceTransform(Plane plane, int targetDisplayWidth, int selectedField)
    {
        var transformed = ApplyOriginal352Transform(plane, targetDisplayWidth);
        if (transformed.Height <= 1)
        {
            return transformed;
        }

        var preferredParity = selectedField switch
        {
            1 => 0,
            0 => 1,
            _ => SelectDominantParity(transformed),
        };

        var packedPixels = new byte[transformed.PackedPixels.Length];
        var writtenPixels = new bool[transformed.WrittenPixels.Length];

        for (var row = 0; row < transformed.Height; row++)
        {
            if ((row & 1) == preferredParity)
            {
                CopyRow(transformed, row, packedPixels, writtenPixels, row);
                continue;
            }

            var previous = FindNearestParityRow(transformed.Height, row - 1, -1, preferredParity);
            var next = FindNearestParityRow(transformed.Height, row + 1, +1, preferredParity);
            if (previous >= 0 && next >= 0)
            {
                BlendRows(transformed, previous, next, packedPixels, writtenPixels, row);
                continue;
            }

            var sourceRow = previous >= 0 ? previous : next;
            if (sourceRow >= 0)
            {
                CopyRow(transformed, sourceRow, packedPixels, writtenPixels, row);
            }
        }

        return new Plane(transformed.Width, transformed.Height, packedPixels, writtenPixels);
    }

    // Recreate the original app's 720->352 packed-UYVY reduction, but keep the full
    // assembled height because the raster is already at 240 lines in this app.
    private static Plane ApplyOriginal352Transform(Plane plane, int targetDisplayWidth)
    {
        var downsampledWidth = Math.Max(plane.Width / 2, 1);
        var packedPixels = new byte[Math.Max(downsampledWidth * plane.Height * 2, 1)];
        var writtenPixels = new bool[Math.Max(downsampledWidth * plane.Height, 1)];

        for (var row = 0; row < plane.Height; row++)
        {
            var sourceRowOffset = row * plane.Width * 2;
            var targetRowOffset = row * downsampledWidth * 2;
            var sourceWrittenRowOffset = row * plane.Width;
            var targetWrittenRowOffset = row * downsampledWidth;

            for (var outputPixel = 0; outputPixel + 1 < downsampledWidth; outputPixel += 2)
            {
                var sourcePixel = outputPixel * 2;
                var sourceOffset = sourceRowOffset + sourcePixel * 2;
                var targetOffset = targetRowOffset + outputPixel * 2;

                if (sourceOffset + 7 >= plane.PackedPixels.Length || targetOffset + 3 >= packedPixels.Length)
                {
                    break;
                }

                packedPixels[targetOffset + 0] = (byte)((plane.PackedPixels[sourceOffset + 0] + plane.PackedPixels[sourceOffset + 4]) >> 1);
                packedPixels[targetOffset + 1] = (byte)((plane.PackedPixels[sourceOffset + 1] + plane.PackedPixels[sourceOffset + 3]) >> 1);
                packedPixels[targetOffset + 2] = (byte)((plane.PackedPixels[sourceOffset + 2] + plane.PackedPixels[sourceOffset + 6]) >> 1);
                packedPixels[targetOffset + 3] = (byte)((plane.PackedPixels[sourceOffset + 5] + plane.PackedPixels[sourceOffset + 7]) >> 1);

                var targetPixelIndex = targetWrittenRowOffset + outputPixel;
                var sourcePixelIndex = sourceWrittenRowOffset + sourcePixel;
                if (sourcePixelIndex + 3 >= plane.WrittenPixels.Length || targetPixelIndex + 1 >= writtenPixels.Length)
                {
                    break;
                }

                writtenPixels[targetPixelIndex] =
                    plane.WrittenPixels[sourcePixelIndex] ||
                    plane.WrittenPixels[sourcePixelIndex + 1];
                writtenPixels[targetPixelIndex + 1] =
                    plane.WrittenPixels[sourcePixelIndex + 2] ||
                    plane.WrittenPixels[sourcePixelIndex + 3];
            }
        }

        var result = new Plane(downsampledWidth, plane.Height, packedPixels, writtenPixels);
        return targetDisplayWidth > 0 && targetDisplayWidth < result.Width
            ? CropPlaneWidth(result, targetDisplayWidth)
            : result;
    }

    // Prefer the field with better current coverage, but keep some hysteresis so the
    // displayed parity does not flicker when both fields are similarly complete.
    private static int SelectBestField(IReadOnlyList<BulkCaptureAnalyzer.RecordSlice> records, int preferredField)
    {
        var videoRecords = records
            .Select(static record => (Record: record, Header: BulkCaptureAnalyzer.DecodeHeader(record.MarkerValue)))
            .Where(static item => item.Header.LooksLikeVideo)
            .ToArray();

        var stats = new[] { 0, 1 }
            .Select(field =>
            {
                var fieldRecords = videoRecords.Where(item => item.Header.Field == field).ToArray();
                return new FieldStats(
                    field,
                    GetStrongLines(fieldRecords).Count,
                    fieldRecords.Length,
                    fieldRecords.Select(static item => item.Header.Block).Distinct().Count());
            })
            .OrderByDescending(static item => item.ActiveLineCount)
            .ThenByDescending(static item => item.RecordCount)
            .ThenByDescending(static item => item.DistinctBlockCount)
            .ThenBy(static item => item.Field)
            .ToArray();
        if (stats.Length == 0)
        {
            return 0;
        }

        var best = stats[0];
        if (preferredField is 0 or 1)
        {
            var preferred = stats.FirstOrDefault(item => item.Field == preferredField);
            if (preferred is not null &&
                preferred.ActiveLineCount > 0 &&
                preferred.RecordCount > 0 &&
                best.Field != preferred.Field)
            {
                var activeDelta = best.ActiveLineCount - preferred.ActiveLineCount;
                var recordDelta = best.RecordCount - preferred.RecordCount;
                var blockDelta = best.DistinctBlockCount - preferred.DistinctBlockCount;
                var clearlyBetter =
                    activeDelta > Math.Max(24, preferred.ActiveLineCount / 8) ||
                    recordDelta > Math.Max(240, preferred.RecordCount / 6) ||
                    blockDelta > 1;
                if (!clearlyBetter)
                {
                    return preferred.Field;
                }
            }
        }

        return best.Field;
    }

    private static List<int> GetStrongLines((BulkCaptureAnalyzer.RecordSlice Record, BulkCaptureAnalyzer.DecodedHeader Header)[] fieldRecords)
    {
        return fieldRecords
            .Where(static item => item.Header.Line > 0)
            .GroupBy(static item => item.Header.Line)
            .OrderBy(static group => group.Key)
            .Where(static group => group.Select(static item => item.Header.Block).Distinct().Count() >= 4)
            .Select(static group => group.Key)
            .ToList();
    }

    // Copy one TM6000 block payload into the packed UYVY raster and mark those
    // pixels as freshly written for this update.
    private static void CopyPackedRow(ReadOnlySpan<byte> payload, byte[] packedPixels, bool[] writtenPixels, int width, int row, int columnStart)
    {
        var rowStride = width * 2;
        var rowOffset = row * rowStride;
        var columnByteOffset = columnStart * 2;
        if (rowOffset >= packedPixels.Length || columnByteOffset >= rowStride)
        {
            return;
        }

        var copyLength = Math.Min(payload.Length, rowStride - columnByteOffset);
        payload[..copyLength].CopyTo(packedPixels.AsSpan(rowOffset + columnByteOffset, copyLength));

        var pixelsWritten = copyLength / 2;
        for (var pixelIndex = 0; pixelIndex < pixelsWritten; pixelIndex++)
        {
            var x = columnStart + pixelIndex;
            if (x >= width)
            {
                break;
            }

            writtenPixels[(row * width) + x] = true;
        }
    }

    private static Plane CropPlaneWidth(Plane plane, int targetWidth)
    {
        var startPixel = Math.Max((plane.Width - targetWidth) / 2, 0);
        var packedPixels = new byte[targetWidth * plane.Height * 2];
        var writtenPixels = new bool[targetWidth * plane.Height];

        for (var row = 0; row < plane.Height; row++)
        {
            var sourcePackedOffset = (row * plane.Width + startPixel) * 2;
            var sourceWrittenOffset = row * plane.Width + startPixel;
            var targetPackedOffset = row * targetWidth * 2;
            var targetWrittenOffset = row * targetWidth;
            Array.Copy(plane.PackedPixels, sourcePackedOffset, packedPixels, targetPackedOffset, targetWidth * 2);
            Array.Copy(plane.WrittenPixels, sourceWrittenOffset, writtenPixels, targetWrittenOffset, targetWidth);
        }

        return new Plane(targetWidth, plane.Height, packedPixels, writtenPixels);
    }

    private static void CopyRow(Plane source, int sourceRow, byte[] targetPackedPixels, bool[] targetWrittenPixels, int targetRow)
    {
        var byteCount = source.Width * 2;
        Array.Copy(source.PackedPixels, sourceRow * byteCount, targetPackedPixels, targetRow * byteCount, byteCount);
        Array.Copy(source.WrittenPixels, sourceRow * source.Width, targetWrittenPixels, targetRow * source.Width, source.Width);
    }

    private static void BlendRows(Plane source, int rowA, int rowB, byte[] targetPackedPixels, bool[] targetWrittenPixels, int targetRow)
    {
        var byteCount = source.Width * 2;
        var offsetA = rowA * byteCount;
        var offsetB = rowB * byteCount;
        var targetOffset = targetRow * byteCount;
        for (var index = 0; index < byteCount; index++)
        {
            targetPackedPixels[targetOffset + index] = (byte)((source.PackedPixels[offsetA + index] + source.PackedPixels[offsetB + index]) >> 1);
        }

        var pixelOffsetA = rowA * source.Width;
        var pixelOffsetB = rowB * source.Width;
        var targetPixelOffset = targetRow * source.Width;
        for (var index = 0; index < source.Width; index++)
        {
            targetWrittenPixels[targetPixelOffset + index] =
                source.WrittenPixels[pixelOffsetA + index] ||
                source.WrittenPixels[pixelOffsetB + index];
        }
    }

    private static int SelectDominantParity(Plane plane)
    {
        var evenCoverage = 0;
        var oddCoverage = 0;
        for (var row = 0; row < plane.Height; row++)
        {
            var rowOffset = row * plane.Width;
            for (var x = 0; x < plane.Width; x++)
            {
                if (!plane.WrittenPixels[rowOffset + x])
                {
                    continue;
                }

                if ((row & 1) == 0)
                {
                    evenCoverage++;
                }
                else
                {
                    oddCoverage++;
                }
            }
        }

        return evenCoverage >= oddCoverage ? 0 : 1;
    }

    private static int FindNearestParityRow(int height, int startRow, int step, int parity)
    {
        for (var row = startRow; row >= 0 && row < height; row += step)
        {
            if ((row & 1) == parity)
            {
                return row;
            }
        }

        return -1;
    }

    // Normalize luma per frame so the preview remains readable even when the analog
    // source is dark or noisy.
    private static (int Min, int Max) ComputeLumaRange(byte[] packedPixels, bool[] writtenPixels, int width, int height)
    {
        var min = 255;
        var max = 0;

        for (var row = 0; row < height; row++)
        {
            for (var x = 0; x < width; x++)
            {
                var pixelIndex = (row * width) + x;
                if (!writtenPixels[pixelIndex])
                {
                    continue;
                }

                var yIndex = GetLumaIndex(row, x, width);
                if (yIndex < 0 || yIndex >= packedPixels.Length)
                {
                    continue;
                }

                min = Math.Min(min, packedPixels[yIndex]);
                max = Math.Max(max, packedPixels[yIndex]);
            }
        }

        return max <= min ? (0, 255) : (min, max);
    }

    // UYVY -> RGB conversion for the final preview bitmap shown in WinForms.
    private static Bitmap BuildColorBitmap(int width, int height, byte[] packedPixels, bool[] writtenPixels, (int Min, int Max) lumaRange)
    {
        var bitmap = new Bitmap(Math.Max(width, 1), Math.Max(height, 1), PixelFormat.Format24bppRgb);
        var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        var data = bitmap.LockBits(rect, ImageLockMode.WriteOnly, bitmap.PixelFormat);

        try
        {
            var stride = data.Stride;
            var buffer = new byte[stride * bitmap.Height];

            for (var row = 0; row < bitmap.Height; row++)
            {
                var rowOffset = row * width * 2;
                for (var x = 0; x < bitmap.Width; x += 2)
                {
                    var pairOffset = rowOffset + (x * 2);
                    if (pairOffset + 3 >= packedPixels.Length)
                    {
                        break;
                    }

                    var u = packedPixels[pairOffset];
                    var y0 = packedPixels[pairOffset + 1];
                    var v = packedPixels[pairOffset + 2];
                    var y1 = packedPixels[pairOffset + 3];

                    WriteColorPixel(buffer, stride, row, x, width, writtenPixels, ScaleLuma(y0, lumaRange), u, v);
                    if (x + 1 < width)
                    {
                        WriteColorPixel(buffer, stride, row, x + 1, width, writtenPixels, ScaleLuma(y1, lumaRange), u, v);
                    }
                }
            }

            Marshal.Copy(buffer, 0, data.Scan0, buffer.Length);
        }
        finally
        {
            bitmap.UnlockBits(data);
        }

        return bitmap;
    }

    private static int GetLumaIndex(int row, int x, int width)
    {
        var rowOffset = row * width * 2;
        var pairOffset = rowOffset + ((x / 2) * 4);
        return pairOffset + (x % 2 == 0 ? 1 : 3);
    }

    private static byte ScaleLuma(byte y, (int Min, int Max) range)
    {
        if (range.Max <= range.Min)
        {
            return y;
        }

        var scaled = (y - range.Min) * 255 / Math.Max(range.Max - range.Min, 1);
        return (byte)Math.Max(0, Math.Min(255, scaled));
    }

    private static void WriteColorPixel(byte[] buffer, int stride, int row, int x, int width, bool[] writtenPixels, byte y, byte u, byte v)
    {
        if (!writtenPixels[(row * width) + x])
        {
            return;
        }

        var c = y - 16;
        var d = u - 128;
        var e = v - 128;

        var r = Clamp((298 * c + 409 * e + 128) >> 8);
        var g = Clamp((298 * c - 100 * d - 208 * e + 128) >> 8);
        var b = Clamp((298 * c + 516 * d + 128) >> 8);

        var offset = (row * stride) + (x * 3);
        buffer[offset] = (byte)b;
        buffer[offset + 1] = (byte)g;
        buffer[offset + 2] = (byte)r;
    }

    private static int Clamp(int value) => Math.Max(0, Math.Min(255, value));

    private sealed record Plane(
        int Width,
        int Height,
        byte[] PackedPixels,
        bool[] WrittenPixels);

    private sealed record FieldStats(
        int Field,
        int ActiveLineCount,
        int RecordCount,
        int DistinctBlockCount);
}
