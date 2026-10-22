using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MinsPDFViewer
{
    public sealed class OcrPageData
    {
        public int PageIndex { get; set; }
        public double WidthPoint { get; set; }
        public double HeightPoint { get; set; }
        public List<OcrWordInfo> Words { get; set; } = new();
    }

    internal sealed class PdfGlyphlessOcrService
    {
        private const string FontResName = "FOcr";
        private readonly Action<string>? _log;

        // 532-byte minimal TrueType font with empty outlines (adv=500, upm=1000)
        private static readonly byte[] GlyphlessTtf = new byte[]
        {
            0x00, 0x01, 0x00, 0x00, 0x00, 0x0A, 0x00, 0x80, 0x00, 0x03, 0x00, 0x20, 0x4F, 0x53, 0x2F, 0x32,
            0x8A, 0x59, 0x7E, 0x3C, 0x00, 0x00, 0x00, 0xB4, 0x00, 0x00, 0x00, 0x60, 0x63, 0x6D, 0x61, 0x70,
            0x3B, 0xFA, 0x00, 0x3D, 0x00, 0x00, 0x01, 0x14, 0x00, 0x00, 0x00, 0x30, 0x67, 0x6C, 0x79, 0x66,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x44, 0x00, 0x00, 0x00, 0x00, 0x68, 0x65, 0x61, 0x64,
            0x0A, 0x7A, 0x4F, 0x81, 0x00, 0x00, 0x01, 0x44, 0x00, 0x00, 0x00, 0x36, 0x68, 0x68, 0x65, 0x61,
            0x09, 0xD4, 0x05, 0xDF, 0x00, 0x00, 0x01, 0x7C, 0x00, 0x00, 0x00, 0x24, 0x68, 0x6D, 0x74, 0x78,
            0x03, 0xE8, 0x03, 0xE8, 0x00, 0x00, 0x01, 0xA0, 0x00, 0x00, 0x00, 0x08, 0x6C, 0x6F, 0x63, 0x61,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0xA8, 0x00, 0x00, 0x00, 0x06, 0x6D, 0x61, 0x78, 0x70,
            0x00, 0x04, 0x00, 0x05, 0x00, 0x00, 0x01, 0xB0, 0x00, 0x00, 0x00, 0x20, 0x6E, 0x61, 0x6D, 0x65,
            0x88, 0xBC, 0x07, 0x98, 0x00, 0x00, 0x01, 0xD0, 0x00, 0x00, 0x00, 0x6C, 0x70, 0x6F, 0x73, 0x74,
            0xFF, 0xDB, 0x00, 0x31, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x01, 0x01, 0xF4,
            0x01, 0x90, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x20,
            0xFF, 0xFF, 0x03, 0xE8, 0x00, 0x00, 0x00, 0x00, 0x03, 0xE8, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x03, 0x00, 0x01, 0x00, 0x00, 0x00, 0x0C,
            0x00, 0x04, 0x00, 0x20, 0x00, 0x00, 0x00, 0x04, 0x00, 0x04, 0x00, 0x01, 0x00, 0x00, 0x00, 0x20,
            0xFF, 0xFF, 0x00, 0x00, 0x00, 0x20, 0xFF, 0xFF, 0xFF, 0xE1, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x6C, 0xA2, 0xD4, 0xDB, 0x5F, 0x0F, 0x3C, 0xF5,
            0x00, 0x0B, 0x03, 0xE8, 0x00, 0x00, 0x00, 0x00, 0xE1, 0xB0, 0x9B, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xE1, 0xB0, 0x9B, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0xF4, 0x03, 0xE8, 0x00, 0x00, 0x00, 0x03,
            0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x03, 0xE8, 0x00, 0x00,
            0x00, 0x00, 0x01, 0xF4, 0x00, 0x00, 0x00, 0x00, 0x01, 0xF4, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x01, 0xF4, 0x00, 0x00,
            0x01, 0xF4, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00,
            0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x1E,
            0x00, 0x03, 0x00, 0x01, 0x04, 0x09, 0x00, 0x01, 0x00, 0x1A, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01,
            0x04, 0x09, 0x00, 0x04, 0x00, 0x1A, 0x00, 0x00, 0x00, 0x47, 0x00, 0x6C, 0x00, 0x79, 0x00, 0x70,
            0x00, 0x68, 0x00, 0x4C, 0x00, 0x65, 0x00, 0x73, 0x00, 0x73, 0x00, 0x46, 0x00, 0x6F, 0x00, 0x6E,
            0x00, 0x74, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
        };

        public PdfGlyphlessOcrService(Action<string>? log = null)
        {
            _log = log;
        }

        public async Task InjectOcrLayerAsync(string pdfPath, List<OcrPageData> ocrPages)
        {
            if (ocrPages == null || ocrPages.Count == 0 || !ocrPages.Any(p => p.Words.Count > 0))
                return;

            await Task.Run(() => InjectOcrLayer(pdfPath, ocrPages));
        }

        public void InjectOcrLayer(string pdfPath, List<OcrPageData> ocrPages)
        {
            byte[] originalBytes = File.ReadAllBytes(pdfPath);
            string originalText = Encoding.Latin1.GetString(originalBytes);

            var trailer = ReadTrailer(originalText);
            var objects = ReadObjects(originalText);
            if (!objects.TryGetValue((trailer.RootObjectNumber, trailer.RootGeneration), out var rootObject))
                throw new InvalidOperationException("PDF catalog object was not found.");

            var pageRefs = ReadPageReferences(objects, rootObject.Content);
            if (pageRefs.Count == 0)
                throw new InvalidOperationException("PDF page tree could not be parsed.");

            int nextObjectNumber = trailer.Size;
            var newObjects = new List<(int Number, long Offset, byte[] Data)>();

            // 1. Build Shared Glyphless Font Objects
            // 1.1 FontFile2 (Compressed TTF)
            byte[] fontFileCompressed = CompressZlib(GlyphlessTtf);
            int fontFileNum = nextObjectNumber++;
            string fontFileHeader = $"<< /Length {fontFileCompressed.Length} /Filter /FlateDecode /Length1 {GlyphlessTtf.Length} >>\nstream\n";
            newObjects.Add((fontFileNum, 0, BuildStreamObject(fontFileNum, 0, fontFileHeader, fontFileCompressed)));

            // 1.2 CIDToGIDMap (65536 CIDs -> GID 1: \x00\x01 * 65536)
            byte[] cidToGidBytes = new byte[65536 * 2];
            for (int i = 0; i < 65536; i++)
            {
                cidToGidBytes[i * 2] = 0x00;
                cidToGidBytes[i * 2 + 1] = 0x01;
            }
            byte[] cidToGidCompressed = CompressZlib(cidToGidBytes);
            int cidToGidNum = nextObjectNumber++;
            string cidToGidHeader = $"<< /Length {cidToGidCompressed.Length} /Filter /FlateDecode >>\nstream\n";
            newObjects.Add((cidToGidNum, 0, BuildStreamObject(cidToGidNum, 0, cidToGidHeader, cidToGidCompressed)));

            // 1.3 ToUnicode CMap (Identity CMap for 256 blocks)
            var cmapSb = new StringBuilder();
            cmapSb.Append("/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n");
            cmapSb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            cmapSb.Append("/CMapName /Adobe-Identity-UCS def /CMapType 2 def\n");
            cmapSb.Append("1 begincodespacerange <0000> <FFFF> endcodespacerange\n");
            cmapSb.Append("256 beginbfrange\n");
            for (int h = 0; h < 256; h++)
            {
                cmapSb.Append($"<{h:X2}00> <{h:X2}FF> <{h:X2}00>\n");
            }
            cmapSb.Append("endbfrange\nendcmap CMapName currentdict /CMap defineresource pop end end\n");
            byte[] cmapBytes = Encoding.ASCII.GetBytes(cmapSb.ToString());
            byte[] cmapCompressed = CompressZlib(cmapBytes);
            int toUnicodeNum = nextObjectNumber++;
            string toUnicodeHeader = $"<< /Length {cmapCompressed.Length} /Filter /FlateDecode >>\nstream\n";
            newObjects.Add((toUnicodeNum, 0, BuildStreamObject(toUnicodeNum, 0, toUnicodeHeader, cmapCompressed)));

            // 1.4 FontDescriptor
            int fontDescNum = nextObjectNumber++;
            string fontDesc = $"<< /Type /FontDescriptor /FontName /GlyphLessFont /Flags 5 /FontBBox [0 0 500 1000] /ItalicAngle 0 /Ascent 1000 /Descent 0 /CapHeight 1000 /StemV 80 /FontFile2 {fontFileNum} 0 R >>";
            newObjects.Add((fontDescNum, 0, Encoding.Latin1.GetBytes(WrapObject(fontDescNum, 0, fontDesc))));

            // 1.5 CIDFontType2
            int cidFontNum = nextObjectNumber++;
            string cidFont = $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /GlyphLessFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {fontDescNum} 0 R /DW 500 /CIDToGIDMap {cidToGidNum} 0 R >>";
            newObjects.Add((cidFontNum, 0, Encoding.Latin1.GetBytes(WrapObject(cidFontNum, 0, cidFont))));

            // 1.6 Type0 Font
            int type0Num = nextObjectNumber++;
            string type0 = $"<< /Type /Font /Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H /DescendantFonts [{cidFontNum} 0 R] /ToUnicode {toUnicodeNum} 0 R >>";
            newObjects.Add((type0Num, 0, Encoding.Latin1.GetBytes(WrapObject(type0Num, 0, type0))));

            // 2. Generate and attach Content Stream for each OCR Page
            int injectedWordCount = 0;
            foreach (var pageData in ocrPages)
            {
                if (pageData.Words == null || pageData.Words.Count == 0)
                    continue;

                if (pageData.PageIndex < 0 || pageData.PageIndex >= pageRefs.Count)
                    continue;

                var pageRef = pageRefs[pageData.PageIndex];
                if (!objects.TryGetValue((pageRef.ObjectNumber, pageRef.Generation), out var pageObj))
                    continue;

                // Build transparent text stream ops (using recursive XY-cut ordering and spacing)
                string streamOps = BuildPageOcrStream(pageData);
                if (string.IsNullOrEmpty(streamOps))
                    continue;

                byte[] rawOpsBytes = Encoding.Latin1.GetBytes(streamOps);
                byte[] compOpsBytes = CompressZlib(rawOpsBytes);

                int contentStreamNum = nextObjectNumber++;
                string contentStreamHeader = $"<< /Length {compOpsBytes.Length} /Filter /FlateDecode >>\nstream\n";
                newObjects.Add((contentStreamNum, 0, BuildStreamObject(contentStreamNum, 0, contentStreamHeader, compOpsBytes)));

                // Update Page Object (Contents and Resources)
                string updatedPageContent = UpdatePageObject(pageObj.Content, contentStreamNum, type0Num, objects, newObjects);
                newObjects.Add((pageRef.ObjectNumber, 0, Encoding.Latin1.GetBytes(WrapObject(pageRef.ObjectNumber, pageRef.Generation, updatedPageContent))));

                injectedWordCount += pageData.Words.Count;
            }

            if (injectedWordCount == 0)
            {
                _log?.Invoke("No OCR words to inject.");
                return;
            }

            // 3. Append incremental objects, xref, and trailer
            var objectStream = new MemoryStream();
            long appendStart = originalBytes.Length;
            long currentOffset = appendStart;

            var xrefEntries = new List<(int Number, long Offset)>();

            for (int i = 0; i < newObjects.Count; i++)
            {
                var obj = newObjects[i];
                xrefEntries.Add((obj.Number, currentOffset));
                objectStream.Write(obj.Data, 0, obj.Data.Length);
                currentOffset += obj.Data.Length;
            }

            long xrefOffset = currentOffset;
            string xref = BuildXrefAndTrailer(xrefEntries, trailer, nextObjectNumber, xrefOffset);
            byte[] xrefBytes = Encoding.Latin1.GetBytes(xref);
            objectStream.Write(xrefBytes, 0, xrefBytes.Length);

            using (var fileStream = new FileStream(pdfPath, FileMode.Append, FileAccess.Write, FileShare.None))
            {
                byte[] appendBytes = objectStream.ToArray();
                fileStream.Write(appendBytes, 0, appendBytes.Length);
            }

            _log?.Invoke($"[GlyphlessOCR] Injected OCR text layer. Words={injectedWordCount}, ObjectsAdded={newObjects.Count}");
        }

        private static string BuildPageOcrStream(OcrPageData pageData)
        {
            var words = pageData.Words.Where(w => !string.IsNullOrWhiteSpace(w.Text) && w.BoundingBox.Width > 0 && w.BoundingBox.Height > 0).ToList();
            if (words.Count == 0) return string.Empty;

            double W = pageData.WidthPoint;
            double H = pageData.HeightPoint;

            // Group into lines and sort using recursive XY-cut (order_lines)
            var lines = GroupAndOrderLines(words);

            var sb = new StringBuilder();
            sb.Append("q\nBT\n3 Tr\n0 Tc\n0 Tw\n");

            foreach (var line in lines)
            {
                for (int i = 0; i < line.Count; i++)
                {
                    var w = line[i];
                    string cleanText = CleanText(w.Text);
                    if (string.IsNullOrEmpty(cleanText)) continue;

                    double x = w.BoundingBox.X;
                    double y = w.BoundingBox.Y;
                    double width = w.BoundingBox.Width;
                    double height = w.BoundingBox.Height;

                    double fs = Math.Max(1.0, height);
                    double ownW = Math.Max(0.5, width);

                    string outText = cleanText;
                    double target = ownW;
                    int nChar = cleanText.Length;

                    // If there is a next word in the same line, append space and stretch Tz to next word
                    if (i + 1 < line.Count)
                    {
                        var nextW = line[i + 1];
                        double adv = nextW.BoundingBox.X - x;
                        outText = cleanText + " ";
                        if (ownW * 0.6 <= adv && adv <= ownW + 4.0 * fs)
                        {
                            target = adv;
                            nChar = outText.Length;
                        }
                    }

                    double tz = Math.Clamp(100.0 * target / (nChar * 0.5 * fs), 5.0, 1000.0);

                    // PDF coordinate: origin at bottom-left
                    double bx = x;
                    double by = H - (y + height);

                    string hex = ToHexUtf16Be(outText);
                    sb.Append(string.Format(CultureInfo.InvariantCulture,
                        "/{0} {1:0.##} Tf {2:0.##} Tz 1 0 0 1 {3:0.##} {4:0.##} Tm <{5}> Tj\n",
                        FontResName, fs, tz, bx, by, hex));
                }
            }

            sb.Append("ET\nQ\n");
            return sb.ToString();
        }

        private static List<List<OcrWordInfo>> GroupAndOrderLines(List<OcrWordInfo> words)
        {
            if (words.Count <= 1)
                return new List<List<OcrWordInfo>> { words };

            // Initial line grouping by vertical proximity
            var sortedWords = words.OrderBy(w => w.BoundingBox.Y).ThenBy(w => w.BoundingBox.X).ToList();
            var rawLines = new List<List<OcrWordInfo>>();

            foreach (var w in sortedWords)
            {
                bool placed = false;
                foreach (var line in rawLines)
                {
                    var lastW = line.Last();
                    double avgH = (lastW.BoundingBox.Height + w.BoundingBox.Height) / 2.0;
                    double yDiff = Math.Abs((lastW.BoundingBox.Y + lastW.BoundingBox.Height / 2.0) - (w.BoundingBox.Y + w.BoundingBox.Height / 2.0));

                    if (yDiff <= avgH * 0.5)
                    {
                        line.Add(w);
                        placed = true;
                        break;
                    }
                }

                if (!placed)
                {
                    rawLines.Add(new List<OcrWordInfo> { w });
                }
            }

            // Sort words within each line by X
            foreach (var line in rawLines)
            {
                line.Sort((a, b) => a.BoundingBox.X.CompareTo(b.BoundingBox.X));
            }

            // Recursive XY-cut ordering on lines
            return RecursiveXyCut(rawLines);
        }

        private static (double MinX, double MinY, double MaxX, double MaxY) GetLineBbox(List<OcrWordInfo> line)
        {
            double minX = line.Min(w => w.BoundingBox.X);
            double minY = line.Min(w => w.BoundingBox.Y);
            double maxX = line.Max(w => w.BoundingBox.X + w.BoundingBox.Width);
            double maxY = line.Max(w => w.BoundingBox.Y + w.BoundingBox.Height);
            return (minX, minY, maxX, maxY);
        }

        private static List<List<OcrWordInfo>> RecursiveXyCut(List<List<OcrWordInfo>> lines)
        {
            if (lines.Count <= 1)
                return lines;

            var boxes = lines.Select(GetLineBbox).ToList();
            var heights = boxes.Select(b => b.MaxY - b.MinY).OrderBy(h => h).ToList();
            double lh = Math.Max(1.0, heights[heights.Count / 2]);

            List<int> Cut(List<int> indices, int depth)
            {
                if (indices.Count <= 1 || depth > 12)
                {
                    return indices.OrderBy(i => Math.Round(((boxes[i].MinY + boxes[i].MaxY) / 2.0) / (0.6 * lh)))
                                  .ThenBy(i => boxes[i].MinX)
                                  .ToList();
                }

                // X gaps (column gap: >= 1.5 * lh)
                var xSpans = indices.Select(i => (Min: boxes[i].MinX, Max: boxes[i].MaxX, Index: i))
                                    .OrderBy(s => s.Min).ToList();
                double maxGap = 0;
                double bestSplitX = 0;
                double reach = xSpans[0].Max;
                for (int i = 1; i < xSpans.Count; i++)
                {
                    double gap = xSpans[i].Min - reach;
                    if (gap >= 1.5 * lh && gap > maxGap)
                    {
                        maxGap = gap;
                        bestSplitX = (reach + xSpans[i].Min) / 2.0;
                    }
                    reach = Math.Max(reach, xSpans[i].Max);
                }

                if (maxGap > 0)
                {
                    var left = indices.Where(i => boxes[i].MaxX <= bestSplitX).ToList();
                    var right = indices.Where(i => boxes[i].MaxX > bestSplitX).ToList();
                    if (left.Count > 0 && right.Count > 0)
                    {
                        var res = Cut(left, depth + 1);
                        res.AddRange(Cut(right, depth + 1));
                        return res;
                    }
                }

                // Y gaps (paragraph gap: >= 1.2 * lh)
                var ySpans = indices.Select(i => (Min: boxes[i].MinY, Max: boxes[i].MaxY, Index: i))
                                    .OrderBy(s => s.Min).ToList();
                double maxGapY = 0;
                double bestSplitY = 0;
                double reachY = ySpans[0].Max;
                for (int i = 1; i < ySpans.Count; i++)
                {
                    double gap = ySpans[i].Min - reachY;
                    if (gap >= 1.2 * lh && gap > maxGapY)
                    {
                        maxGapY = gap;
                        bestSplitY = (reachY + ySpans[i].Min) / 2.0;
                    }
                    reachY = Math.Max(reachY, ySpans[i].Max);
                }

                if (maxGapY > 0)
                {
                    var top = indices.Where(i => boxes[i].MaxY <= bestSplitY).ToList();
                    var bottom = indices.Where(i => boxes[i].MaxY > bestSplitY).ToList();
                    if (top.Count > 0 && bottom.Count > 0)
                    {
                        var res = Cut(top, depth + 1);
                        res.AddRange(Cut(bottom, depth + 1));
                        return res;
                    }
                }

                return indices.OrderBy(i => Math.Round(((boxes[i].MinY + boxes[i].MaxY) / 2.0) / (0.6 * lh)))
                              .ThenBy(i => boxes[i].MinX)
                              .ToList();
            }

            var orderedIndices = Cut(Enumerable.Range(0, lines.Count).ToList(), 0);
            return orderedIndices.Select(i => lines[i]).ToList();
        }

        private static string UpdatePageObject(
            string pageContent,
            int newContentStreamNum,
            int type0FontNum,
            Dictionary<(int Number, int Generation), PdfObjectInfo> objects,
            List<(int Number, long Offset, byte[] Data)> newObjects)
        {
            string dict = ExtractDictionary(pageContent, pageContent.IndexOf("<<", StringComparison.Ordinal));

            // 1. Update /Contents
            var contentsMatch = Regex.Match(dict, @"/Contents\s+(\[[^\]]*\]|\d+\s+\d+\s+R)");
            if (contentsMatch.Success)
            {
                string origContents = contentsMatch.Groups[1].Value.Trim();
                string newContents;
                if (origContents.StartsWith("["))
                {
                    string inner = origContents.Substring(1, origContents.Length - 2).Trim();
                    newContents = $"[{inner} {newContentStreamNum} 0 R]";
                }
                else
                {
                    newContents = $"[{origContents} {newContentStreamNum} 0 R]";
                }
                dict = dict.Substring(0, contentsMatch.Index) + $"/Contents {newContents}" + dict.Substring(contentsMatch.Index + contentsMatch.Length);
            }
            else
            {
                int endIdx = dict.LastIndexOf(">>", StringComparison.Ordinal);
                dict = dict.Substring(0, endIdx) + $" /Contents {newContentStreamNum} 0 R >>";
            }

            // 2. Update /Resources /Font
            var resMatch = Regex.Match(dict, @"/Resources\s+(\d+\s+\d+\s+R|<<[\s\S]*?>>)");
            if (resMatch.Success)
            {
                string resVal = resMatch.Groups[1].Value;
                if (resVal.StartsWith("<<"))
                {
                    string updatedRes = InjectFontIntoResourcesDict(resVal, type0FontNum);
                    dict = dict.Substring(0, resMatch.Index) + $"/Resources {updatedRes}" + dict.Substring(resMatch.Index + resMatch.Length);
                }
                else
                {
                    // Indirect object reference for Resources (e.g. /Resources 5 0 R)
                    var refMatch = Regex.Match(resVal, @"(\d+)\s+(\d+)\s+R");
                    if (refMatch.Success)
                    {
                        int resNum = int.Parse(refMatch.Groups[1].Value);
                        int resGen = int.Parse(refMatch.Groups[2].Value);
                        if (objects.TryGetValue((resNum, resGen), out var resObj))
                        {
                            string updatedRes = InjectFontIntoResourcesDict(resObj.Content, type0FontNum);
                            objects[(resNum, resGen)] = new PdfObjectInfo(resNum, resGen, updatedRes);

                            byte[] resBytes = Encoding.Latin1.GetBytes(WrapObject(resNum, resGen, updatedRes));
                            int existingIdx = newObjects.FindIndex(o => o.Number == resNum);
                            if (existingIdx >= 0)
                                newObjects[existingIdx] = (resNum, 0, resBytes);
                            else
                                newObjects.Add((resNum, 0, resBytes));
                        }
                    }
                }
            }
            else
            {
                int endIdx = dict.LastIndexOf(">>", StringComparison.Ordinal);
                dict = dict.Substring(0, endIdx) + $" /Resources << /Font << /{FontResName} {type0FontNum} 0 R >> >> >>";
            }

            return dict;
        }

        private static string InjectFontIntoResourcesDict(string resDict, int type0FontNum)
        {
            var fontMatch = Regex.Match(resDict, @"/Font\s+(<<[\s\S]*?>>)");
            if (fontMatch.Success)
            {
                string fontSubDict = fontMatch.Groups[1].Value;
                int close = fontSubDict.LastIndexOf(">>", StringComparison.Ordinal);
                string newFontSub = fontSubDict.Substring(0, close) + $" /{FontResName} {type0FontNum} 0 R >>";
                return resDict.Substring(0, fontMatch.Index) + $"/Font {newFontSub}" + resDict.Substring(fontMatch.Index + fontMatch.Length);
            }
            else
            {
                int close = resDict.LastIndexOf(">>", StringComparison.Ordinal);
                if (close >= 0)
                {
                    return resDict.Substring(0, close) + $" /Font << /{FontResName} {type0FontNum} 0 R >> >>";
                }
                return resDict;
            }
        }

        private static byte[] CompressZlib(byte[] data)
        {
            using var ms = new MemoryStream();
            using (var zlib = new ZLibStream(ms, CompressionLevel.Optimal, true))
            {
                zlib.Write(data, 0, data.Length);
            }
            return ms.ToArray();
        }

        private static byte[] BuildStreamObject(int number, int generation, string header, byte[] streamBytes)
        {
            var sb = new StringBuilder();
            sb.Append($"{number} {generation} obj\n");
            sb.Append(header);

            byte[] prefix = Encoding.Latin1.GetBytes(sb.ToString());
            byte[] suffix = Encoding.Latin1.GetBytes("\nendstream\nendobj\n");

            byte[] result = new byte[prefix.Length + streamBytes.Length + suffix.Length];
            Buffer.BlockCopy(prefix, 0, result, 0, prefix.Length);
            Buffer.BlockCopy(streamBytes, 0, result, prefix.Length, streamBytes.Length);
            Buffer.BlockCopy(suffix, 0, result, prefix.Length + streamBytes.Length, suffix.Length);
            return result;
        }

        private static string WrapObject(int number, int generation, string content)
        {
            return $"{number} {generation} obj\n{content}\nendobj\n";
        }

        private static string CleanText(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var sb = new StringBuilder();
            foreach (char ch in s.Replace('\n', ' ').Replace('\t', ' '))
            {
                if (ch >= 0x20 && !char.IsSurrogate(ch))
                    sb.Append(ch);
            }
            return sb.ToString().Trim();
        }

        private static string ToHexUtf16Be(string s)
        {
            var sb = new StringBuilder();
            foreach (char ch in s)
            {
                sb.Append(((int)ch).ToString("X4"));
            }
            return sb.ToString();
        }

        private static string BuildXrefAndTrailer(
            List<(int Number, long Offset)> entries,
            TrailerInfo trailer,
            int nextObjectNumber,
            long xrefOffset)
        {
            var sb = new StringBuilder();
            sb.Append("xref\n");

            var sorted = entries.OrderBy(e => e.Number).ToList();
            int i = 0;
            while (i < sorted.Count)
            {
                int start = sorted[i].Number;
                int count = 1;
                while (i + count < sorted.Count && sorted[i + count].Number == start + count)
                    count++;

                sb.Append($"{start} {count}\n");
                for (int j = 0; j < count; j++)
                    sb.Append($"{sorted[i + j].Offset:D10} 00000 n \n");

                i += count;
            }

            sb.Append("trailer\n<< ");
            sb.Append($"/Size {nextObjectNumber} ");
            sb.Append($"/Root {trailer.RootObjectNumber} {trailer.RootGeneration} R ");
            sb.Append($"/Prev {trailer.PreviousXrefOffset} ");
            if (!string.IsNullOrEmpty(trailer.IdRaw))
                sb.Append($"/ID {trailer.IdRaw} ");
            sb.Append(">>\n");
            sb.Append($"startxref\n{xrefOffset}\n%%EOF\n");

            return sb.ToString();
        }

        private static TrailerInfo ReadTrailer(string text)
        {
            int startXref = text.LastIndexOf("startxref", StringComparison.Ordinal);
            if (startXref < 0)
                throw new InvalidOperationException("PDF startxref marker was not found.");

            var startXrefMatch = Regex.Match(text.Substring(startXref), @"startxref\s+(\d+)");
            if (!startXrefMatch.Success)
                throw new InvalidOperationException("PDF startxref offset could not be parsed.");

            long previousXrefOffset = long.Parse(startXrefMatch.Groups[1].Value, CultureInfo.InvariantCulture);
            int trailerIndex = text.IndexOf("trailer", (int)Math.Min(previousXrefOffset, int.MaxValue), StringComparison.Ordinal);
            if (trailerIndex < 0)
                trailerIndex = text.LastIndexOf("trailer", startXref, StringComparison.Ordinal);
            if (trailerIndex < 0)
                throw new InvalidOperationException("PDF trailer was not found.");

            string trailerDictionary = ExtractDictionary(text, text.IndexOf("<<", trailerIndex, StringComparison.Ordinal));
            var root = ReadReference(trailerDictionary, "Root");
            int size = ReadInteger(trailerDictionary, "Size");
            string? id = ReadRawArray(trailerDictionary, "ID");

            return new TrailerInfo(size, root.ObjectNumber, root.Generation, previousXrefOffset, id);
        }

        private static Dictionary<(int Number, int Generation), PdfObjectInfo> ReadObjects(string text)
        {
            var objects = new Dictionary<(int, int), PdfObjectInfo>();
            foreach (Match match in Regex.Matches(text, @"(?s)(\d+)\s+(\d+)\s+obj\s*(.*?)\s*endobj"))
            {
                int number = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                int generation = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                objects[(number, generation)] = new PdfObjectInfo(number, generation, match.Groups[3].Value.Trim());
            }
            return objects;
        }

        private static List<PageRef> ReadPageReferences(Dictionary<(int Number, int Generation), PdfObjectInfo> objects, string catalog)
        {
            var pagesRef = ReadReference(catalog, "Pages");
            var result = new List<PageRef>();
            var visited = new HashSet<(int, int)>();
            VisitPageTree(pagesRef, objects, visited, result);

            if (result.Count == 0)
            {
                foreach (var obj in objects.Values.OrderBy(o => o.Number))
                {
                    if (Regex.IsMatch(obj.Content, @"/Type\s*/Page\b") && !Regex.IsMatch(obj.Content, @"/Type\s*/Pages\b"))
                        result.Add(new PageRef(obj.Number, obj.Generation));
                }
            }

            return result;
        }

        private static void VisitPageTree(
            PageRef node,
            Dictionary<(int Number, int Generation), PdfObjectInfo> objects,
            HashSet<(int, int)> visited,
            List<PageRef> pages)
        {
            if (!visited.Add((node.ObjectNumber, node.Generation)))
                return;

            if (!objects.TryGetValue((node.ObjectNumber, node.Generation), out var obj))
                return;

            if (Regex.IsMatch(obj.Content, @"/Type\s*/Page\b") && !Regex.IsMatch(obj.Content, @"/Type\s*/Pages\b"))
            {
                pages.Add(node);
                return;
            }

            foreach (var child in ReadReferencesFromArray(obj.Content, "Kids"))
                VisitPageTree(child, objects, visited, pages);
        }

        private static List<PageRef> ReadReferencesFromArray(string text, string key)
        {
            var list = new List<PageRef>();
            var match = Regex.Match(text, @"/" + key + @"\s*\[(.*?)\]", RegexOptions.Singleline);
            if (!match.Success)
                return list;

            foreach (Match item in Regex.Matches(match.Groups[1].Value, @"(\d+)\s+(\d+)\s+R"))
            {
                list.Add(new PageRef(
                    int.Parse(item.Groups[1].Value, CultureInfo.InvariantCulture),
                    int.Parse(item.Groups[2].Value, CultureInfo.InvariantCulture)));
            }

            return list;
        }

        private static PageRef ReadReference(string text, string key)
        {
            var match = Regex.Match(text, @"/" + key + @"\s+(\d+)\s+(\d+)\s+R");
            if (!match.Success)
                throw new InvalidOperationException($"Reference for /{key} was not found.");

            return new PageRef(
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        }

        private static int ReadInteger(string text, string key)
        {
            var match = Regex.Match(text, @"/" + key + @"\s+(\d+)");
            if (!match.Success)
                throw new InvalidOperationException($"Integer value for /{key} was not found.");

            return int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        private static string? ReadRawArray(string text, string key)
        {
            var match = Regex.Match(text, @"/" + key + @"\s*(\[[^\]]*\])", RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string ExtractDictionary(string text, int startIndex)
        {
            if (startIndex < 0 || startIndex + 1 >= text.Length || text[startIndex] != '<' || text[startIndex + 1] != '<')
                throw new InvalidOperationException("Dictionary start marker was not found.");

            int depth = 0;
            for (int i = startIndex; i < text.Length - 1; i++)
            {
                if (text[i] == '<' && text[i + 1] == '<')
                {
                    depth++;
                    i++;
                }
                else if (text[i] == '>' && text[i + 1] == '>')
                {
                    depth--;
                    i++;
                    if (depth == 0)
                        return text.Substring(startIndex, i - startIndex + 1);
                }
            }

            throw new InvalidOperationException("Unterminated dictionary in PDF.");
        }
        private sealed record TrailerInfo(int Size, int RootObjectNumber, int RootGeneration, long PreviousXrefOffset, string? IdRaw);
        private sealed record PdfObjectInfo(int Number, int Generation, string Content);
        private sealed record PageRef(int ObjectNumber, int Generation);
    }
}