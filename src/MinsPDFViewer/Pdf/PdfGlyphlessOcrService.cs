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
        /// <summary>저장본에서의 페이지 위치 (0부터).</summary>
        public int PageIndex { get; set; }

        /// <summary>앱에서 더한 회전(도). 저장본 /Rotate에서 빼서 OCR 당시(원본) 회전을 구한다.</summary>
        public int UserRotation { get; set; }

        /// <summary>OCR 좌표계(원본 페이지 화면) 크기, 포인트. 0이면 저장본의 페이지 상자를 그대로 쓴다.</summary>
        public double WidthPoint { get; set; }
        public double HeightPoint { get; set; }

        /// <summary>단어 상자: 포인트, 원본 페이지 화면 기준 왼쪽 위 원점.</summary>
        public List<OcrWordInfo> Words { get; set; } = new();
    }

    /// <summary>
    /// 보이지 않는 OCR 텍스트 층(searchable PDF)을 증분 업데이트로 덧붙인다. Tesseract·MobileScan과 같은 방식:
    /// 글리프 없는 TrueType(내장) + Type0/Identity-H + CIDToGIDMap(전부 GID 1) + ToUnicode(CID = UTF-16),
    /// 렌더 모드 3(투명), 단어마다 Tz로 폭을 맞춘다.
    /// </summary>
    internal sealed class PdfGlyphlessOcrService
    {
        internal const string FontResName = "FMinsOcr";
        private const string OcrStreamMarker = "/MinsOcrLayer true";
        private const string PrefixStreamMarker = "/MinsOcrPrefix true";
        private readonly Action<string>? _log;
        // 글리프 없는 TrueType (632바이트, upm 1000, 전진폭 500, Ascent 880 / Descent -120).
        // GID 1은 넓이 0인 직선 윤곽([0 -120 500 880] 대각선)이라 칠할 것이 없다: 렌더 모드 3을 무시하는 뷰어에서도 안 보인다.
        // 완전히 빈 글리프였던 이전 폰트는 PDFium이 세로로 놓인(회전 페이지) 글자를 텍스트로 인식하지 못했다.
        // fontTools FontBuilder로 생성 (hhea·hmtx·OS/2·cmap·post 정상).
        private static readonly byte[] GlyphlessTtf = Convert.FromBase64String(
            "AAEAAAAKAIAAAwAgT1MvMkT8RFAAAAEoAAAAYGNtYXAADABzAAABkAAAADRnbHlmGOgTdQAAAcwAAAAYaGVhZC8L04QAAACsAAAANmhoZWEFZgF+AAAA5AAAACRobXR4AfQAAAAAAYgAAAAGbG9jYQAMAAAAAAHEAAAABm1heHAABAAFAAABCAAAACBuYW1lGZ8ZNAAAAeQAAABycG9zdAADAAAAAAJYAAAAIAABAAAAAQAASYgP3V8PPPUAAwPoAAAAAOcAR9YAAAAA5wBH1gAA/4gB9ANwAAAAAwACAAAAAAAAAAEAAANw/4gAAAH0AAAAAAH0AAEAAAAAAAAAAAAAAAAAAAABAAEAAAACAAMAAQAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAwH0AZAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAPz8/PwAAACAAIANw/4gAAANwAHgAAAAAAAAAAAAAAAAAAAAgAAAB9AAAAAAAAAAAAAIAAAADAAAAFAADAAEAAAAUAAQAIAAAAAQABAABAAAAIP//AAAAIP///+EAAQAAAAAAAAAAAAwAAAABAAD/iAH0A3AAAgAAFRMT+vp4AfQB9AAAAAQANgABAAAAAAABAA0AAAABAAAAAAACAAcADQADAAEECQABABoAFAADAAEECQACAA4ALkdseXBoTGVzc0ZvbnRSZWd1bGFyAEcAbAB5AHAAaABMAGUAcwBzAEYAbwBuAHQAUgBlAGcAdQBsAGEAcgAAAAMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=");

        public PdfGlyphlessOcrService(Action<string>? log = null)
        {
            _log = log;
        }

        /// <returns>텍스트 층을 넣었으면 true. 넣을 단어가 없거나 암호화된 PDF면 false.</returns>
        public async Task<bool> InjectOcrLayerAsync(string pdfPath, List<OcrPageData> ocrPages)
        {
            if (ocrPages == null || ocrPages.Count == 0 || !ocrPages.Any(p => p.Words.Count > 0))
                return false;

            return await Task.Run(() => InjectOcrLayer(pdfPath, ocrPages));
        }

        public bool InjectOcrLayer(string pdfPath, List<OcrPageData> ocrPages)
        {
            byte[] originalBytes = File.ReadAllBytes(pdfPath);
            string originalText = Encoding.Latin1.GetString(originalBytes);

            var trailer = ReadTrailer(originalText);
            if (trailer.IsEncrypted)
            {
                // 암호화 문서는 새 객체도 암호화해야 한다. 평문 스트림을 덧붙이면 깨지므로 건너뛴다.
                _log?.Invoke("[GlyphlessOCR] Skipped: encrypted PDF (text layer would need encryption).");
                return false;
            }

            var objects = ReadObjects(originalText);
            if (!objects.TryGetValue((trailer.RootObjectNumber, trailer.RootGeneration), out var rootObject))
                throw new InvalidOperationException("PDF catalog object was not found.");

            var pageRefs = ReadPageReferences(objects, rootObject.Content);
            if (pageRefs.Count == 0)
                throw new InvalidOperationException("PDF page tree could not be parsed.");

            int nextObjectNumber = trailer.Size;
            var writes = new List<(int Number, int Generation, byte[] Data)>();
            void Put(int number, int generation, byte[] data)
            {
                int existing = writes.FindIndex(w => w.Number == number && w.Generation == generation);
                if (existing >= 0)
                    writes[existing] = (number, generation, data);
                else
                    writes.Add((number, generation, data));
            }

            // 1. 공유 글꼴 객체
            byte[] fontFileCompressed = CompressZlib(GlyphlessTtf);
            int fontFileNum = nextObjectNumber++;
            Put(fontFileNum, 0, BuildStreamObject(fontFileNum, 0,
                $"<< /Length {fontFileCompressed.Length} /Filter /FlateDecode /Length1 {GlyphlessTtf.Length} >>\nstream\n",
                fontFileCompressed));

            // CIDToGIDMap: 65536개 CID 모두 GID 1(빈 글리프)
            byte[] cidToGidBytes = new byte[65536 * 2];
            for (int i = 1; i < cidToGidBytes.Length; i += 2)
                cidToGidBytes[i] = 0x01;
            byte[] cidToGidCompressed = CompressZlib(cidToGidBytes);
            int cidToGidNum = nextObjectNumber++;
            Put(cidToGidNum, 0, BuildStreamObject(cidToGidNum, 0,
                $"<< /Length {cidToGidCompressed.Length} /Filter /FlateDecode >>\nstream\n", cidToGidCompressed));

            byte[] cmapCompressed = CompressZlib(Encoding.ASCII.GetBytes(BuildIdentityToUnicodeCMap()));
            int toUnicodeNum = nextObjectNumber++;
            Put(toUnicodeNum, 0, BuildStreamObject(toUnicodeNum, 0,
                $"<< /Length {cmapCompressed.Length} /Filter /FlateDecode >>\nstream\n", cmapCompressed));

            int fontDescNum = nextObjectNumber++;
            Put(fontDescNum, 0, Encoding.Latin1.GetBytes(WrapObject(fontDescNum, 0,
                $"<< /Type /FontDescriptor /FontName /GlyphLessFont /Flags 5 /FontBBox [0 -120 500 880] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 880 /StemV 80 /FontFile2 {fontFileNum} 0 R >>")));

            int cidFontNum = nextObjectNumber++;
            Put(cidFontNum, 0, Encoding.Latin1.GetBytes(WrapObject(cidFontNum, 0,
                $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /GlyphLessFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {fontDescNum} 0 R /DW 500 /CIDToGIDMap {cidToGidNum} 0 R >>")));

            int type0Num = nextObjectNumber++;
            Put(type0Num, 0, Encoding.Latin1.GetBytes(WrapObject(type0Num, 0,
                $"<< /Type /Font /Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H /DescendantFonts [{cidFontNum} 0 R] /ToUnicode {toUnicodeNum} 0 R >>")));

            // 원본 내용을 q ... Q로 감싸기 위한 공유 "q" 스트림. 원본이 그래픽 상태(CTM 등)를 남겨도 텍스트 층 위치가 틀어지지 않게.
            byte[] prefixBytes = Encoding.ASCII.GetBytes("q\n");
            int prefixNum = nextObjectNumber++;
            Put(prefixNum, 0, BuildStreamObject(prefixNum, 0, $"<< /Length {prefixBytes.Length} {PrefixStreamMarker} >>\nstream\n", prefixBytes));

            // 2. 페이지마다 텍스트 층 스트림을 만들고 페이지 객체를 갱신
            int injectedWordCount = 0;
            int injectedPageCount = 0;
            foreach (var pageData in ocrPages)
            {
                if (pageData.Words == null || pageData.Words.Count == 0)
                    continue;

                if (pageData.PageIndex < 0 || pageData.PageIndex >= pageRefs.Count)
                {
                    _log?.Invoke($"[GlyphlessOCR] Page index {pageData.PageIndex} is out of range (pages={pageRefs.Count}).");
                    continue;
                }

                var pageRef = pageRefs[pageData.PageIndex];
                if (!objects.TryGetValue((pageRef.ObjectNumber, pageRef.Generation), out var pageObj))
                    continue;

                string pageDict = PdfSyntax.ExtractLeadingDictionary(pageObj.Content);
                var geometry = PageGeometry.Read(pageDict, objects, pageData.UserRotation);

                string streamOps = BuildPageOcrStream(pageData, geometry, out int wordCount);
                if (wordCount == 0)
                    continue;

                byte[] compOpsBytes = CompressZlib(Encoding.Latin1.GetBytes(streamOps));
                int contentStreamNum = nextObjectNumber++;
                Put(contentStreamNum, 0, BuildStreamObject(contentStreamNum, 0,
                    $"<< /Length {compOpsBytes.Length} /Filter /FlateDecode {OcrStreamMarker} >>\nstream\n", compOpsBytes));

                string updatedPage = UpdatePageObject(pageDict, prefixNum, contentStreamNum, type0Num, objects, Put);
                objects[(pageRef.ObjectNumber, pageRef.Generation)] = new PdfObjectInfo(pageRef.ObjectNumber, pageRef.Generation, updatedPage);
                Put(pageRef.ObjectNumber, pageRef.Generation, Encoding.Latin1.GetBytes(WrapObject(pageRef.ObjectNumber, pageRef.Generation, updatedPage)));

                injectedWordCount += wordCount;
                injectedPageCount++;
            }

            if (injectedWordCount == 0)
            {
                _log?.Invoke("No OCR words to inject.");
                return false;
            }

            // 3. 증분 객체 + xref + trailer 덧붙이기
            var objectStream = new MemoryStream();
            long currentOffset = originalBytes.Length;
            // 원본이 줄바꿈 없이 %%EOF로 끝나면 새 객체가 같은 줄에 붙지 않게
            if (originalBytes.Length > 0 && originalBytes[^1] != (byte)'\n' && originalBytes[^1] != (byte)'\r')
            {
                objectStream.WriteByte((byte)'\n');
                currentOffset++;
            }

            var xrefEntries = new List<(int Number, int Generation, long Offset)>();
            foreach (var obj in writes)
            {
                xrefEntries.Add((obj.Number, obj.Generation, currentOffset));
                objectStream.Write(obj.Data, 0, obj.Data.Length);
                currentOffset += obj.Data.Length;
            }

            string xref = BuildXrefAndTrailer(xrefEntries, trailer, nextObjectNumber, currentOffset);
            byte[] xrefBytes = Encoding.Latin1.GetBytes(xref);
            objectStream.Write(xrefBytes, 0, xrefBytes.Length);

            using (var fileStream = new FileStream(pdfPath, FileMode.Append, FileAccess.Write, FileShare.None))
            {
                byte[] appendBytes = objectStream.ToArray();
                fileStream.Write(appendBytes, 0, appendBytes.Length);
            }

            _log?.Invoke($"[GlyphlessOCR] Injected OCR text layer. Pages={injectedPageCount}, Words={injectedWordCount}, ObjectsWritten={writes.Count}");
            return true;
        }

        /// <summary>ToUnicode: CID = UTF-16 코드 단위. beginbfrange 한 블록은 100개까지(PDF 규격) → 100·100·56으로 나눈다.</summary>
        internal static string BuildIdentityToUnicodeCMap()
        {
            var sb = new StringBuilder();
            sb.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
            sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            sb.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
            sb.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
            for (int block = 0; block < 256; block += 100)
            {
                int n = Math.Min(100, 256 - block);
                sb.Append(n).Append(" beginbfrange\n");
                for (int hi = block; hi < block + n; hi++)
                    sb.Append($"<{hi:X2}00> <{hi:X2}FF> <{hi:X2}00>\n");
                sb.Append("endbfrange\n");
            }
            sb.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
            return sb.ToString();
        }

        /// <summary>
        /// 페이지 상자와 회전. OCR 좌표(원본 화면, 왼쪽 위 원점)를 PDF 사용자 공간으로 바꾼다.
        /// 화면 = 유효 CropBox(MediaBox와 교집합)를 원본 /Rotate만큼 시계 방향으로 돌린 것 (PDFium과 같음).
        /// </summary>
        internal readonly record struct PageGeometry(double Llx, double Lly, double Urx, double Ury, int SourceRotation)
        {
            public double ViewWidth => SourceRotation % 180 == 0 ? Urx - Llx : Ury - Lly;
            public double ViewHeight => SourceRotation % 180 == 0 ? Ury - Lly : Urx - Llx;

            public (double X, double Y) ToUser(double vx, double vy) => SourceRotation switch
            {
                90 => (Llx + vy, Lly + vx),
                180 => (Urx - vx, Lly + vy),
                270 => (Urx - vy, Ury - vx),
                _ => (Llx + vx, Ury - vy)
            };

            /// <summary>텍스트 행렬의 축: x축 = 화면 오른쪽, y축 = 화면 위쪽.</summary>
            public (int A, int B, int C, int D) TextAxes => SourceRotation switch
            {
                90 => (0, 1, -1, 0),
                180 => (-1, 0, 0, -1),
                270 => (0, -1, 1, 0),
                _ => (1, 0, 0, 1)
            };

            internal static int NormalizeRotation(int degrees)
            {
                int r = ((degrees % 360) + 360) % 360;
                return (int)(Math.Round(r / 90.0) * 90) % 360;
            }

            internal static PageGeometry Read(string pageDict, Dictionary<(int Number, int Generation), PdfObjectInfo> objects, int userRotation)
            {
                double[] media = ReadBox(GetInheritedValue(pageDict, "MediaBox", objects), objects) ?? new double[] { 0, 0, 612, 792 };
                double[] crop = ReadBox(GetInheritedValue(pageDict, "CropBox", objects), objects) ?? media;

                double llx = Math.Max(crop[0], media[0]);
                double lly = Math.Max(crop[1], media[1]);
                double urx = Math.Min(crop[2], media[2]);
                double ury = Math.Min(crop[3], media[3]);
                if (urx - llx <= 0 || ury - lly <= 0)
                    (llx, lly, urx, ury) = (media[0], media[1], media[2], media[3]);

                int finalRotation = 0;
                string? rotateValue = ResolveValue(GetInheritedValue(pageDict, "Rotate", objects), objects);
                if (rotateValue != null && PdfSyntax.TryParseNumber(rotateValue, out double rot))
                    finalRotation = (int)rot;

                int sourceRotation = NormalizeRotation(NormalizeRotation(finalRotation) - NormalizeRotation(userRotation));
                return new PageGeometry(llx, lly, urx, ury, sourceRotation);
            }

            private static double[]? ReadBox(string? value, Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
            {
                value = ResolveValue(value, objects);
                if (value == null)
                    return null;

                var items = PdfSyntax.ParseArrayItems(value);
                if (items.Count != 4)
                    return null;

                var numbers = new double[4];
                for (int i = 0; i < 4; i++)
                {
                    string item = ResolveValue(items[i], objects) ?? "";
                    if (!PdfSyntax.TryParseNumber(item, out numbers[i]))
                        return null;
                }

                return new[]
                {
                    Math.Min(numbers[0], numbers[2]), Math.Min(numbers[1], numbers[3]),
                    Math.Max(numbers[0], numbers[2]), Math.Max(numbers[1], numbers[3])
                };
            }
        }

        /// <summary>간접 참조면 대상 객체 본문(스트림이면 사전만)을, 아니면 값 그대로 돌려준다.</summary>
        private static string? ResolveValue(string? value, Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
        {
            if (value == null)
                return null;
            if (!PdfSyntax.TryParseReference(value, out int num, out int gen))
                return value;
            if (!objects.TryGetValue((num, gen), out var obj))
                return null;
            string content = obj.Content;
            int streamIdx = content.IndexOf("stream", StringComparison.Ordinal);
            if (content.TrimStart().StartsWith("<<", StringComparison.Ordinal))
                return PdfSyntax.ExtractLeadingDictionary(content);
            return streamIdx >= 0 ? content.Substring(0, streamIdx).Trim() : content.Trim();
        }

        /// <summary>페이지 사전에서 상속 가능한 키(MediaBox, CropBox, Rotate, Resources)를 /Parent를 따라 찾는다.</summary>
        private static string? GetInheritedValue(string pageDict, string key, Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
        {
            string current = pageDict;
            var visited = new HashSet<(int, int)>();
            for (int depth = 0; depth < 64; depth++)
            {
                string? value = PdfSyntax.GetValue(current, key);
                if (value != null)
                    return value;

                string? parent = PdfSyntax.GetValue(current, "Parent");
                if (!PdfSyntax.TryParseReference(parent, out int num, out int gen) || !visited.Add((num, gen)))
                    return null;
                if (!objects.TryGetValue((num, gen), out var parentObj))
                    return null;
                current = PdfSyntax.ExtractLeadingDictionary(parentObj.Content);
            }
            return null;
        }

        private static string BuildPageOcrStream(OcrPageData pageData, PageGeometry geometry, out int wordCount)
        {
            wordCount = 0;
            var words = pageData.Words.Where(w => !string.IsNullOrWhiteSpace(w.Text) && w.BoundingBox.Width > 0 && w.BoundingBox.Height > 0).ToList();
            if (words.Count == 0) return string.Empty;

            // OCR 당시 화면 크기와 저장본 페이지 상자가 다르면(드묾) 비율을 맞춘다.
            double sx = pageData.WidthPoint > 0 ? geometry.ViewWidth / pageData.WidthPoint : 1.0;
            double sy = pageData.HeightPoint > 0 ? geometry.ViewHeight / pageData.HeightPoint : 1.0;
            if (double.IsNaN(sx) || double.IsInfinity(sx) || sx <= 0) sx = 1.0;
            if (double.IsNaN(sy) || double.IsInfinity(sy) || sy <= 0) sy = 1.0;

            var (a, b, c, d) = geometry.TextAxes;

            // Group into lines and sort using recursive XY-cut (order_lines)
            var lines = GroupAndOrderLines(words);

            var sb = new StringBuilder();
            // 앞의 Q는 원본 내용 앞에 넣은 q를 닫는다.
            sb.Append("Q\nq\nBT\n3 Tr\n0 Tc\n0 Tw\n");

            foreach (var line in lines)
            {
                for (int i = 0; i < line.Count; i++)
                {
                    var w = line[i];
                    string cleanText = CleanText(w.Text);
                    if (string.IsNullOrEmpty(cleanText)) continue;

                    double x = w.BoundingBox.X * sx;
                    double y = w.BoundingBox.Y * sy;
                    double width = w.BoundingBox.Width * sx;
                    double height = w.BoundingBox.Height * sy;

                    double fs = Math.Max(1.0, height);
                    double ownW = Math.Max(0.5, width);

                    string outText = cleanText;
                    double target = ownW;
                    int nChar = cleanText.Length;

                    // 같은 줄에 다음 단어가 있으면 공백을 붙이고 다음 단어 시작까지 늘린다 (단어 구분·복사 품질)
                    if (i + 1 < line.Count)
                    {
                        double adv = line[i + 1].BoundingBox.X * sx - x;
                        outText = cleanText + " ";
                        if (ownW * 0.6 <= adv && adv <= ownW + 4.0 * fs)
                        {
                            target = adv;
                            nChar = outText.Length;
                        }
                    }

                    // DW 500 = 글자당 0.5 em
                    double tz = Math.Clamp(100.0 * target / (nChar * 0.5 * fs), 5.0, 1000.0);

                    // 기준선 = 단어 아래 변에서 0.12 em 위 (MobileScan과 같음). Ascent 880/Descent -120이라
                    // 뷰어의 선택 상자가 단어 상자와 정확히 겹친다. Descent 0은 poppler 등이 '없음'으로 보고 -0.35를 써서 상자가 커진다.
                    var (e, f) = geometry.ToUser(x, y + height - 0.12 * fs);

                    string hex = ToHexUtf16Be(outText);
                    sb.Append(string.Format(CultureInfo.InvariantCulture,
                        "/{0} {1:0.###} Tf {2:0.###} Tz {3} {4} {5} {6} {7:0.###} {8:0.###} Tm <{9}> Tj\n",
                        FontResName, fs, tz, a, b, c, d, e, f, hex));
                    wordCount++;
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

        /// <summary>
        /// /Contents = [q-스트림, 원본 내용..., OCR 스트림]. 예전에 넣은 OCR 층(표식 있는 스트림)은 빼서 다시 OCR해도 겹치지 않게.
        /// /Resources /Font에 OCR 글꼴 추가 — 중첩 사전, 간접 참조, 부모에서 상속된 Resources 모두 처리.
        /// </summary>
        private static string UpdatePageObject(
            string pageDict,
            int prefixStreamNum,
            int newContentStreamNum,
            int type0FontNum,
            Dictionary<(int Number, int Generation), PdfObjectInfo> objects,
            Action<int, int, byte[]> put)
        {
            // 1. /Contents
            var contentItems = new List<string>();
            string? contents = PdfSyntax.GetValue(pageDict, "Contents");
            if (contents != null)
            {
                if (contents.StartsWith("[", StringComparison.Ordinal))
                {
                    contentItems.AddRange(PdfSyntax.ParseArrayItems(contents));
                }
                else if (PdfSyntax.TryParseReference(contents, out int cNum, out int cGen) &&
                         objects.TryGetValue((cNum, cGen), out var cObj) &&
                         cObj.Content.TrimStart().StartsWith("[", StringComparison.Ordinal))
                {
                    // /Contents가 배열 객체를 가리키는 경우: 원소를 펼친다
                    contentItems.AddRange(PdfSyntax.ParseArrayItems(cObj.Content.Trim()));
                }
                else
                {
                    contentItems.Add(contents);
                }
            }

            contentItems.RemoveAll(item => IsOwnStream(item, objects));
            var newContents = new StringBuilder("[");
            newContents.Append(prefixStreamNum).Append(" 0 R");
            foreach (var item in contentItems)
                newContents.Append(' ').Append(item.Trim());
            newContents.Append(' ').Append(newContentStreamNum).Append(" 0 R]");
            pageDict = PdfSyntax.SetValue(pageDict, "Contents", newContents.ToString());

            // 2. /Resources
            string fontRef = $"{type0FontNum} 0 R";
            string? resources = PdfSyntax.GetValue(pageDict, "Resources");
            if (resources == null)
            {
                // 상속된 Resources를 페이지에 복사해 넣는다 (그냥 새로 만들면 원본 이미지·글꼴을 잃는다)
                string? parent = PdfSyntax.GetValue(pageDict, "Parent");
                string? inherited = null;
                if (PdfSyntax.TryParseReference(parent, out int pNum, out int pGen) && objects.TryGetValue((pNum, pGen), out var pObj))
                    inherited = GetInheritedValue(PdfSyntax.ExtractLeadingDictionary(pObj.Content), "Resources", objects);
                string baseRes = ResolveValue(inherited, objects) ?? "<< >>";
                if (!baseRes.StartsWith("<<", StringComparison.Ordinal))
                    baseRes = "<< >>";
                pageDict = PdfSyntax.SetValue(pageDict, "Resources", AddFontToResources(baseRes, fontRef, objects));
            }
            else if (PdfSyntax.TryParseReference(resources, out int rNum, out int rGen))
            {
                if (objects.TryGetValue((rNum, rGen), out var resObj))
                {
                    // 공유 Resources 객체: 그 객체를 고친다 (여러 페이지가 함께 써도 키는 한 번만)
                    string resDict = PdfSyntax.ExtractLeadingDictionary(resObj.Content);
                    string updated = AddFontToResources(resDict, fontRef, objects);
                    objects[(rNum, rGen)] = new PdfObjectInfo(rNum, rGen, updated);
                    put(rNum, rGen, Encoding.Latin1.GetBytes(WrapObject(rNum, rGen, updated)));
                }
                else
                {
                    pageDict = PdfSyntax.SetValue(pageDict, "Resources", $"<< /Font << /{FontResName} {fontRef} >> >>");
                }
            }
            else
            {
                pageDict = PdfSyntax.SetValue(pageDict, "Resources", AddFontToResources(resources, fontRef, objects));
            }

            return pageDict;
        }

        private static bool IsOwnStream(string item, Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
        {
            if (!PdfSyntax.TryParseReference(item, out int num, out int gen) || !objects.TryGetValue((num, gen), out var obj))
                return false;
            try
            {
                string dict = PdfSyntax.ExtractLeadingDictionary(obj.Content);
                return PdfSyntax.GetValue(dict, "MinsOcrLayer") != null || PdfSyntax.GetValue(dict, "MinsOcrPrefix") != null;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private static string AddFontToResources(string resDict, string fontRef, Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
        {
            string? font = PdfSyntax.GetValue(resDict, "Font");
            string fontDict;
            if (font == null)
                fontDict = "<< >>";
            else if (PdfSyntax.TryParseReference(font, out _, out _))
                fontDict = ResolveValue(font, objects) ?? "<< >>"; // 간접 /Font는 이 Resources 안으로 복사
            else
                fontDict = font;

            if (!fontDict.StartsWith("<<", StringComparison.Ordinal))
                fontDict = "<< >>";

            fontDict = PdfSyntax.SetValue(fontDict, FontResName, fontRef);
            return PdfSyntax.SetValue(resDict, "Font", fontDict);
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
            byte[] prefix = Encoding.Latin1.GetBytes($"{number} {generation} obj\n{header}");
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
            var sb = new StringBuilder(s.Length * 4);
            foreach (char ch in s)
                sb.Append(((int)ch).ToString("X4"));
            return sb.ToString();
        }

        private static string BuildXrefAndTrailer(
            List<(int Number, int Generation, long Offset)> entries,
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
                    sb.Append($"{sorted[i + j].Offset:D10} {sorted[i + j].Generation:D5} n \n");

                i += count;
            }

            sb.Append("trailer\n<< ");
            sb.Append($"/Size {nextObjectNumber} ");
            sb.Append($"/Root {trailer.RootObjectNumber} {trailer.RootGeneration} R ");
            if (!string.IsNullOrEmpty(trailer.InfoRef))
                sb.Append($"/Info {trailer.InfoRef} ");
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
            if (previousXrefOffset < 0 || previousXrefOffset >= text.Length ||
                string.CompareOrdinal(text, (int)previousXrefOffset, "xref", 0, 4) != 0)
            {
                // 교차 참조 스트림(PDF 1.5+)은 이 단순 증분 기록기가 이어 쓸 수 없다.
                throw new InvalidOperationException("PDF uses a cross-reference stream; OCR text layer cannot be appended.");
            }

            int trailerIndex = text.IndexOf("trailer", (int)previousXrefOffset, StringComparison.Ordinal);
            if (trailerIndex < 0)
                throw new InvalidOperationException("PDF trailer was not found.");

            string trailerDictionary = PdfSyntax.ExtractLeadingDictionary(text.Substring(trailerIndex, Math.Min(64 * 1024, text.Length - trailerIndex)));
            if (!PdfSyntax.TryParseReference(PdfSyntax.GetValue(trailerDictionary, "Root"), out int rootNum, out int rootGen))
                throw new InvalidOperationException("Reference for /Root was not found.");
            string? sizeValue = PdfSyntax.GetValue(trailerDictionary, "Size");
            if (sizeValue == null || !int.TryParse(sizeValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int size))
                throw new InvalidOperationException("Integer value for /Size was not found.");

            string? infoValue = PdfSyntax.GetValue(trailerDictionary, "Info");
            string? infoRef = PdfSyntax.TryParseReference(infoValue, out _, out _) ? infoValue : null;

            return new TrailerInfo(
                size,
                rootNum,
                rootGen,
                previousXrefOffset,
                PdfSyntax.GetValue(trailerDictionary, "ID"),
                infoRef,
                PdfSyntax.GetValue(trailerDictionary, "Encrypt") != null);
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
            var result = new List<PageRef>();
            string catalogDict = PdfSyntax.ExtractLeadingDictionary(catalog);
            if (PdfSyntax.TryParseReference(PdfSyntax.GetValue(catalogDict, "Pages"), out int pagesNum, out int pagesGen))
            {
                var visited = new HashSet<(int, int)>();
                VisitPageTree(new PageRef(pagesNum, pagesGen), objects, visited, result);
            }

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

            string dict;
            try
            {
                dict = PdfSyntax.ExtractLeadingDictionary(obj.Content);
            }
            catch (FormatException)
            {
                return;
            }

            string? type = PdfSyntax.GetValue(dict, "Type");
            if (type == "/Page")
            {
                pages.Add(node);
                return;
            }

            string? kids = ResolveValue(PdfSyntax.GetValue(dict, "Kids"), objects);
            if (kids == null)
                return;

            foreach (var item in PdfSyntax.ParseArrayItems(kids))
            {
                if (PdfSyntax.TryParseReference(item, out int num, out int gen))
                    VisitPageTree(new PageRef(num, gen), objects, visited, pages);
            }
        }

        private sealed record TrailerInfo(int Size, int RootObjectNumber, int RootGeneration, long PreviousXrefOffset, string? IdRaw, string? InfoRef, bool IsEncrypted);
        internal sealed record PdfObjectInfo(int Number, int Generation, string Content);
        private sealed record PageRef(int ObjectNumber, int Generation);
    }
}
