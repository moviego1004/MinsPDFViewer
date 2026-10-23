using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace MinsPDFViewer
{
    /// <summary>간접 객체 하나. Content = obj와 endobj 사이의 값 (스트림이면 사전만).</summary>
    internal sealed record PdfObjectInfo(int Number, int Generation, string Content);

    /// <summary>
    /// 증분 저장(북마크·서명·OCR 텍스트 층)과 서명 검증을 위한 PDF 객체 읽기.
    /// 교차 참조 표(xref)와 교차 참조 스트림(PDF 1.5+, 객체 스트림 포함), /Prev 사슬, 혼합형(/XRefStm), 선형화 파일을 읽는다.
    /// 파일은 필요한 부분만 읽는다 (1GB 파일도 통째로 메모리에 올리지 않음).
    /// 암호화 문서는 읽기만 하고(문자열·스트림은 암호화된 그대로) 기록기는 거부한다.
    /// </summary>
    internal sealed class PdfObjectStore : IDisposable
    {
        private readonly FileStream _file;
        private readonly Dictionary<int, XrefEntry> _xref = new();
        private readonly Dictionary<(int, int), PdfObjectInfo> _overrides = new();
        private readonly Dictionary<(int, int), PdfObjectInfo?> _cache = new();
        private readonly Dictionary<int, ObjectStream> _objectStreams = new();

        private readonly record struct XrefEntry(int Type, long Offset, int Generation);

        private sealed class ObjectStream
        {
            public string Text = "";
            public int First;
            public readonly Dictionary<int, (int Start, int End)> Objects = new();
        }

        public string FilePath { get; }
        public long FileLength { get; }
        /// <summary>trailer /Size (모든 xref 구역 중 최대, 가장 큰 객체 번호 + 1 이상).</summary>
        public int Size { get; private set; }
        public int RootNumber { get; private set; }
        public int RootGeneration { get; private set; }
        public string? InfoRef { get; private set; }
        public string? IdRaw { get; private set; }
        public bool IsEncrypted { get; private set; }
        /// <summary>마지막(가장 최근) xref 위치 = 새 구역의 /Prev.</summary>
        public long LastXrefOffset { get; private set; }
        /// <summary>마지막 xref가 교차 참조 스트림인지. 그러면 새 구역도 스트림으로 쓴다.</summary>
        public bool LastXrefIsStream { get; private set; }
        public bool EndsWithNewline { get; private set; }

        private PdfObjectStore(FileStream file, string path)
        {
            _file = file;
            FilePath = path;
            FileLength = file.Length;
        }

        public static PdfObjectStore Open(string path)
        {
            var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 64 * 1024, FileOptions.RandomAccess);
            try
            {
                var store = new PdfObjectStore(file, path);
                store.Load();
                return store;
            }
            catch
            {
                file.Dispose();
                throw;
            }
        }

        public void Dispose() => _file.Dispose();

        // ── 원시 읽기 ────────────────────────────────────────────────────

        private byte[] ReadBytes(long position, int count)
        {
            if (position < 0 || position >= FileLength)
                return Array.Empty<byte>();
            count = (int)Math.Min(count, FileLength - position);
            var buffer = new byte[count];
            _file.Position = position;
            int done = 0;
            while (done < count)
            {
                int read = _file.Read(buffer, done, count - done);
                if (read <= 0)
                    break;
                done += read;
            }
            if (done < count)
                Array.Resize(ref buffer, done);
            return buffer;
        }

        private string ReadText(long position, int count) => Encoding.Latin1.GetString(ReadBytes(position, count));

        // ── xref 사슬 ────────────────────────────────────────────────────

        private void Load()
        {
            if (FileLength < 16)
                throw new InvalidOperationException("PDF 파일이 너무 짧습니다.");

            var tail = ReadBytes(Math.Max(0, FileLength - 2048), 2048);
            EndsWithNewline = tail.Length > 0 && (tail[^1] == (byte)'\n' || tail[^1] == (byte)'\r');
            string tailText = Encoding.Latin1.GetString(tail);
            int sx = tailText.LastIndexOf("startxref", StringComparison.Ordinal);
            if (sx < 0)
                throw new InvalidOperationException("PDF startxref를 찾을 수 없습니다.");
            int p = PdfSyntax.SkipWhite(tailText, sx + 9);
            int e = p;
            while (e < tailText.Length && char.IsDigit(tailText[e]))
                e++;
            if (e == p || !long.TryParse(tailText.AsSpan(p, e - p), NumberStyles.None, CultureInfo.InvariantCulture, out long startXref))
                throw new InvalidOperationException("PDF startxref 값을 읽을 수 없습니다.");

            LastXrefOffset = startXref;
            var visited = new HashSet<long>();
            long? offset = startXref;
            bool first = true;
            int maxSize = 0;
            while (offset.HasValue && visited.Add(offset.Value))
            {
                string trailer = ReadXrefSection(offset.Value, out bool isStream);
                if (first)
                {
                    LastXrefIsStream = isStream;
                    first = false;
                }
                ApplyTrailer(trailer);
                if (int.TryParse(PdfSyntax.GetValue(trailer, "Size"), NumberStyles.None, CultureInfo.InvariantCulture, out int size))
                    maxSize = Math.Max(maxSize, size);

                // 혼합형: 표 구역에 딸린 xref 스트림 (표에 없는 객체를 이전 구역보다 먼저 찾는다)
                if (!isStream && long.TryParse(PdfSyntax.GetValue(trailer, "XRefStm"), NumberStyles.None, CultureInfo.InvariantCulture, out long xrefStm) &&
                    visited.Add(xrefStm))
                {
                    ReadXrefSection(xrefStm, out _);
                }

                offset = long.TryParse(PdfSyntax.GetValue(trailer, "Prev"), NumberStyles.None, CultureInfo.InvariantCulture, out long prev)
                    ? prev
                    : null;
            }

            if (RootNumber <= 0)
                throw new InvalidOperationException("PDF trailer에 /Root가 없습니다.");
            int maxNumber = _xref.Count > 0 ? _xref.Keys.Max() : 0;
            Size = Math.Max(maxSize, maxNumber + 1);
        }

        /// <summary>새 구역의 값이 먼저 들어오므로, 아직 없는 값만 채운다.</summary>
        private void ApplyTrailer(string trailer)
        {
            if (RootNumber <= 0 && PdfSyntax.TryParseReference(PdfSyntax.GetValue(trailer, "Root"), out int rn, out int rg))
            {
                RootNumber = rn;
                RootGeneration = rg;
            }
            if (InfoRef == null)
            {
                string? info = PdfSyntax.GetValue(trailer, "Info");
                if (PdfSyntax.TryParseReference(info, out _, out _))
                    InfoRef = info;
            }
            IdRaw ??= PdfSyntax.GetValue(trailer, "ID");
            if (PdfSyntax.GetValue(trailer, "Encrypt") != null)
                IsEncrypted = true;
        }

        private void AddEntry(int number, XrefEntry entry)
        {
            if (number <= 0 || _xref.ContainsKey(number))
                return; // 0번은 늘 free, 이미 있으면 더 새 구역의 값
            _xref[number] = entry;
        }

        /// <summary>offset의 xref 구역을 읽어 항목을 더하고 trailer 사전을 돌려준다.</summary>
        private string ReadXrefSection(long offset, out bool isStream)
        {
            string head = ReadText(offset, 64);
            int i = PdfSyntax.SkipWhite(head, 0);
            if (string.CompareOrdinal(head, i, "xref", 0, 4) == 0)
            {
                isStream = false;
                return ReadXrefTable(offset + i + 4);
            }

            isStream = true;
            return ReadXrefStream(offset);
        }

        private string ReadXrefTable(long position)
        {
            var reader = new TokenReader(this, position);
            while (true)
            {
                string? token = reader.Next();
                if (token == null)
                    throw new InvalidOperationException("PDF xref 표가 끝나지 않습니다.");
                if (token == "trailer")
                    break;
                if (!int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out int start) ||
                    !int.TryParse(reader.Next(), NumberStyles.None, CultureInfo.InvariantCulture, out int count))
                {
                    throw new InvalidOperationException("PDF xref 표를 읽을 수 없습니다.");
                }

                for (int k = 0; k < count; k++)
                {
                    string? off = reader.Next();
                    string? gen = reader.Next();
                    string? type = reader.Next();
                    if (off == null || gen == null || type == null)
                        throw new InvalidOperationException("PDF xref 항목이 잘렸습니다.");
                    long.TryParse(off, NumberStyles.None, CultureInfo.InvariantCulture, out long o);
                    int.TryParse(gen, NumberStyles.None, CultureInfo.InvariantCulture, out int g);
                    AddEntry(start + k, type == "n" ? new XrefEntry(1, o, g) : new XrefEntry(0, 0, g));
                }
            }

            return ReadDictionaryAt(reader.Position);
        }

        private string ReadDictionaryAt(long position)
        {
            for (int size = 4096; ; size *= 4)
            {
                string text = ReadText(position, size);
                try
                {
                    return PdfSyntax.ExtractLeadingDictionary(text);
                }
                catch (FormatException) when (text.Length == size && size < 64 * 1024 * 1024)
                {
                }
            }
        }

        private string ReadXrefStream(long offset)
        {
            var obj = ReadIndirectAt(offset, out int number, out _);
            if (obj.StreamStart < 0)
                throw new InvalidOperationException("PDF xref 위치에 xref 표도 스트림도 없습니다.");
            string dict = obj.Value;
            if (PdfSyntax.GetValue(dict, "Type") != "/XRef")
                throw new InvalidOperationException("PDF xref 스트림이 아닙니다.");

            if (!int.TryParse(PdfSyntax.GetValue(dict, "Length"), NumberStyles.None, CultureInfo.InvariantCulture, out int length))
                throw new InvalidOperationException("PDF xref 스트림 /Length가 직접 값이 아닙니다.");
            byte[] data = DecodeStream(dict, ReadBytes(obj.StreamStart, length));

            var w = PdfSyntax.ParseArrayItems(PdfSyntax.GetValue(dict, "W") ?? "")
                .Select(s => int.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            if (w.Length != 3)
                throw new InvalidOperationException("PDF xref 스트림 /W가 잘못됐습니다.");
            int size = int.Parse(PdfSyntax.GetValue(dict, "Size") ?? "0", CultureInfo.InvariantCulture);
            var index = PdfSyntax.ParseArrayItems(PdfSyntax.GetValue(dict, "Index") ?? $"[0 {size}]")
                .Select(s => int.Parse(s, CultureInfo.InvariantCulture)).ToArray();

            int entryLength = w[0] + w[1] + w[2];
            int pos = 0;
            for (int s = 0; s + 1 < index.Length; s += 2)
            {
                for (int k = 0; k < index[s + 1]; k++)
                {
                    if (pos + entryLength > data.Length)
                        break;
                    long type = w[0] == 0 ? 1 : ReadField(data, pos, w[0]);
                    long f2 = ReadField(data, pos + w[0], w[1]);
                    long f3 = ReadField(data, pos + w[0] + w[1], w[2]);
                    pos += entryLength;
                    int objectNumber = index[s] + k;
                    switch (type)
                    {
                        case 1: AddEntry(objectNumber, new XrefEntry(1, f2, (int)f3)); break;
                        case 2: AddEntry(objectNumber, new XrefEntry(2, f2, (int)f3)); break;
                        case 0: AddEntry(objectNumber, new XrefEntry(0, 0, (int)f3)); break;
                    }
                }
            }

            // xref 스트림 객체 자신
            AddEntry(number, new XrefEntry(1, offset, 0));
            return dict;
        }

        private static long ReadField(byte[] data, int pos, int width)
        {
            long v = 0;
            for (int i = 0; i < width; i++)
                v = (v << 8) | data[pos + i];
            return v;
        }

        // ── 객체 읽기 ────────────────────────────────────────────────────

        private readonly record struct RawObject(string Value, long StreamStart);

        /// <summary>파일 위치의 "N G obj 값 [stream]"을 읽는다. 스트림이면 StreamStart = 데이터 시작.</summary>
        private RawObject ReadIndirectAt(long offset, out int number, out int generation)
        {
            for (int size = 16 * 1024; ; size *= 4)
            {
                string text = ReadText(offset, size);
                bool complete = text.Length < size;
                try
                {
                    int i = PdfSyntax.SkipWhite(text, 0);
                    int e = ReadDigits(text, i);
                    number = int.Parse(text.AsSpan(i, e - i), NumberStyles.None, CultureInfo.InvariantCulture);
                    i = PdfSyntax.SkipWhite(text, e);
                    e = ReadDigits(text, i);
                    generation = int.Parse(text.AsSpan(i, e - i), NumberStyles.None, CultureInfo.InvariantCulture);
                    i = PdfSyntax.SkipWhite(text, e);
                    if (string.CompareOrdinal(text, i, "obj", 0, 3) != 0)
                        throw new InvalidOperationException($"PDF 객체 머리를 찾을 수 없습니다 (위치 {offset}).");
                    int valueStart = PdfSyntax.SkipWhite(text, i + 3);
                    int valueEnd = PdfSyntax.SkipObject(text, valueStart);
                    if (valueEnd >= text.Length && !complete)
                        throw new FormatException("more");
                    string value = text.Substring(valueStart, valueEnd - valueStart).Trim();

                    int after = PdfSyntax.SkipWhite(text, valueEnd);
                    long streamStart = -1;
                    if (string.CompareOrdinal(text, after, "stream", 0, 6) == 0)
                    {
                        int d = after + 6;
                        if (d < text.Length && text[d] == '\r')
                            d++;
                        if (d < text.Length && text[d] == '\n')
                            d++;
                        streamStart = offset + d;
                    }
                    return new RawObject(value, streamStart);
                }
                catch (Exception ex) when ((ex is FormatException || ex is ArgumentOutOfRangeException) && !complete && size < 256 * 1024 * 1024)
                {
                    // 값이 읽은 범위보다 길다 → 더 읽는다
                }
            }
        }

        private static int ReadDigits(string text, int i)
        {
            int e = i;
            while (e < text.Length && char.IsDigit(text[e]))
                e++;
            if (e == i)
                throw new FormatException("number expected");
            return e;
        }

        /// <summary>기록기가 고친 객체 (이후 읽기에 반영).</summary>
        public PdfObjectInfo this[(int Number, int Generation) key]
        {
            set => _overrides[key] = value;
        }

        public bool TryGetValue((int Number, int Generation) key, out PdfObjectInfo obj)
        {
            if (_overrides.TryGetValue(key, out obj!))
                return true;
            obj = GetObject(key.Number, key.Generation)!;
            return obj != null;
        }

        public PdfObjectInfo? GetObject(int number, int generation)
        {
            if (_cache.TryGetValue((number, generation), out var cached))
                return cached;

            PdfObjectInfo? result = null;
            if (_xref.TryGetValue(number, out var entry))
            {
                try
                {
                    if (entry.Type == 1 && entry.Generation == generation)
                    {
                        var raw = ReadIndirectAt(entry.Offset, out int n, out _);
                        if (n == number)
                            result = new PdfObjectInfo(number, generation, raw.Value);
                    }
                    else if (entry.Type == 2 && generation == 0)
                    {
                        result = ReadFromObjectStream(number, (int)entry.Offset);
                    }
                }
                catch (Exception)
                {
                    result = null;
                }
            }

            _cache[(number, generation)] = result;
            return result;
        }

        private PdfObjectInfo? ReadFromObjectStream(int number, int streamNumber)
        {
            if (!_objectStreams.TryGetValue(streamNumber, out var os))
            {
                byte[]? data = GetStreamData(streamNumber, 0);
                if (data == null)
                    return null;
                os = new ObjectStream { Text = Encoding.Latin1.GetString(data) };
                var dictObj = GetObject(streamNumber, 0);
                string dict = dictObj?.Content ?? "<< >>";
                int count = int.Parse(PdfSyntax.GetValue(dict, "N") ?? "0", CultureInfo.InvariantCulture);
                os.First = int.Parse(PdfSyntax.GetValue(dict, "First") ?? "0", CultureInfo.InvariantCulture);
                var header = new TokenReader(os.Text.Substring(0, Math.Min(os.First, os.Text.Length)));
                var pairs = new List<(int Number, int Offset)>();
                for (int k = 0; k < count; k++)
                {
                    if (!int.TryParse(header.Next(), out int objNum) || !int.TryParse(header.Next(), out int objOff))
                        break;
                    pairs.Add((objNum, objOff));
                }
                for (int k = 0; k < pairs.Count; k++)
                {
                    int start = os.First + pairs[k].Offset;
                    int end = k + 1 < pairs.Count ? os.First + pairs[k + 1].Offset : os.Text.Length;
                    os.Objects[pairs[k].Number] = (start, Math.Min(end, os.Text.Length));
                }
                _objectStreams[streamNumber] = os;
            }

            if (!os.Objects.TryGetValue(number, out var range) || range.Start >= range.End)
                return null;
            string value = os.Text.Substring(range.Start, range.End - range.Start);
            int vs = PdfSyntax.SkipWhite(value, 0);
            int ve = PdfSyntax.SkipObject(value, vs);
            return new PdfObjectInfo(number, 0, value.Substring(vs, ve - vs).Trim());
        }

        /// <summary>스트림 객체의 데이터 (FlateDecode·PNG 예측자 풀어서). 다른 필터면 null.</summary>
        public byte[]? GetStreamData(int number, int generation)
        {
            if (!_xref.TryGetValue(number, out var entry) || entry.Type != 1)
                return null;
            var raw = ReadIndirectAt(entry.Offset, out _, out _);
            if (raw.StreamStart < 0)
                return null;
            string? lengthValue = ResolveValue(PdfSyntax.GetValue(raw.Value, "Length"));
            if (!int.TryParse(lengthValue, NumberStyles.None, CultureInfo.InvariantCulture, out int length))
                return null;
            try
            {
                return DecodeStream(raw.Value, ReadBytes(raw.StreamStart, length));
            }
            catch (NotSupportedException)
            {
                return null;
            }
        }

        private static byte[] DecodeStream(string dict, byte[] data)
        {
            string? filter = PdfSyntax.GetValue(dict, "Filter");
            if (filter == null)
                return data;
            var filters = filter.StartsWith("[", StringComparison.Ordinal) ? PdfSyntax.ParseArrayItems(filter) : new List<string> { filter };
            string? parms = PdfSyntax.GetValue(dict, "DecodeParms");
            if (parms != null && parms.StartsWith("[", StringComparison.Ordinal))
                parms = PdfSyntax.ParseArrayItems(parms).FirstOrDefault(p => p.StartsWith("<<", StringComparison.Ordinal));

            foreach (var f in filters)
            {
                if (f.Trim() != "/FlateDecode")
                    throw new NotSupportedException($"PDF 필터 {f}는 지원하지 않습니다.");
                data = Inflate(data);
                if (parms != null && parms.StartsWith("<<", StringComparison.Ordinal))
                    data = ApplyPredictor(data, parms);
            }
            return data;
        }

        private static byte[] Inflate(byte[] data)
        {
            try
            {
                using var input = new MemoryStream(data);
                using var z = new ZLibStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                z.CopyTo(output);
                return output.ToArray();
            }
            catch (InvalidDataException)
            {
                // zlib 머리가 깨진 파일: 원시 deflate로
                using var input = new MemoryStream(data, 2, Math.Max(0, data.Length - 2));
                using var d = new DeflateStream(input, CompressionMode.Decompress);
                using var output = new MemoryStream();
                d.CopyTo(output);
                return output.ToArray();
            }
        }

        private static byte[] ApplyPredictor(byte[] data, string parms)
        {
            int predictor = int.TryParse(PdfSyntax.GetValue(parms, "Predictor"), out int pr) ? pr : 1;
            if (predictor < 10)
            {
                if (predictor == 1)
                    return data;
                throw new NotSupportedException("TIFF 예측자는 지원하지 않습니다.");
            }
            int columns = int.TryParse(PdfSyntax.GetValue(parms, "Columns"), out int c) ? c : 1;
            int colors = int.TryParse(PdfSyntax.GetValue(parms, "Colors"), out int cl) ? cl : 1;
            int bpc = int.TryParse(PdfSyntax.GetValue(parms, "BitsPerComponent"), out int b) ? b : 8;
            int bpp = Math.Max(1, (colors * bpc + 7) / 8);
            int rowLength = (columns * colors * bpc + 7) / 8;
            int rows = data.Length / (rowLength + 1);
            var output = new byte[rows * rowLength];
            var previous = new byte[rowLength];
            for (int r = 0; r < rows; r++)
            {
                int filterType = data[r * (rowLength + 1)];
                int src = r * (rowLength + 1) + 1;
                int dst = r * rowLength;
                for (int x = 0; x < rowLength; x++)
                {
                    int raw = data[src + x];
                    int left = x >= bpp ? output[dst + x - bpp] : 0;
                    int up = previous[x];
                    int upLeft = x >= bpp ? previous[x - bpp] : 0;
                    int value = filterType switch
                    {
                        1 => raw + left,
                        2 => raw + up,
                        3 => raw + ((left + up) >> 1),
                        4 => raw + Paeth(left, up, upLeft),
                        _ => raw
                    };
                    output[dst + x] = (byte)value;
                }
                Array.Copy(output, dst, previous, 0, rowLength);
            }
            return output;
        }

        private static int Paeth(int a, int b, int c)
        {
            int p = a + b - c;
            int pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
            return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
        }

        // ── 구조 도우미 ──────────────────────────────────────────────────

        public string RootDictionary =>
            GetObject(RootNumber, RootGeneration)?.Content is { } c
                ? PdfSyntax.ExtractLeadingDictionary(c)
                : throw new InvalidOperationException("PDF 카탈로그를 읽을 수 없습니다.");

        /// <summary>간접 참조면 대상 객체 값(사전·배열·숫자)을, 아니면 값 그대로.</summary>
        public string? ResolveValue(string? value)
        {
            if (value == null)
                return null;
            if (!PdfSyntax.TryParseReference(value, out int num, out int gen))
                return value;
            return TryGetValue((num, gen), out var obj) ? obj.Content : null;
        }

        /// <summary>페이지 사전에서 상속 가능한 키(MediaBox, CropBox, Rotate, Resources)를 /Parent를 따라 찾는다.</summary>
        public string? GetInheritedValue(string pageDict, string key)
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
                if (!TryGetValue((num, gen), out var parentObj))
                    return null;
                current = PdfSyntax.ExtractLeadingDictionary(parentObj.Content);
            }
            return null;
        }

        /// <summary>페이지 순서대로 페이지 객체 참조.</summary>
        public List<(int Number, int Generation)> GetPageRefs()
        {
            var pages = new List<(int, int)>();
            if (PdfSyntax.TryParseReference(PdfSyntax.GetValue(RootDictionary, "Pages"), out int num, out int gen))
                VisitPageTree(num, gen, new HashSet<(int, int)>(), pages, 0);
            return pages;
        }

        private void VisitPageTree(int number, int generation, HashSet<(int, int)> visited, List<(int, int)> pages, int depth)
        {
            if (depth > 64 || !visited.Add((number, generation)) || !TryGetValue((number, generation), out var obj))
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

            string? kids = ResolveValue(PdfSyntax.GetValue(dict, "Kids"));
            if (PdfSyntax.GetValue(dict, "Type") == "/Page" || kids == null)
            {
                pages.Add((number, generation));
                return;
            }
            foreach (var item in PdfSyntax.ParseArrayItems(kids))
            {
                if (PdfSyntax.TryParseReference(item, out int n, out int g))
                    VisitPageTree(n, g, visited, pages, depth + 1);
            }
        }

        /// <summary>파일에 있는 객체 번호(최신 xref 기준, free 제외).</summary>
        public IEnumerable<int> ObjectNumbers => _xref.Where(kv => kv.Value.Type != 0).Select(kv => kv.Key);

        /// <summary>객체가 파일에서 시작하는 위치 (객체 스트림 안이면 -1).</summary>
        public long GetObjectOffset(int number) =>
            _xref.TryGetValue(number, out var e) && e.Type == 1 ? e.Offset : -1;

        // ── 토큰 읽기 (xref 표, 객체 스트림 머리) ───────────────────────

        private sealed class TokenReader
        {
            private readonly PdfObjectStore? _store;
            private readonly string? _text;
            private long _position;
            private byte[] _buffer = Array.Empty<byte>();
            private long _bufferStart;

            public TokenReader(PdfObjectStore store, long position)
            {
                _store = store;
                _position = position;
            }

            public TokenReader(string text)
            {
                _text = text;
            }

            public long Position => _position;

            private int ByteAt(long position)
            {
                if (_text != null)
                    return position < _text.Length ? _text[(int)position] : -1;
                if (position < _bufferStart || position >= _bufferStart + _buffer.Length)
                {
                    _bufferStart = position;
                    _buffer = _store!.ReadBytes(position, 64 * 1024);
                    if (_buffer.Length == 0)
                        return -1;
                }
                return _buffer[position - _bufferStart];
            }

            public string? Next()
            {
                int c;
                while ((c = ByteAt(_position)) >= 0 && PdfSyntax.IsWhite((char)c))
                    _position++;
                if (c < 0)
                    return null;
                var sb = new StringBuilder();
                while ((c = ByteAt(_position)) >= 0 && !PdfSyntax.IsWhite((char)c))
                {
                    sb.Append((char)c);
                    _position++;
                }
                return sb.ToString();
            }
        }
    }

    /// <summary>증분 갱신 구역(새 객체 + xref + trailer)을 만든다. 원본 xref가 스트림이면 스트림으로, 표면 표로.</summary>
    internal static class PdfIncrementalWriter
    {
        public static byte[] WrapObject(int number, int generation, string content) =>
            Encoding.Latin1.GetBytes($"{number} {generation} obj\n{content}\nendobj\n");

        public static byte[] WrapStream(int number, int generation, string dictionary, byte[] data)
        {
            byte[] head = Encoding.Latin1.GetBytes($"{number} {generation} obj\n{dictionary}\nstream\n");
            byte[] tail = Encoding.Latin1.GetBytes("\nendstream\nendobj\n");
            var result = new byte[head.Length + data.Length + tail.Length];
            Buffer.BlockCopy(head, 0, result, 0, head.Length);
            Buffer.BlockCopy(data, 0, result, head.Length, data.Length);
            Buffer.BlockCopy(tail, 0, result, head.Length + data.Length, tail.Length);
            return result;
        }

        /// <summary>
        /// 원본 끝에 붙일 바이트. offsets = 객체 번호 → 파일 안 절대 위치.
        /// </summary>
        public static byte[] BuildAppend(PdfObjectStore store, IReadOnlyList<(int Number, int Generation, byte[] Data)> objects,
            out Dictionary<int, long> offsets)
        {
            if (store.IsEncrypted)
                throw new InvalidOperationException("암호화된 PDF에는 덧붙여 저장할 수 없습니다.");

            offsets = new Dictionary<int, long>();
            using var output = new MemoryStream();
            long baseOffset = store.FileLength;
            if (!store.EndsWithNewline)
                output.WriteByte((byte)'\n');

            var entries = new List<(int Number, int Generation, long Offset)>();
            foreach (var obj in objects)
            {
                long offset = baseOffset + output.Position;
                offsets[obj.Number] = offset;
                entries.Add((obj.Number, obj.Generation, offset));
                output.Write(obj.Data, 0, obj.Data.Length);
            }

            int size = Math.Max(store.Size, objects.Count == 0 ? 0 : objects.Max(o => o.Number) + 1);
            var trailer = new StringBuilder();
            trailer.Append($"/Root {store.RootNumber} {store.RootGeneration} R ");
            if (store.InfoRef != null)
                trailer.Append($"/Info {store.InfoRef} ");
            if (store.IdRaw != null)
                trailer.Append($"/ID {store.IdRaw} ");
            trailer.Append($"/Prev {store.LastXrefOffset} ");

            if (store.LastXrefIsStream)
            {
                int xrefNumber = size;
                long xrefOffset = baseOffset + output.Position;
                entries.Add((xrefNumber, 0, xrefOffset));
                size = xrefNumber + 1;

                var sorted = entries.OrderBy(e => e.Number).ToList();
                var data = new MemoryStream();
                var index = new StringBuilder();
                foreach (var group in GroupConsecutive(sorted))
                {
                    index.Append(group[0].Number).Append(' ').Append(group.Count).Append(' ');
                    foreach (var e in group)
                    {
                        data.WriteByte(1);
                        for (int shift = 32; shift >= 0; shift -= 8)
                            data.WriteByte((byte)(e.Offset >> shift));
                        data.WriteByte((byte)(e.Generation >> 8));
                        data.WriteByte((byte)e.Generation);
                    }
                }
                byte[] xrefData = data.ToArray();
                string dict = $"<< /Type /XRef /Size {size} /Index [{index.ToString().Trim()}] /W [1 5 2] {trailer}/Length {xrefData.Length} >>";
                byte[] xrefObject = WrapStream(xrefNumber, 0, dict, xrefData);
                output.Write(xrefObject, 0, xrefObject.Length);
                byte[] end = Encoding.ASCII.GetBytes($"startxref\n{xrefOffset}\n%%EOF\n");
                output.Write(end, 0, end.Length);
            }
            else
            {
                long xrefOffset = baseOffset + output.Position;
                var sb = new StringBuilder("xref\n");
                foreach (var group in GroupConsecutive(entries.OrderBy(e => e.Number).ToList()))
                {
                    sb.Append(group[0].Number).Append(' ').Append(group.Count).Append('\n');
                    foreach (var e in group)
                        sb.Append($"{e.Offset:D10} {e.Generation:D5} n \n");
                }
                sb.Append($"trailer\n<< /Size {size} {trailer}>>\nstartxref\n{xrefOffset}\n%%EOF\n");
                byte[] xref = Encoding.ASCII.GetBytes(sb.ToString());
                output.Write(xref, 0, xref.Length);
            }

            return output.ToArray();
        }

        private static List<List<(int Number, int Generation, long Offset)>> GroupConsecutive(List<(int Number, int Generation, long Offset)> sorted)
        {
            var groups = new List<List<(int, int, long)>>();
            foreach (var e in sorted)
            {
                if (groups.Count == 0 || groups[^1][^1].Item1 + 1 != e.Number)
                    groups.Add(new List<(int, int, long)>());
                groups[^1].Add(e);
            }
            return groups;
        }

        public static void AppendToFile(string path, byte[] append)
        {
            using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.None);
            stream.Write(append, 0, append.Length);
            stream.Flush(true);
        }
    }
}
