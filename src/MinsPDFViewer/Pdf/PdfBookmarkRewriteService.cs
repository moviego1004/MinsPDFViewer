using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MinsPDFViewer
{
    internal sealed class PdfBookmarkRewriteService
    {
        private readonly Action<string>? _log;

        public PdfBookmarkRewriteService(Action<string>? log = null)
        {
            _log = log;
        }

        public async Task RewriteBookmarksAsync(string pdfPath, List<BookmarkSaveData> bookmarksSnapshot)
        {
            await Task.Run(() => RewriteBookmarks(pdfPath, bookmarksSnapshot));
        }

        private void RewriteBookmarks(string pdfPath, List<BookmarkSaveData> bookmarksSnapshot)
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

            var objectBuilder = new StringBuilder();
            var rewrittenObjects = new List<(int Number, long Offset, string Content)>();
            int nextObjectNumber = trailer.Size;

            int outlineRootObjectNumber = nextObjectNumber++;
            var outlineItems = new List<OutlineItem>();
            BuildOutlineItems(bookmarksSnapshot, outlineRootObjectNumber, ref nextObjectNumber, outlineItems, pageRefs);

            string catalog = RewriteCatalog(rootObject.Content, outlineRootObjectNumber);
            rewrittenObjects.Add((trailer.RootObjectNumber, 0, WrapObject(trailer.RootObjectNumber, trailer.RootGeneration, catalog)));

            string outlineRoot = BuildOutlineRoot(outlineItems);
            rewrittenObjects.Add((outlineRootObjectNumber, 0, WrapObject(outlineRootObjectNumber, 0, outlineRoot)));

            foreach (var item in outlineItems)
                rewrittenObjects.Add((item.ObjectNumber, 0, WrapObject(item.ObjectNumber, 0, BuildOutlineItemObject(item))));

            long appendStart = originalBytes.Length;
            long currentOffset = appendStart;
            for (int i = 0; i < rewrittenObjects.Count; i++)
            {
                var obj = rewrittenObjects[i];
                string normalized = EnsureLineEnd(obj.Content);
                rewrittenObjects[i] = (obj.Number, currentOffset, normalized);
                objectBuilder.Append(normalized);
                currentOffset += Encoding.Latin1.GetByteCount(normalized);
            }

            long xrefOffset = currentOffset;
            string xref = BuildXrefAndTrailer(rewrittenObjects, trailer, outlineItems, xrefOffset);
            objectBuilder.Append(xref);

            using var stream = new FileStream(pdfPath, FileMode.Append, FileAccess.Write, FileShare.None);
            byte[] appendBytes = Encoding.Latin1.GetBytes(objectBuilder.ToString());
            stream.Write(appendBytes, 0, appendBytes.Length);

            _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfiumWithBookmarkRewrite}] Bookmark rewrite successful. Count={bookmarksSnapshot.Count}");
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
                    if (IsPageObject(obj.Content))
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

            if (IsPageObject(obj.Content))
            {
                pages.Add(node);
                return;
            }

            foreach (var child in ReadReferencesFromArray(obj.Content, "Kids"))
                VisitPageTree(child, objects, visited, pages);
        }

        private static bool IsPageObject(string content)
        {
            return Regex.IsMatch(content, @"/Type\s*/Page\b") &&
                   !Regex.IsMatch(content, @"/Type\s*/Pages\b");
        }

        private static void BuildOutlineItems(
            List<BookmarkSaveData> bookmarks,
            int parentObjectNumber,
            ref int nextObjectNumber,
            List<OutlineItem> allItems,
            List<PageRef> pageRefs)
        {
            var siblings = new List<OutlineItem>();
            foreach (var bookmark in bookmarks)
            {
                int pageIndex = Math.Max(0, Math.Min(bookmark.PageIndex, pageRefs.Count - 1));
                var item = new OutlineItem
                {
                    ObjectNumber = nextObjectNumber++,
                    ParentObjectNumber = parentObjectNumber,
                    Title = bookmark.Title ?? string.Empty,
                    DestinationPage = pageRefs[pageIndex]
                };
                siblings.Add(item);
                allItems.Add(item);
            }

            for (int i = 0; i < siblings.Count; i++)
            {
                siblings[i].PreviousObjectNumber = i > 0 ? siblings[i - 1].ObjectNumber : 0;
                siblings[i].NextObjectNumber = i + 1 < siblings.Count ? siblings[i + 1].ObjectNumber : 0;
            }

            foreach (var pair in bookmarks.Zip(siblings, (bookmark, item) => (bookmark, item)))
            {
                var before = allItems.Count;
                BuildOutlineItems(pair.bookmark.Children, pair.item.ObjectNumber, ref nextObjectNumber, allItems, pageRefs);
                var children = allItems.Skip(before).Where(i => i.ParentObjectNumber == pair.item.ObjectNumber).ToList();
                if (children.Count > 0)
                {
                    pair.item.FirstChildObjectNumber = children.First().ObjectNumber;
                    pair.item.LastChildObjectNumber = children.Last().ObjectNumber;
                    pair.item.OpenDescendantCount = CountOpenDescendants(pair.item, allItems);
                }
            }
        }

        private static int CountOpenDescendants(OutlineItem item, List<OutlineItem> allItems)
        {
            int count = 0;
            foreach (var child in allItems.Where(i => i.ParentObjectNumber == item.ObjectNumber))
            {
                count++;
                count += CountOpenDescendants(child, allItems);
            }
            return count;
        }

        private static string BuildOutlineRoot(List<OutlineItem> outlineItems)
        {
            var topLevel = outlineItems.Where(i => i.ParentObjectNumber == outlineItems.FirstOrDefault()?.ParentObjectNumber).ToList();
            if (outlineItems.Count == 0)
                return "<< /Type /Outlines /Count 0 >>";

            int rootObjectNumber = outlineItems.First().ParentObjectNumber;
            topLevel = outlineItems.Where(i => i.ParentObjectNumber == rootObjectNumber).ToList();
            int visibleCount = topLevel.Count + topLevel.Sum(i => CountOpenDescendants(i, outlineItems));
            return $"<< /Type /Outlines /First {topLevel.First().ObjectNumber} 0 R /Last {topLevel.Last().ObjectNumber} 0 R /Count {visibleCount} >>";
        }

        private static string BuildOutlineItemObject(OutlineItem item)
        {
            var sb = new StringBuilder();
            sb.Append("<< ");
            sb.Append($"/Title {EncodePdfTextString(item.Title)} ");
            sb.Append($"/Parent {item.ParentObjectNumber} 0 R ");
            if (item.PreviousObjectNumber > 0) sb.Append($"/Prev {item.PreviousObjectNumber} 0 R ");
            if (item.NextObjectNumber > 0) sb.Append($"/Next {item.NextObjectNumber} 0 R ");
            if (item.FirstChildObjectNumber > 0) sb.Append($"/First {item.FirstChildObjectNumber} 0 R ");
            if (item.LastChildObjectNumber > 0) sb.Append($"/Last {item.LastChildObjectNumber} 0 R ");
            if (item.OpenDescendantCount > 0) sb.Append($"/Count {item.OpenDescendantCount} ");
            sb.Append($"/Dest [{item.DestinationPage.ObjectNumber} {item.DestinationPage.Generation} R /XYZ null null null] ");
            sb.Append(">>");
            return sb.ToString();
        }

        private static string RewriteCatalog(string catalogContent, int outlineRootObjectNumber)
        {
            string dictionary = ExtractDictionary(catalogContent, catalogContent.IndexOf("<<", StringComparison.Ordinal));
            dictionary = Regex.Replace(dictionary, @"/Outlines\s+\d+\s+\d+\s+R", "");
            dictionary = Regex.Replace(dictionary, @"/PageMode\s*/\w+", "");
            int insert = dictionary.LastIndexOf(">>", StringComparison.Ordinal);
            if (insert < 0)
                throw new InvalidOperationException("PDF catalog dictionary could not be parsed.");

            return dictionary.Insert(insert, $" /Outlines {outlineRootObjectNumber} 0 R /PageMode /UseOutlines ");
        }

        private static string BuildXrefAndTrailer(
            List<(int Number, long Offset, string Content)> objects,
            TrailerInfo trailer,
            List<OutlineItem> outlineItems,
            long xrefOffset)
        {
            var ordered = objects.OrderBy(o => o.Number).ToList();
            var sb = new StringBuilder();
            sb.Append("xref\n");
            foreach (var group in GroupConsecutive(ordered))
            {
                sb.Append($"{group.First().Number} {group.Count}\n");
                foreach (var obj in group)
                    sb.Append($"{obj.Offset:0000000000} 00000 n \n");
            }

            int newSize = Math.Max(trailer.Size, ordered.Max(o => o.Number) + 1);
            sb.Append("trailer\n");
            sb.Append("<< ");
            sb.Append($"/Size {newSize} ");
            sb.Append($"/Root {trailer.RootObjectNumber} {trailer.RootGeneration} R ");
            if (!string.IsNullOrWhiteSpace(trailer.IdArray))
                sb.Append($"/ID {trailer.IdArray} ");
            sb.Append($"/Prev {trailer.PreviousXrefOffset} ");
            sb.Append(">>\n");
            sb.Append("startxref\n");
            sb.Append($"{xrefOffset}\n");
            sb.Append("%%EOF\n");
            return sb.ToString();
        }

        private static List<List<(int Number, long Offset, string Content)>> GroupConsecutive(List<(int Number, long Offset, string Content)> objects)
        {
            var groups = new List<List<(int, long, string)>>();
            foreach (var obj in objects)
            {
                if (groups.Count == 0 || groups[^1][^1].Item1 + 1 != obj.Number)
                    groups.Add(new List<(int, long, string)>());
                groups[^1].Add(obj);
            }
            return groups;
        }

        private static string WrapObject(int number, int generation, string content)
        {
            return $"{number} {generation} obj\n{content}\nendobj\n";
        }

        private static string EnsureLineEnd(string value)
        {
            return value.EndsWith("\n", StringComparison.Ordinal) ? value : value + "\n";
        }

        private static string ExtractDictionary(string text, int dictionaryStart)
        {
            if (dictionaryStart < 0 || dictionaryStart + 1 >= text.Length)
                throw new InvalidOperationException("PDF dictionary start was not found.");

            int depth = 0;
            for (int i = dictionaryStart; i < text.Length - 1; i++)
            {
                if (text[i] == '<' && text[i + 1] == '<')
                {
                    depth++;
                    i++;
                    continue;
                }

                if (text[i] == '>' && text[i + 1] == '>')
                {
                    depth--;
                    i++;
                    if (depth == 0)
                        return text.Substring(dictionaryStart, i - dictionaryStart + 1);
                }
            }

            throw new InvalidOperationException("PDF dictionary end was not found.");
        }

        private static PageRef ReadReference(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s+(\d+)\s+(\d+)\s+R");
            if (!match.Success)
                throw new InvalidOperationException($"PDF reference /{key} was not found.");

            return new PageRef(
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        }

        private static int ReadInteger(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s+(\d+)");
            if (!match.Success)
                throw new InvalidOperationException($"PDF integer /{key} was not found.");

            return int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        private static string? ReadRawArray(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s*(\[[^\]]+\])", RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static List<PageRef> ReadReferencesFromArray(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s*\[(.*?)\]", RegexOptions.Singleline);
            if (!match.Success)
                return new List<PageRef>();

            return Regex.Matches(match.Groups[1].Value, @"(\d+)\s+(\d+)\s+R")
                .Select(m => new PageRef(
                    int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture),
                    int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)))
                .ToList();
        }

        private static string EncodePdfTextString(string value)
        {
            byte[] unicode = Encoding.BigEndianUnicode.GetBytes(value ?? string.Empty);
            var bytes = new byte[unicode.Length + 2];
            bytes[0] = 0xFE;
            bytes[1] = 0xFF;
            Buffer.BlockCopy(unicode, 0, bytes, 2, unicode.Length);
            return "<" + Convert.ToHexString(bytes) + ">";
        }

        private sealed record TrailerInfo(int Size, int RootObjectNumber, int RootGeneration, long PreviousXrefOffset, string? IdArray);
        private sealed record PdfObjectInfo(int Number, int Generation, string Content);
        private sealed record PageRef(int ObjectNumber, int Generation);

        private sealed class OutlineItem
        {
            public int ObjectNumber { get; set; }
            public int ParentObjectNumber { get; set; }
            public int PreviousObjectNumber { get; set; }
            public int NextObjectNumber { get; set; }
            public int FirstChildObjectNumber { get; set; }
            public int LastChildObjectNumber { get; set; }
            public int OpenDescendantCount { get; set; }
            public string Title { get; set; } = string.Empty;
            public PageRef DestinationPage { get; set; } = new(0, 0);
        }
    }
}
