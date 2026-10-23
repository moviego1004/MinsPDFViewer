using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
            byte[] append;
            using (var store = PdfObjectStore.Open(pdfPath))
            {
                if (store.IsEncrypted)
                    throw new InvalidOperationException("암호화된 PDF에는 북마크를 저장할 수 없습니다.");

                var pageRefs = store.GetPageRefs().Select(r => new PageRef(r.Number, r.Generation)).ToList();
                if (pageRefs.Count == 0)
                    throw new InvalidOperationException("PDF page tree could not be parsed.");

                int nextObjectNumber = store.Size;
                int outlineRootObjectNumber = nextObjectNumber++;
                var outlineItems = new List<OutlineItem>();
                BuildOutlineItems(bookmarksSnapshot, outlineRootObjectNumber, ref nextObjectNumber, outlineItems, pageRefs);

                string catalog = PdfSyntax.SetValue(store.RootDictionary, "Outlines", $"{outlineRootObjectNumber} 0 R");
                catalog = PdfSyntax.SetValue(catalog, "PageMode", outlineItems.Count > 0 ? "/UseOutlines" : "/UseNone");

                var objects = new List<(int Number, int Generation, byte[] Data)>
                {
                    (store.RootNumber, store.RootGeneration, PdfIncrementalWriter.WrapObject(store.RootNumber, store.RootGeneration, catalog)),
                    (outlineRootObjectNumber, 0, PdfIncrementalWriter.WrapObject(outlineRootObjectNumber, 0, BuildOutlineRoot(outlineItems)))
                };
                foreach (var item in outlineItems)
                    objects.Add((item.ObjectNumber, 0, PdfIncrementalWriter.WrapObject(item.ObjectNumber, 0, BuildOutlineItemObject(item))));

                append = PdfIncrementalWriter.BuildAppend(store, objects, out _);
            }

            PdfIncrementalWriter.AppendToFile(pdfPath, append);
            _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfiumWithBookmarkRewrite}] Bookmark rewrite successful. Count={bookmarksSnapshot.Count}");
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

        private static string EncodePdfTextString(string value)
        {
            byte[] unicode = Encoding.BigEndianUnicode.GetBytes(value ?? string.Empty);
            var bytes = new byte[unicode.Length + 2];
            bytes[0] = 0xFE;
            bytes[1] = 0xFF;
            Buffer.BlockCopy(unicode, 0, bytes, 2, unicode.Length);
            return "<" + Convert.ToHexString(bytes) + ">";
        }

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
