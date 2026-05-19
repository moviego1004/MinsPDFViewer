using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MinsPDFViewer
{
    internal sealed class PdfSharpBookmarkService
    {
        private readonly Action<string>? _log;

        public PdfSharpBookmarkService(Action<string>? log = null)
        {
            _log = log;
        }

        public async Task RewriteBookmarksAsync(string pdfPath, List<BookmarkSaveData> bookmarksSnapshot)
        {
            PdfSharpRuntime.EnsureInitialized();

            await Task.Run(() =>
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");
                try
                {
                    using (var sourceDoc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
                    using (var outputDoc = new PdfDocument())
                    {
                        var pageMapping = new Dictionary<int, PdfPage>();
                        for (int i = 0; i < sourceDoc.PageCount; i++)
                        {
                            var page = outputDoc.AddPage(sourceDoc.Pages[i]);
                            pageMapping[i] = page;
                        }

                        foreach (var bm in bookmarksSnapshot)
                            AddBookmarkToPdf(outputDoc.Outlines, bm, pageMapping);

                        outputDoc.Save(tempPath);
                    }

                    File.Copy(tempPath, pdfPath, true);
                    _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfiumWithPdfSharpBookmarkRewrite}] Bookmark rewrite successful. Count={bookmarksSnapshot.Count}");
                }
                catch (Exception ex)
                {
                    _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfiumWithPdfSharpBookmarkRewrite}] Bookmark rewrite failed: {ex}");
                    throw;
                }
                finally
                {
                    try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                }
            });
        }

        internal static void AddBookmarkToPdf(
            PdfOutlineCollection collection,
            BookmarkSaveData bmData,
            Dictionary<int, PdfPage> mapping)
        {
            PdfPage? destPage = null;
            if (mapping.ContainsKey(bmData.OriginalPageIndex))
                destPage = mapping[bmData.OriginalPageIndex];

            var outline = destPage != null
                ? collection.Add(bmData.Title, destPage, true)
                : collection.Add(bmData.Title, null, true);

            foreach (var child in bmData.Children)
                AddBookmarkToPdf(outline.Outlines, child, mapping);
        }
    }
}
