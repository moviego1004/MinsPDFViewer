using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using PdfiumViewer;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;
using PdfSharpRectangle = PdfSharp.Pdf.PdfRectangle;

namespace MinsPDFViewer
{
    internal sealed class PdfSharpLegacySaveService
    {
        private readonly Action<string>? _log;

        public PdfSharpLegacySaveService(Action<string>? log = null)
        {
            _log = log;
        }

        public async Task<PdfSaveEngine> SaveAsync(
            string originalFilePath,
            string outputPath,
            List<PageSaveData> pagesSnapshot,
            List<BookmarkSaveData> bookmarksSnapshot)
        {
            PdfSharpRuntime.EnsureInitialized();

            return await Task.Run(() =>
            {
                PdfSaveEngine engine = PdfSaveEngine.PdfSharpLegacy;
                string tempWorkPath = Path.GetTempFileName();
                File.Copy(originalFilePath, tempWorkPath, true);
                _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfSharpLegacy}] Starting PDFsharp legacy save. Temp={tempWorkPath}");

                bool standardSaveSuccess = false;
                var pageMapping = new Dictionary<int, PdfPage>();

                try
                {
                    using (var sourceDoc = PdfReader.Open(tempWorkPath, PdfDocumentOpenMode.Import))
                    using (var outputDoc = new PdfSharpDocument())
                    {
                        _log?.Invoke($"Document opened for Import. PageCount: {sourceDoc.PageCount}, SnapshotCount: {pagesSnapshot.Count}");

                        for (int i = 0; i < pagesSnapshot.Count; i++)
                        {
                            var pageData = pagesSnapshot[i];
                            int originalIdx = pageData.OriginalPageIndex;
                            if (originalIdx < 0 || originalIdx >= sourceDoc.PageCount)
                                continue;

                            var pdfPage = outputDoc.AddPage(sourceDoc.Pages[originalIdx]);
                            pageMapping[originalIdx] = pdfPage;
                            pdfPage.Annotations.Clear();

                            DrawOcrText(outputDoc, pdfPage, pageData.OcrWords, pageData.Width, pageData.Height);
                            DrawAnnotationsOnPage(outputDoc, pdfPage, i, pagesSnapshot);
                        }

                        foreach (var bm in bookmarksSnapshot)
                            PdfSharpBookmarkService.AddBookmarkToPdf(outputDoc.Outlines, bm, pageMapping);

                        var catalog = outputDoc.Internals.Catalog;
                        var acroForm = catalog.Elements.GetDictionary("/AcroForm");
                        if (acroForm == null)
                        {
                            acroForm = new PdfDictionary(outputDoc);
                            catalog.Elements["/AcroForm"] = acroForm;
                        }
                        acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);

                        outputDoc.Save(outputPath);
                        standardSaveSuccess = true;
                        _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfSharpLegacy}] Standard import save successful.");
                    }
                }
                catch (Exception ex)
                {
                    _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfSharpLegacy}] Standard import save failed: {ex}. Trying raster fallback.");
                }

                if (!standardSaveSuccess)
                {
                    try
                    {
                        engine = PdfSaveEngine.PdfSharpLegacyRasterFallback;
                        pageMapping.Clear();

                        using (var pdfiumDoc = PdfiumViewer.PdfDocument.Load(tempWorkPath))
                        using (var outputDoc = new PdfSharpDocument())
                        {
                            _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfSharpLegacyRasterFallback}] Starting raster fallback save.");

                            for (int i = 0; i < pagesSnapshot.Count; i++)
                            {
                                var pageData = pagesSnapshot[i];
                                int originalIdx = pageData.OriginalPageIndex;
                                if (originalIdx < 0 || originalIdx >= pdfiumDoc.PageCount)
                                    continue;

                                var size = pdfiumDoc.PageSizes[originalIdx];
                                using (var bitmap = pdfiumDoc.Render(originalIdx, (int)size.Width * 2, (int)size.Height * 2, 192, 192, PdfRenderFlags.Annotations))
                                using (var ms = new MemoryStream())
                                {
                                    bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                    ms.Position = 0;

                                    var page = outputDoc.AddPage();
                                    page.Width = size.Width;
                                    page.Height = size.Height;
                                    pageMapping[originalIdx] = page;

                                    using (var gfx = XGraphics.FromPdfPage(page))
                                    using (var xImage = XImage.FromStream(ms))
                                    {
                                        gfx.DrawImage(xImage, 0, 0, page.Width, page.Height);
                                        DrawOcrText(outputDoc, page, pageData.OcrWords, pageData.Width, pageData.Height);
                                        DrawAnnotationsOnPage(outputDoc, page, i, pagesSnapshot);
                                    }
                                }
                            }

                            foreach (var bm in bookmarksSnapshot)
                                PdfSharpBookmarkService.AddBookmarkToPdf(outputDoc.Outlines, bm, pageMapping);

                            outputDoc.Save(outputPath);
                            _log?.Invoke($"[SaveEngine:{PdfSaveEngine.PdfSharpLegacyRasterFallback}] Raster fallback save successful.");
                        }
                    }
                    catch (Exception ex2)
                    {
                        _log?.Invoke($"Fallback save failed: {ex2}");
                        throw;
                    }
                }

                try { if (File.Exists(tempWorkPath)) File.Delete(tempWorkPath); } catch { }
                return engine;
            });
        }

        private static PdfFormXObject? GetPdfForm(XForm form)
        {
            var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            return prop?.GetValue(form) as PdfFormXObject;
        }

        private void DrawOcrText(PdfSharpDocument doc, PdfPage page, List<OcrWordInfo> words, double viewWidth, double viewHeight)
        {
            if (words == null || words.Count == 0) return;

            try
            {
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    var transparentBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0, 0));
                    double pdfPageW = page.Width.Point;
                    double pdfPageH = page.Height.Point;
                    double scaleX = pdfPageW / viewWidth;
                    double scaleY = pdfPageH / viewHeight;

                    foreach (var word in words)
                    {
                        double fSize = word.BoundingBox.Height * scaleY;
                        if (fSize <= 0) fSize = 10;
                        var font = new XFont("Malgun Gothic", fSize, XFontStyleEx.Regular);
                        double x = word.BoundingBox.X * scaleX;
                        double y = word.BoundingBox.Y * scaleY;
                        gfx.DrawString(word.Text, font, transparentBrush, x, y + (fSize * 0.8));
                    }
                }
            }
            catch (Exception ex)
            {
                _log?.Invoke($"Error drawing OCR text: {ex.Message}");
            }
        }

        private void DrawAnnotationsOnPage(PdfSharpDocument doc, PdfPage pdfPage, int pageIndex, List<PageSaveData> snapshots)
        {
            if (pageIndex >= snapshots.Count) return;
            var pageData = snapshots[pageIndex];

            foreach (var ann in pageData.Annotations)
            {
                try
                {
                    double pdfPageH = pdfPage.Height.Point;
                    double pdfPageW = pdfPage.Width.Point;
                    double scaleX = pdfPageW / pageData.Width;
                    double scaleY = pdfPageH / pageData.Height;

                    var rect = new PdfSharpRectangle(new XRect(
                        ann.X * scaleX,
                        pdfPageH - ((ann.Y + ann.Height) * scaleY),
                        ann.Width * scaleX,
                        ann.Height * scaleY));

                    if (ann.Type == AnnotationType.FreeText)
                    {
                        AddFreeTextAnnotation(doc, pdfPage, ann, rect, scaleX, scaleY);
                    }
                    else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                    {
                        AddMarkupAnnotation(doc, pdfPage, ann, rect, scaleX, scaleY);
                    }
                }
                catch (Exception ex)
                {
                    _log?.Invoke($"Annotation drawing error: {ex.Message}");
                }
            }
        }

        private static void AddFreeTextAnnotation(
            PdfSharpDocument doc,
            PdfPage pdfPage,
            AnnotationSaveData ann,
            PdfSharpRectangle rect,
            double scaleX,
            double scaleY)
        {
            var pdfAnnot = new GenericPdfAnnotation(doc);
            pdfAnnot.Elements["/Subtype"] = new PdfName("/FreeText");
            pdfAnnot.Elements["/Rect"] = rect;
            pdfAnnot.Elements["/Contents"] = new PdfString(ann.TextContent, PdfStringEncoding.Unicode);

            var form = new XForm(doc, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY));
            using (var gfx = XGraphics.FromForm(form))
            {
                var fontName = "Noto Sans KR";
                var font = new XFont(fontName, ann.FontSize * scaleY, ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);
                if (font.FontFamily.Name != fontName)
                    font = new XFont("Malgun Gothic", ann.FontSize * scaleY, ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);

                var brush = new XSolidBrush(XColor.FromArgb(ann.ForegroundColor.A, ann.ForegroundColor.R, ann.ForegroundColor.G, ann.ForegroundColor.B));
                gfx.DrawString(ann.TextContent, font, brush, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY), XStringFormats.TopLeft);
            }

            var pdfForm = GetPdfForm(form);
            string fontKey = "/F1";
            var dr = new PdfDictionary(doc);
            var fontDict = new PdfDictionary(doc);
            bool fontFound = false;

            if (pdfForm != null)
            {
                var resources = pdfForm.Elements.GetDictionary("/Resources");
                var formFontDict = resources?.Elements.GetDictionary("/Font");
                if (formFontDict != null && formFontDict.Elements.Count > 0)
                {
                    foreach (var key in formFontDict.Elements.Keys)
                    {
                        fontDict.Elements[key] = formFontDict.Elements[key];
                        fontKey = key;
                        fontFound = true;
                    }
                }
            }

            if (!fontFound)
            {
                PdfDictionary? fallbackFont = null;
                foreach (var obj in doc.Internals.GetAllObjects())
                {
                    if (obj is PdfDictionary d && d.Elements.GetName("/Type") == "/Font")
                    {
                        var baseFont = d.Elements.GetName("/BaseFont");
                        var subtype = d.Elements.GetName("/Subtype");
                        if (baseFont.Contains("Noto") || baseFont.Contains("Malgun") || subtype == "/Type0")
                        {
                            fontDict.Elements["/F1"] = d.Reference;
                            fontKey = "/F1";
                            fontFound = true;
                            break;
                        }

                        fallbackFont ??= d;
                    }
                }

                if (!fontFound && fallbackFont != null)
                {
                    fontDict.Elements["/F1"] = fallbackFont.Reference;
                    fontKey = "/F1";
                    fontFound = true;
                }
            }

            if (!fontFound)
            {
                var f = new PdfDictionary(doc);
                f.Elements["/Type"] = new PdfName("/Font");
                f.Elements["/Subtype"] = new PdfName("/Type1");
                f.Elements["/BaseFont"] = new PdfName("/Helvetica");
                f.Elements["/Encoding"] = new PdfName("/WinAnsiEncoding");
                doc.Internals.AddObject(f);
                fontDict.Elements["/F1"] = f.Reference;
                fontKey = "/F1";
            }

            dr.Elements["/Font"] = fontDict;
            pdfAnnot.Elements["/DR"] = dr;

            if (pdfForm != null)
            {
                var apDict = new PdfDictionary(doc);
                apDict.Elements["/N"] = pdfForm.Reference;
                pdfAnnot.Elements["/AP"] = apDict;
            }

            double finalFontSize = Math.Max(1.0, ann.FontSize * scaleY);
            string colorStr = string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} rg",
                ann.ForegroundColor.R / 255.0, ann.ForegroundColor.G / 255.0, ann.ForegroundColor.B / 255.0);

            pdfAnnot.Elements["/DA"] = new PdfString(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "{0} {1:0.###} Tf {2}", fontKey, finalFontSize, colorStr));

            pdfPage.Annotations.Add(pdfAnnot);
        }

        private static void AddMarkupAnnotation(
            PdfSharpDocument doc,
            PdfPage pdfPage,
            AnnotationSaveData ann,
            PdfSharpRectangle rect,
            double scaleX,
            double scaleY)
        {
            var pdfAnnot = new GenericPdfAnnotation(doc);
            pdfAnnot.Elements["/Subtype"] = new PdfName(ann.Type == AnnotationType.Highlight ? "/Highlight" : "/Underline");
            pdfAnnot.Elements["/Rect"] = rect;

            var form = new XForm(doc, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY));
            using (var gfx = XGraphics.FromForm(form))
            {
                if (ann.Type == AnnotationType.Highlight)
                {
                    var brush = new XSolidBrush(XColor.FromArgb(ann.BackgroundColor.A, ann.BackgroundColor.R, ann.BackgroundColor.G, ann.BackgroundColor.B));
                    gfx.DrawRectangle(brush, 0, 0, ann.Width * scaleX, ann.Height * scaleY);
                }
                else
                {
                    var pen = new XPen(XColors.Black, 1 * scaleY);
                    gfx.DrawLine(pen, 0, (ann.Height * scaleY) - 1, ann.Width * scaleX, (ann.Height * scaleY) - 1);
                }
            }

            var apDict = new PdfDictionary(doc);
            var pdfForm = GetPdfForm(form);
            if (pdfForm != null) apDict.Elements["/N"] = pdfForm.Reference;
            pdfAnnot.Elements["/AP"] = apDict;
            pdfPage.Annotations.Add(pdfAnnot);
        }
    }
}
