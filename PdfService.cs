using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text; // [추가] Encoding, StringBuilder 사용을 위해 필수
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfiumViewer;
using PdfSharp.Drawing;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;
using PdfiumDoc = PdfiumViewer.PdfDocument;
// 네임스페이스 충돌 방지 별칭
using PdfSharpDoc = PdfSharp.Pdf.PdfDocument;
using PdfSharpPage = PdfSharp.Pdf.PdfPage;
using PdfSharpRect = PdfSharp.Pdf.PdfRectangle;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private const double RENDER_SCALE = 2.0;

        // 1. PDF 로드
        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            return await Task.Run(() =>
            {
                try
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    var memoryStream = new MemoryStream(fileBytes);
                    IPdfDocument? pdfDoc = null;

                    lock (PdfiumLock)
                    {
                        pdfDoc = PdfiumDoc.Load(memoryStream);
                    }

                    var model = new PdfDocumentModel
                    {
                        FilePath = filePath,
                        FileName = Path.GetFileName(filePath),
                        PdfDocument = pdfDoc,
                        FileStream = memoryStream
                    };
                    LoadBookmarks(model);
                    return model;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                    return null;
                }
            });
        }

        // 2. 문서 초기화
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.PdfDocument == null || model.IsDisposed)
                return;
            await Task.Run(() =>
            {
                if (model.IsDisposed)
                    return;
                int pageCount = 0;
                lock (PdfiumLock)
                {
                    if (model.PdfDocument == null)
                        return;
                    pageCount = model.PdfDocument.PageCount;
                }
                if (pageCount == 0)
                    return;

                var tempPageList = new List<PdfPageViewModel>();
                for (int i = 0; i < pageCount; i++)
                {
                    if (model.IsDisposed)
                        return;
                    double pageW = 0, pageH = 0;
                    lock (PdfiumLock)
                    {
                        if (model.PdfDocument != null)
                        {
                            var size = model.PdfDocument.PageSizes[i];
                            pageW = size.Width;
                            pageH = size.Height;
                        }
                    }
                    tempPageList.Add(new PdfPageViewModel
                    {
                        PageIndex = i,
                        OriginalFilePath = model.FilePath,
                        OriginalPageIndex = i,
                        Width = pageW,
                        Height = pageH,
                        PdfPageWidthPoint = pageW,
                        PdfPageHeightPoint = pageH,
                        CropWidthPoint = pageW,
                        CropHeightPoint = pageH
                    });
                }
                if (!model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() => { foreach (var p in tempPageList) model.Pages.Add(p); });
                }
            });
        }

        // 3. 렌더링
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage)
                return;
            if (pageVM.ImageSource == null)
            {
                BitmapSource? bmpSource = null;
                lock (PdfiumLock)
                {
                    if (model.PdfDocument != null && !model.IsDisposed)
                    {
                        try
                        {
                            int renderW = (int)(pageVM.Width * RENDER_SCALE);
                            int renderH = (int)(pageVM.Height * RENDER_SCALE);
                            int dpi = (int)(96.0 * RENDER_SCALE);

                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, dpi, dpi, PdfRenderFlags.Annotations))
                            {
                                bmpSource = ToBitmapSource((DrawingBitmap)bitmap);
                                bmpSource.Freeze();
                            }
                        }
                        catch { }
                    }
                }
                if (bmpSource != null && !model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() => { pageVM.ImageSource = bmpSource; });
                }
            }
            LoadAnnotationsLazy(model, pageVM);
        }

        public static BitmapSource ToBitmapSource(DrawingBitmap bitmap)
        {
            var rect = new DrawingRectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            try
            {
                return BitmapSource.Create(bitmap.Width, bitmap.Height, bitmap.HorizontalResolution, bitmap.VerticalResolution, PixelFormats.Bgra32, null, bitmapData.Scan0, bitmapData.Stride * bitmap.Height, bitmapData.Stride);
            }
            finally { bitmap.UnlockBits(bitmapData); }
        }

        // 4. 주석 로드
        private void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (pageVM.Annotations.Count > 0)
                return;
            Task.Run(() =>
            {
                try
                {
                    if (model.IsDisposed)
                        return;
                    string path = pageVM.OriginalFilePath ?? model.FilePath;
                    if (!File.Exists(path))
                        return;
                    List<MPdfAnnotation> extracted = new List<MPdfAnnotation>();
                    try
                    {
                        using (var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                        {
                            if (pageVM.OriginalPageIndex < doc.PageCount)
                            {
                                var pdfPage = doc.Pages[pageVM.OriginalPageIndex];
                                extracted = ExtractAnnotationsFromPage(pdfPage);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"주석 로드 실패: {ex.Message}");
                    }

                    if (extracted != null && extracted.Count > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                            {
                                foreach (var item in extracted)
                                {
                                    var newAnnot = new PdfAnnotation
                                    {
                                        Type = item.Type,
                                        X = item.X,
                                        Y = item.Y,
                                        Width = item.Width,
                                        Height = item.Height,
                                        TextContent = item.TextContent,
                                        FontSize = item.FontSize,
                                        Foreground = item.Foreground ?? Brushes.Black,
                                        Background = item.Background ?? Brushes.Transparent
                                    };
                                    pageVM.Annotations.Add(newAnnot);
                                }
                            }
                        });
                    }
                }
                catch { }
            });
        }

        private List<MPdfAnnotation> ExtractAnnotationsFromPage(PdfSharpPage page)
        {
            var list = new List<MPdfAnnotation>();
            try
            {
                if (page.Annotations == null)
                    return list;
            }
            catch { return list; }

            double pageH = page.Height.Point;

            foreach (var item in page.Annotations)
            {
                try
                {
                    PdfSharp.Pdf.PdfDictionary? annotDict = null;
                    if (item is PdfSharp.Pdf.PdfDictionary dict)
                        annotDict = dict;
                    else if (item is PdfSharp.Pdf.Advanced.PdfReference reference)
                        annotDict = reference.Value as PdfSharp.Pdf.PdfDictionary;

                    if (annotDict == null)
                        continue;

                    string subtype = annotDict.Elements.ContainsKey("/Subtype") ? annotDict.Elements.GetString("/Subtype") : "";

                    var rect = annotDict.Elements.GetRectangle("/Rect");
                    double pdfLeft = rect.X1;
                    double pdfBottom = rect.Y1;
                    double pdfRight = rect.X2;
                    double pdfTop = rect.Y2;

                    double w = Math.Abs(pdfRight - pdfLeft);
                    double h = Math.Abs(pdfTop - pdfBottom);

                    double x = pdfLeft;
                    double y = pageH - Math.Max(pdfBottom, pdfTop);

                    if (subtype == "/FreeText")
                    {
                        string contents = annotDict.Elements.GetString("/Contents");
                        string da = annotDict.Elements.GetString("/DA");

                        double fontSize = 12;
                        byte fr = 0, fg = 0, fb = 0;

                        try
                        {
                            var sizeMatch = Regex.Match(da ?? "", @"(\d+(\.\d+)?) Tf");
                            if (sizeMatch.Success)
                                double.TryParse(sizeMatch.Groups[1].Value, out fontSize);

                            var colorMatch = Regex.Match(da ?? "", @"(\d+(\.\d+)?) (\d+(\.\d+)?) (\d+(\.\d+)?) rg");
                            if (colorMatch.Success)
                            {
                                fr = (byte)(double.Parse(colorMatch.Groups[1].Value) * 255);
                                fg = (byte)(double.Parse(colorMatch.Groups[3].Value) * 255);
                                fb = (byte)(double.Parse(colorMatch.Groups[5].Value) * 255);
                            }
                        }
                        catch { }

                        var textBrush = new SolidColorBrush(Color.FromRgb(fr, fg, fb));
                        textBrush.Freeze();

                        list.Add(new MPdfAnnotation { Type = AnnotationType.FreeText, X = x, Y = y, Width = w, Height = h, TextContent = contents, FontSize = fontSize, Foreground = textBrush, Background = Brushes.Transparent });
                    }
                    else if (subtype == "/Highlight")
                    {
                        byte br = 255, bg = 255, bb = 0;
                        try
                        {
                            if (annotDict.Elements.ContainsKey("/C"))
                            {
                                var cArr = annotDict.Elements.GetArray("/C");
                                if (cArr != null && cArr.Elements.Count >= 3)
                                {
                                    br = (byte)(((PdfSharp.Pdf.PdfReal)cArr.Elements[0]).Value * 255);
                                    bg = (byte)(((PdfSharp.Pdf.PdfReal)cArr.Elements[1]).Value * 255);
                                    bb = (byte)(((PdfSharp.Pdf.PdfReal)cArr.Elements[2]).Value * 255);
                                }
                            }
                        }
                        catch { }

                        var highlightBrush = new SolidColorBrush(Color.FromRgb(br, bg, bb)) { Opacity = 0.5 };
                        highlightBrush.Freeze();

                        list.Add(new MPdfAnnotation { Type = AnnotationType.Highlight, X = x, Y = y, Width = w, Height = h, Background = highlightBrush });
                    }
                }
                catch { }
            }
            return list;
        }

        // 5. 저장
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            var pagesSnapshot = model.Pages.Select(p => new PageSaveData
            {
                IsBlankPage = p.IsBlankPage,
                Width = p.Width,
                Height = p.Height,
                OriginalPageIndex = p.OriginalPageIndex,
                Annotations = p.Annotations.Cast<dynamic>().Select(a => new AnnotationSaveData
                {
                    Type = (AnnotationType)a.Type,
                    X = (double)a.X,
                    Y = (double)a.Y,
                    Width = (double)a.Width,
                    Height = (double)a.Height,
                    TextContent = (string)a.TextContent,
                    FontSize = (double)a.FontSize,
                    ForeR = (a.Foreground as SolidColorBrush)?.Color.R ?? 0,
                    ForeG = (a.Foreground as SolidColorBrush)?.Color.G ?? 0,
                    ForeB = (a.Foreground as SolidColorBrush)?.Color.B ?? 0,
                    BackR = (a.Background as SolidColorBrush)?.Color.R ?? 255,
                    BackG = (a.Background as SolidColorBrush)?.Color.G ?? 255,
                    BackB = (a.Background as SolidColorBrush)?.Color.B ?? 255,
                    IsHighlight = (AnnotationType)a.Type == AnnotationType.Highlight
                }).ToList()
            }).ToList();

            string tempFilePath = Path.GetTempFileName();
            await Task.Run(() =>
            {
                string tempSourceCopy = Path.GetTempFileName();
                File.Copy(model.FilePath, tempSourceCopy, true);

                try
                {
                    using (var outputDoc = new PdfSharpDoc())
                    {
                        var acroForm = new PdfSharp.Pdf.PdfDictionary(outputDoc);
                        outputDoc.Internals.Catalog.Elements["/AcroForm"] = acroForm;

                        using (var form = XPdfForm.FromFile(tempSourceCopy))
                        {
                            for (int i = 0; i < pagesSnapshot.Count; i++)
                            {
                                var pageData = pagesSnapshot[i];
                                PdfSharpPage newPage = outputDoc.AddPage();

                                if (pageData.Width > 0 && pageData.Height > 0)
                                {
                                    newPage.Width = XUnit.FromPoint(pageData.Width);
                                    newPage.Height = XUnit.FromPoint(pageData.Height);
                                }

                                if (!pageData.IsBlankPage && pageData.OriginalPageIndex < form.PageCount)
                                {
                                    form.PageNumber = pageData.OriginalPageIndex + 1;
                                    using (var gfx = XGraphics.FromPdfPage(newPage))
                                    {
                                        gfx.DrawImage(form, 0, 0, newPage.Width.Point, newPage.Height.Point);
                                    }
                                }
                                AddAnnotations(outputDoc, newPage, pageData.Annotations);
                            }
                        }
                        outputDoc.Save(tempFilePath);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[1차 저장 실패] {ex.Message}.");
                    throw;
                }
                finally { if (File.Exists(tempSourceCopy)) File.Delete(tempSourceCopy); }
            });

            try
            {
                model.Dispose();
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex) { throw new IOException($"파일 저장 최종 실패: {ex.Message}"); }
            finally { if (File.Exists(tempFilePath)) File.Delete(tempFilePath); }

            var newModel = await LoadPdfAsync(outputPath);
            if (newModel != null)
            {
                model.IsDisposed = false;
                model.PdfDocument = newModel.PdfDocument;
                model.FileStream = newModel.FileStream;
                model.FilePath = outputPath;
                Application.Current.Dispatcher.Invoke(() => { foreach (var p in model.Pages) p.Annotations.Clear(); });
            }
        }

        private void AddAnnotations(PdfSharpDoc doc, PdfSharpPage page, List<AnnotationSaveData> annotations)
        {
            var acroForm = GetPdfDictionary(doc.Internals.Catalog, "/AcroForm");
            if (acroForm == null)
            {
                acroForm = new PdfSharp.Pdf.PdfDictionary(doc);
                doc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
            }

            foreach (var ann in annotations)
            {
                double pdfX = ann.X;
                double pdfY = page.Height.Point - ann.Y - ann.Height;
                double pdfW = ann.Width;
                double pdfH = ann.Height;

                var rect = new PdfSharpRect(new XRect(pdfX, pdfY, pdfW, pdfH));
                var annotDict = new PdfSharp.Pdf.PdfDictionary(doc);
                annotDict.Elements["/Type"] = new PdfSharp.Pdf.PdfName("/Annot");
                annotDict.Elements["/Rect"] = rect;

                annotDict.Elements["/NM"] = new PdfSharp.Pdf.PdfString(Guid.NewGuid().ToString());
                annotDict.Elements["/M"] = new PdfSharp.Pdf.PdfString(DateTime.Now.ToString("yyyyMMddHHmmss"));

                if (ann.Type == AnnotationType.SignaturePlaceholder)
                {
                    annotDict.Elements["/Subtype"] = new PdfSharp.Pdf.PdfName("/Widget");
                    annotDict.Elements["/FT"] = new PdfSharp.Pdf.PdfName("/Sig");
                    annotDict.Elements["/T"] = new PdfSharp.Pdf.PdfString($"Signature_{Guid.NewGuid().ToString("N").Substring(0, 8)}");
                    annotDict.Elements["/F"] = new PdfSharp.Pdf.PdfInteger(4);
                    annotDict.Elements["/P"] = page;
                    AddAnnotToPageArray(page, annotDict, doc);
                    if (!acroForm.Elements.ContainsKey("/Fields"))
                        acroForm.Elements["/Fields"] = new PdfSharp.Pdf.PdfArray(doc);
                    acroForm.Elements.GetArray("/Fields")?.Elements.Add(annotDict);
                    continue;
                }

                if (ann.Type == AnnotationType.FreeText)
                {
                    annotDict.Elements["/Subtype"] = new PdfSharp.Pdf.PdfName("/FreeText");
                    annotDict.Elements["/Contents"] = new PdfSharp.Pdf.PdfString(ann.TextContent ?? "");
                    annotDict.Elements["/F"] = new PdfSharp.Pdf.PdfInteger(4);

                    double fSize = ann.FontSize < 5 ? 10 : ann.FontSize;
                    string r = (ann.ForeR / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                    string g = (ann.ForeG / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                    string b = (ann.ForeB / 255.0).ToString("0.##", CultureInfo.InvariantCulture);

                    annotDict.Elements["/DA"] = new PdfSharp.Pdf.PdfString($"/Helvetica {fSize} Tf {r} {g} {b} rg");
                    annotDict.Elements["/Q"] = new PdfSharp.Pdf.PdfInteger(0);

                    // 인코딩 문제 임시 해결을 위한 Hex 처리 적용 Appearance 생성
                    CreateAppearanceStream(doc, annotDict, pdfW, pdfH, ann.TextContent ?? "", fSize, r, g, b);
                    AddAnnotToPageArray(page, annotDict, doc);
                }
                else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                {
                    string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                    annotDict.Elements["/Subtype"] = new PdfSharp.Pdf.PdfName(subtype);
                    annotDict.Elements["/F"] = new PdfSharp.Pdf.PdfInteger(4);

                    double tl_x = pdfX;
                    double tl_y = pdfY + pdfH;
                    double tr_x = pdfX + pdfW;
                    double tr_y = pdfY + pdfH;
                    double bl_x = pdfX;
                    double bl_y = pdfY;
                    double br_x = pdfX + pdfW;
                    double br_y = pdfY;

                    var qp = new PdfSharp.Pdf.PdfArray(doc);
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(tl_x));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(tl_y));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(tr_x));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(tr_y));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(bl_x));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(bl_y));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(br_x));
                    qp.Elements.Add(new PdfSharp.Pdf.PdfReal(br_y));

                    annotDict.Elements["/QuadPoints"] = qp;

                    if (ann.Type == AnnotationType.Highlight)
                    {
                        var c = new PdfSharp.Pdf.PdfArray(doc);
                        c.Elements.Add(new PdfSharp.Pdf.PdfReal(ann.BackR / 255.0));
                        c.Elements.Add(new PdfSharp.Pdf.PdfReal(ann.BackG / 255.0));
                        c.Elements.Add(new PdfSharp.Pdf.PdfReal(ann.BackB / 255.0));
                        annotDict.Elements["/C"] = c;
                    }
                    AddAnnotToPageArray(page, annotDict, doc);
                }
            }
        }

        private void AddAnnotToPageArray(PdfSharpPage page, PdfSharp.Pdf.PdfDictionary annotDict, PdfSharpDoc doc)
        {
            if (!page.Elements.ContainsKey("/Annots"))
                page.Elements["/Annots"] = new PdfSharp.Pdf.PdfArray(doc);
            page.Elements.GetArray("/Annots")?.Elements.Add(annotDict);
        }

        private void CreateAppearanceStream(PdfSharpDoc doc, PdfSharp.Pdf.PdfDictionary annotation, double w, double h, string text, double fSize, string r, string g, string b)
        {
            var form = new PdfSharp.Pdf.PdfDictionary(doc);
            form.Elements["/Type"] = new PdfSharp.Pdf.PdfName("/XObject");
            form.Elements["/Subtype"] = new PdfSharp.Pdf.PdfName("/Form");

            var bbox = new PdfSharp.Pdf.PdfArray(doc);
            bbox.Elements.Add(new PdfSharp.Pdf.PdfReal(0));
            bbox.Elements.Add(new PdfSharp.Pdf.PdfReal(0));
            bbox.Elements.Add(new PdfSharp.Pdf.PdfReal(w));
            bbox.Elements.Add(new PdfSharp.Pdf.PdfReal(h));
            form.Elements["/BBox"] = bbox;

            var helvRef = InjectHelveticaToAcroForm(doc);
            var resFont = new PdfSharp.Pdf.PdfDictionary(doc);
            resFont.Elements["/Helv"] = helvRef;
            var resources = new PdfSharp.Pdf.PdfDictionary(doc);
            resources.Elements["/Font"] = resFont;
            form.Elements["/Resources"] = resources;

            string safeText = text.Replace("(", "\\(").Replace(")", "\\)");

            StringBuilder sb = new StringBuilder();
            sb.Append("q\n");
            sb.Append($"{r} {g} {b} rg\n");
            sb.Append("BT\n");
            sb.AppendFormat(CultureInfo.InvariantCulture, "/Helv {0:0.##} Tf\n", fSize);
            double textY = h - fSize - 2;
            if (textY < 0)
                textY = 0;
            sb.AppendFormat(CultureInfo.InvariantCulture, "2 {0:0.##} Td\n", textY);
            sb.AppendFormat("({0}) Tj\n", safeText);
            sb.Append("ET\n");
            sb.Append("Q");

            // [Fix] Encoding namespace required here
            form.CreateStream(Encoding.UTF8.GetBytes(sb.ToString()));

            var ap = new PdfSharp.Pdf.PdfDictionary(doc);
            annotation.Elements["/AP"] = ap;
            ap.Elements["/N"] = form;
        }

        private PdfSharp.Pdf.PdfDictionary? GetPdfDictionary(PdfSharp.Pdf.PdfDictionary? parent, string key)
        {
            if (parent == null || !parent.Elements.ContainsKey(key))
                return null;
            var item = parent.Elements[key];
            if (item is PdfSharp.Pdf.Advanced.PdfReference reference)
                return reference.Value as PdfSharp.Pdf.PdfDictionary;
            return item as PdfSharp.Pdf.PdfDictionary;
        }

        private PdfSharp.Pdf.PdfDictionary? InjectHelveticaToAcroForm(PdfSharpDoc doc)
        {
            var catalog = doc.Internals.Catalog;
            var acroForm = GetPdfDictionary(catalog, "/AcroForm");
            if (acroForm == null)
            {
                acroForm = new PdfSharp.Pdf.PdfDictionary(doc);
                catalog.Elements["/AcroForm"] = acroForm;
            }
            var dr = GetPdfDictionary(acroForm, "/DR");
            if (dr == null)
            {
                dr = new PdfSharp.Pdf.PdfDictionary(doc);
                acroForm.Elements["/DR"] = dr;
            }
            var fontDict = GetPdfDictionary(dr, "/Font");
            if (fontDict == null)
            {
                fontDict = new PdfSharp.Pdf.PdfDictionary(doc);
                dr.Elements["/Font"] = fontDict;
            }
            if (!fontDict.Elements.ContainsKey("/Helvetica"))
            {
                var helvetica = new PdfSharp.Pdf.PdfDictionary(doc);
                helvetica.Elements["/Type"] = new PdfSharp.Pdf.PdfName("/Font");
                helvetica.Elements["/Subtype"] = new PdfSharp.Pdf.PdfName("/Type1");
                helvetica.Elements["/BaseFont"] = new PdfSharp.Pdf.PdfName("/Helvetica");
                helvetica.Elements["/Encoding"] = new PdfSharp.Pdf.PdfName("/WinAnsiEncoding");
                fontDict.Elements["/Helvetica"] = helvetica;
                return helvetica;
            }
            return GetPdfDictionary(fontDict, "/Helvetica");
        }

        public void LoadBookmarks(PdfDocumentModel model)
        {
            lock (PdfiumLock)
            {
                if (model.PdfDocument == null)
                    return;
                var bookmarks = model.PdfDocument.Bookmarks;
                if (bookmarks != null)
                {
                    Application.Current.Dispatcher.Invoke(() => model.Bookmarks.Clear());
                    foreach (var bm in bookmarks)
                    {
                        var vm = ConvertToViewModel(bm, null);
                        Application.Current.Dispatcher.Invoke(() => model.Bookmarks.Add(vm));
                    }
                }
            }
        }
        private PdfBookmarkViewModel ConvertToViewModel(PdfBookmark bookmark, PdfBookmarkViewModel? parent)
        {
            var vm = new PdfBookmarkViewModel { Title = bookmark.Title, PageIndex = bookmark.PageIndex, Parent = parent, IsExpanded = true };
            foreach (var child in bookmark.Children)
                vm.Children.Add(ConvertToViewModel(child, vm));
            return vm;
        }

        public class MPdfAnnotation
        {
            public AnnotationType Type
            {
                get; set;
            }
            public double X
            {
                get; set;
            }
            public double Y
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public string TextContent { get; set; } = ""; public double FontSize
            {
                get; set;
            }
            public Brush? Foreground
            {
                get; set;
            }
            public Brush? Background
            {
                get; set;
            }
        }
        private class PageSaveData
        {
            public bool IsBlankPage
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public int OriginalPageIndex
            {
                get; set;
            }
            public List<AnnotationSaveData> Annotations { get; set; } = new List<AnnotationSaveData>();
        }
        private class AnnotationSaveData
        {
            public AnnotationType Type
            {
                get; set;
            }
            public double X
            {
                get; set;
            }
            public double Y
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public string TextContent { get; set; } = ""; public double FontSize
            {
                get; set;
            }
            public byte ForeR
            {
                get; set;
            }
            public byte ForeG
            {
                get; set;
            }
            public byte ForeB
            {
                get; set;
            }
            public byte BackR
            {
                get; set;
            }
            public byte BackG
            {
                get; set;
            }
            public byte BackB
            {
                get; set;
            }
            public bool IsHighlight
            {
                get; set;
            }
        }
    }
}