using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Globalization;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.AcroForms;

// [중요] 네임스페이스 충돌 방지
using PdfSharpRect = PdfSharp.Pdf.PdfRectangle;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private readonly IDocLib _docLib;
        private const double RENDER_SCALE = 2.0;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        // =========================================================
        // 1. PDF 로드
        // =========================================================
        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            return await Task.Run(() =>
            {
                try
                {
                    IDocReader docReader;
                    byte[] fileBytes = ReadFileSafely(filePath);

                    lock (PdfiumLock)
                    {
                        docReader = Application.Current.Dispatcher.Invoke(() =>
                            _docLib.GetDocReader(fileBytes, new PageDimensions(RENDER_SCALE)));
                    }

                    var model = new PdfDocumentModel
                    {
                        FilePath = filePath,
                        FileName = Path.GetFileName(filePath),
                        DocLib = _docLib,
                        DocReader = docReader
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

        // =========================================================
        // 2. 초기화
        // =========================================================
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null || model.IsDisposed)
                return;

            await Task.Run(() =>
            {
                if (model.IsDisposed)
                    return;
                int pageCount = 0;

                lock (PdfiumLock)
                {
                    if (model.DocReader == null)
                        return;
                    pageCount = model.DocReader.GetPageCount();
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
                        if (model.DocReader != null)
                        {
                            using (var pr = model.DocReader.GetPageReader(i))
                            {
                                pageW = pr.GetPageWidth();
                                pageH = pr.GetPageHeight();
                            }
                        }
                    }

                    double originalW = pageW / RENDER_SCALE;
                    double originalH = pageH / RENDER_SCALE;

                    tempPageList.Add(new PdfPageViewModel
                    {
                        PageIndex = i,
                        OriginalFilePath = model.FilePath,
                        OriginalPageIndex = i,
                        Width = originalW,
                        Height = originalH,
                        PdfPageWidthPoint = originalW,
                        PdfPageHeightPoint = originalH,
                        CropWidthPoint = originalW,
                        CropHeightPoint = originalH
                    });
                }

                if (!model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        foreach (var p in tempPageList)
                            model.Pages.Add(p);
                    });
                }
            });
        }

        // =========================================================
        // 3. 렌더링 및 주석 읽기
        // =========================================================
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage)
                return;

            if (pageVM.ImageSource == null)
            {
                byte[]? imgBytes = null;
                int w = 0, h = 0;

                lock (PdfiumLock)
                {
                    if (model.DocReader != null && !model.IsDisposed)
                    {
                        try
                        {
                            using (var pr = model.DocReader.GetPageReader(pageVM.OriginalPageIndex))
                            {
                                imgBytes = pr.GetImage(RenderFlags.RenderAnnotations);
                                w = pr.GetPageWidth();
                                h = pr.GetPageHeight();
                            }
                        }
                        catch { }
                    }
                }

                if (imgBytes != null && !model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, imgBytes, w * 4);
                        bmp.Freeze();
                        pageVM.ImageSource = bmp;
                    });
                }
            }

            LoadAnnotationsLazy(model, pageVM);
        }

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

                    List<MinsPDFViewer.PdfAnnotation> extracted = new List<MinsPDFViewer.PdfAnnotation>();

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
                        System.Diagnostics.Debug.WriteLine($"[주석 로드 경고] {ex.Message}");
                    }

                    if (extracted != null && extracted.Count > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                            {
                                foreach (var appAnnotation in extracted)
                                    pageVM.Annotations.Add(appAnnotation);
                            }
                        });
                    }
                }
                catch { }
            });
        }

        private List<MinsPDFViewer.PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            var list = new List<MinsPDFViewer.PdfAnnotation>();
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
                    PdfDictionary annotDict = item as PdfDictionary;
                    if (annotDict == null && item is PdfReference pdfRef)
                        annotDict = pdfRef.Value as PdfDictionary;

                    if (annotDict == null)
                        continue;

                    string subtype = annotDict.Elements.ContainsKey("/Subtype") ? annotDict.Elements.GetString("/Subtype") : "";

                    if (subtype == "/FreeText")
                    {
                        var rect = annotDict.Elements.GetRectangle("/Rect");
                        string contents = annotDict.Elements.GetString("/Contents");
                        string da = annotDict.Elements.GetString("/DA");

                        double w = rect.Width;
                        double h = rect.Height;
                        double x = rect.X1;
                        double y = pageH - rect.Y1 - h;

                        double fontSize = 12;
                        byte fr = 0, fg = 0, fb = 0;

                        var sizeMatch = Regex.Match(da ?? "", @"(\d+(\.\d+)?) Tf");
                        if (sizeMatch.Success)
                            double.TryParse(sizeMatch.Groups[1].Value, out fontSize);

                        var colorMatch = Regex.Match(da ?? "", @"(\d+(\.\d+)?) (\d+(\.\d+)?) (\d+(\.\d+)?) rg");
                        if (colorMatch.Success)
                        {
                            double rVal = double.Parse(colorMatch.Groups[1].Value);
                            double gVal = double.Parse(colorMatch.Groups[3].Value);
                            double bVal = double.Parse(colorMatch.Groups[5].Value);
                            fr = (byte)(rVal * 255);
                            fg = (byte)(gVal * 255);
                            fb = (byte)(bVal * 255);
                        }

                        var textBrush = new SolidColorBrush(Color.FromRgb(fr, fg, fb));
                        textBrush.Freeze(); // [중요]

                        var newAnn = new MinsPDFViewer.PdfAnnotation
                        {
                            Type = AnnotationType.FreeText,
                            X = x,
                            Y = y,
                            Width = w,
                            Height = h,
                            TextContent = contents,
                            FontSize = fontSize,
                            Foreground = textBrush,
                            Background = Brushes.Transparent
                        };
                        list.Add(newAnn);
                    }
                    else if (subtype == "/Highlight")
                    {
                        var rect = annotDict.Elements.GetRectangle("/Rect");
                        byte br = 255, bg = 255, bb = 0;
                        if (annotDict.Elements.ContainsKey("/C"))
                        {
                            var cArr = annotDict.Elements.GetArray("/C");
                            if (cArr != null && cArr.Elements.Count >= 3)
                            {
                                double cr = ((PdfReal)cArr.Elements[0]).Value;
                                double cg = ((PdfReal)cArr.Elements[1]).Value;
                                double cb = ((PdfReal)cArr.Elements[2]).Value;
                                br = (byte)(cr * 255);
                                bg = (byte)(cg * 255);
                                bb = (byte)(cb * 255);
                            }
                        }
                        double w = rect.Width;
                        double h = rect.Height;
                        double x = rect.X1;
                        double y = pageH - rect.Y1 - h;

                        var highlightBrush = new SolidColorBrush(Color.FromRgb(br, bg, bb)) { Opacity = 0.5 };
                        highlightBrush.Freeze(); // [중요]

                        var newAnn = new MinsPDFViewer.PdfAnnotation
                        {
                            Type = AnnotationType.Highlight,
                            X = x,
                            Y = y,
                            Width = w,
                            Height = h,
                            Background = highlightBrush
                        };
                        list.Add(newAnn);
                    }
                }
                catch { }
            }
            return list;
        }

        // =========================================================
        // 4. 저장 (순수 XPdfForm 방식 - PdfReader.Open 제거)
        // =========================================================
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            // UI 스냅샷
            var pagesSnapshot = model.Pages.Select(p => new PageSaveData
            {
                IsBlankPage = p.IsBlankPage,
                Width = p.Width,
                Height = p.Height,
                OriginalPageIndex = p.OriginalPageIndex,
                Annotations = p.Annotations.Select(a => new AnnotationSaveData
                {
                    Type = a.Type,
                    X = a.X,
                    Y = a.Y,
                    Width = a.Width,
                    Height = a.Height,
                    TextContent = a.TextContent,
                    FontSize = a.FontSize,
                    ForeR = (a.Foreground as SolidColorBrush)?.Color.R ?? 0,
                    ForeG = (a.Foreground as SolidColorBrush)?.Color.G ?? 0,
                    ForeB = (a.Foreground as SolidColorBrush)?.Color.B ?? 0,
                    BackR = (a.Background as SolidColorBrush)?.Color.R ?? 255,
                    BackG = (a.Background as SolidColorBrush)?.Color.G ?? 255,
                    BackB = (a.Background as SolidColorBrush)?.Color.B ?? 255,
                    IsHighlight = (a.Type == AnnotationType.Highlight)
                }).ToList()
            }).ToList();

            string tempFilePath = Path.GetTempFileName();

            await Task.Run(() =>
            {
                try
                {
                    // 1. 임시 파일로 복사
                    string tempSourcePath = Path.GetTempFileName();
                    File.Copy(model.FilePath, tempSourcePath, true);

                    // 2. 새 문서 생성
                    using (var outputDoc = new PdfDocument())
                    {
                        var acroForm = new PdfDictionary(outputDoc);
                        outputDoc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
                        acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);

                        InjectHelveticaToAcroForm(outputDoc);

                        // 3. [핵심] XPdfForm만 사용 (PdfReader.Open 사용 안함)
                        // 이렇게 하면 원본 파일의 내부 구조 파싱 에러(NotImplemented)를 피할 수 있습니다.
                        using (var form = XPdfForm.FromFile(tempSourcePath))
                        {
                            for (int i = 0; i < pagesSnapshot.Count; i++)
                            {
                                var pageData = pagesSnapshot[i];
                                PdfPage newPage = outputDoc.AddPage();

                                // 페이지 크기 설정
                                if (pageData.Width > 0 && pageData.Height > 0)
                                {
                                    newPage.Width = XUnit.FromPoint(pageData.Width);
                                    newPage.Height = XUnit.FromPoint(pageData.Height);
                                }

                                // 원본 그리기 (배경)
                                if (!pageData.IsBlankPage && pageData.OriginalPageIndex < form.PageCount)
                                {
                                    form.PageNumber = pageData.OriginalPageIndex + 1; // 1-based index
                                    using (var gfx = XGraphics.FromPdfPage(newPage))
                                    {
                                        gfx.DrawImage(form, 0, 0, newPage.Width.Point, newPage.Height.Point);
                                    }
                                }

                                // 주석 및 서명 추가 (Dictionary 방식 사용 - 빌드 에러 해결)
                                foreach (var ann in pageData.Annotations)
                                {
                                    double pdfX = ann.X;
                                    double pdfY = newPage.Height.Point - ann.Y - ann.Height;
                                    double pdfW = ann.Width;
                                    double pdfH = ann.Height;
                                    var rect = new PdfSharpRect(new XRect(pdfX, pdfY, pdfW, pdfH));

                                    // (1) 전자서명 필드
                                    if (ann.Type == AnnotationType.SignaturePlaceholder)
                                    {
                                        string fieldName = $"Signature_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                                        var sigField = new GenericPdfAnnotation(outputDoc);
                                        sigField.Rectangle = rect;
                                        sigField.Elements["/Subtype"] = new PdfName("/Widget");
                                        sigField.Elements["/FT"] = new PdfName("/Sig");
                                        sigField.Elements["/T"] = new PdfString(fieldName);
                                        sigField.Elements["/P"] = newPage;
                                        sigField.Elements["/F"] = new PdfInteger(4);

                                        newPage.Annotations.Add(sigField);

                                        if (!acroForm.Elements.ContainsKey("/Fields"))
                                            acroForm.Elements["/Fields"] = new PdfArray(outputDoc);

                                        acroForm.Elements.GetArray("/Fields")?.Elements.Add(sigField);
                                        continue;
                                    }

                                    // (2) 일반 주석 (텍스트, 형광펜)
                                    if (pdfW < 5)
                                        pdfW = 20;
                                    if (pdfH < 5)
                                        pdfH = 10;

                                    var newPdfAnn = new GenericPdfAnnotation(outputDoc);
                                    newPdfAnn.Rectangle = rect;
                                    newPdfAnn.Elements["/F"] = new PdfInteger(4);

                                    if (ann.Type == AnnotationType.FreeText)
                                    {
                                        newPdfAnn.Elements["/Subtype"] = new PdfName("/FreeText");
                                        newPdfAnn.Elements["/Contents"] = new PdfString(ann.TextContent ?? "");

                                        double fSize = ann.FontSize < 5 ? 10 : ann.FontSize;
                                        string r = (ann.ForeR / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                                        string g = (ann.ForeG / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                                        string b = (ann.ForeB / 255.0).ToString("0.##", CultureInfo.InvariantCulture);

                                        newPdfAnn.Elements["/DA"] = new PdfString($"/Helvetica {fSize} Tf {r} {g} {b} rg");
                                        newPdfAnn.Elements["/Q"] = new PdfInteger(0);

                                        // AP Stream 직접 생성 (가시성 확보)
                                        CreateAppearanceStream(outputDoc, newPdfAnn, pdfW, pdfH, ann.TextContent ?? "", fSize, r, g, b);
                                        newPage.Annotations.Add(newPdfAnn);
                                    }
                                    else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                                    {
                                        var qp = new PdfArray(outputDoc);
                                        qp.Elements.Add(new PdfReal(pdfX));
                                        qp.Elements.Add(new PdfReal(pdfY + pdfH));
                                        qp.Elements.Add(new PdfReal(pdfX + pdfW));
                                        qp.Elements.Add(new PdfReal(pdfY + pdfH));
                                        qp.Elements.Add(new PdfReal(pdfX));
                                        qp.Elements.Add(new PdfReal(pdfY));
                                        qp.Elements.Add(new PdfReal(pdfX + pdfW));
                                        qp.Elements.Add(new PdfReal(pdfY));
                                        newPdfAnn.Elements["/QuadPoints"] = qp;

                                        string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                        newPdfAnn.Elements["/Subtype"] = new PdfName(subtype);

                                        if (ann.Type == AnnotationType.Highlight)
                                        {
                                            var c = new PdfArray(outputDoc);
                                            c.Elements.Add(new PdfReal(ann.BackR / 255.0));
                                            c.Elements.Add(new PdfReal(ann.BackG / 255.0));
                                            c.Elements.Add(new PdfReal(ann.BackB / 255.0));
                                            newPdfAnn.Elements["/C"] = c;
                                        }
                                        newPage.Annotations.Add(newPdfAnn);
                                    }
                                }
                            }
                        }

                        outputDoc.Save(tempFilePath);
                    }

                    if (File.Exists(tempSourcePath))
                        File.Delete(tempSourcePath);
                }
                catch (Exception ex)
                {
                    throw new Exception($"PDF 저장 실패: {ex.Message}");
                }
            });

            // 파일 이동
            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}");
            }
            finally
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }

            // 모델 리로드
            if (string.Equals(model.FilePath, outputPath, StringComparison.OrdinalIgnoreCase))
            {
                lock (PdfiumLock)
                {
                    model.DocReader = null;
                }
                byte[] newFileBytes = ReadFileSafely(outputPath);
                lock (PdfiumLock)
                {
                    model.DocReader = Application.Current.Dispatcher.Invoke(() =>
                        _docLib.GetDocReader(newFileBytes, new PageDimensions(RENDER_SCALE)));
                }
                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var p in model.Pages)
                        p.Annotations.Clear();
                });
                model.IsDisposed = false;
            }
        }

        private void CreateAppearanceStream(PdfDocument doc, PdfSharp.Pdf.Annotations.PdfAnnotation annotation, double w, double h, string text, double fSize, string r, string g, string b)
        {
            var form = new PdfDictionary(doc);
            form.Elements["/Type"] = new PdfName("/XObject");
            form.Elements["/Subtype"] = new PdfName("/Form");
            var bbox = new PdfArray(doc);
            bbox.Elements.Add(new PdfReal(0));
            bbox.Elements.Add(new PdfReal(0));
            bbox.Elements.Add(new PdfReal(w));
            bbox.Elements.Add(new PdfReal(h));
            form.Elements["/BBox"] = bbox;
            var helvRef = InjectHelveticaToAcroForm(doc);
            var resFont = new PdfDictionary(doc);
            resFont.Elements["/Helv"] = helvRef;
            var resources = new PdfDictionary(doc);
            resources.Elements["/Font"] = resFont;
            form.Elements["/Resources"] = resources;
            string safeText = text.Replace("(", "\\(").Replace(")", "\\)");
            StringBuilder sb = new StringBuilder();
            sb.Append("q\n");
            sb.Append($"{r} {g} {b} rg\n");
            sb.Append("BT\n");
            sb.AppendFormat(CultureInfo.InvariantCulture, "/Helv {0:0.##} Tf\n", fSize);
            sb.AppendFormat(CultureInfo.InvariantCulture, "2 {0:0.##} Td\n", h - fSize - 2);
            sb.AppendFormat("({0}) Tj\n", safeText);
            sb.Append("ET\n");
            sb.Append("Q");
            form.CreateStream(Encoding.UTF8.GetBytes(sb.ToString()));
            var ap = new PdfDictionary(doc);
            annotation.Elements["/AP"] = ap;
            ap.Elements["/N"] = form;
        }

        private byte[] ReadFileSafely(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                return ms.ToArray();
            }
        }
        private PdfDictionary? GetPdfDictionary(PdfDictionary? parent, string key)
        {
            if (parent == null)
                return null;
            if (!parent.Elements.ContainsKey(key))
                return null;
            var item = parent.Elements[key];
            if (item is PdfReference reference)
                return reference.Value as PdfDictionary;
            return item as PdfDictionary;
        }
        private PdfDictionary? InjectHelveticaToAcroForm(PdfDocument doc)
        {
            var catalog = doc.Internals.Catalog;
            var acroForm = GetPdfDictionary(catalog, "/AcroForm");
            if (acroForm == null)
            {
                acroForm = new PdfDictionary(doc);
                catalog.Elements["/AcroForm"] = acroForm;
            }
            var dr = GetPdfDictionary(acroForm, "/DR");
            if (dr == null)
            {
                dr = new PdfDictionary(doc);
                acroForm.Elements["/DR"] = dr;
            }
            var fontDict = GetPdfDictionary(dr, "/Font");
            if (fontDict == null)
            {
                fontDict = new PdfDictionary(doc);
                dr.Elements["/Font"] = fontDict;
            }
            if (!fontDict.Elements.ContainsKey("/Helvetica"))
            {
                var helvetica = new PdfDictionary(doc);
                helvetica.Elements["/Type"] = new PdfName("/Font");
                helvetica.Elements["/Subtype"] = new PdfName("/Type1");
                helvetica.Elements["/BaseFont"] = new PdfName("/Helvetica");
                helvetica.Elements["/Encoding"] = new PdfName("/WinAnsiEncoding");
                fontDict.Elements["/Helvetica"] = helvetica;
                return helvetica;
            }
            return GetPdfDictionary(fontDict, "/Helvetica");
        }
        private void InjectHelveticaToPage(PdfPage page)
        {
            var resources = GetPdfDictionary(page, "/Resources");
            if (resources == null)
            {
                resources = new PdfDictionary(page.Owner);
                page.Elements["/Resources"] = resources;
            }
            var fontDict = GetPdfDictionary(resources, "/Font");
            if (fontDict == null)
            {
                fontDict = new PdfDictionary(page.Owner);
                resources.Elements["/Font"] = fontDict;
            }
        }
        public void LoadBookmarks(PdfDocumentModel model)
        {
            try
            {
                if (!File.Exists(model.FilePath))
                    return;
                using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Import))
                {
                    Application.Current.Dispatcher.Invoke(() => model.Bookmarks.Clear());
                    foreach (var ol in doc.Outlines)
                    {
                        var vm = ConvertOutlineToViewModel(ol, doc, null);
                        if (vm != null)
                            Application.Current.Dispatcher.Invoke(() => model.Bookmarks.Add(vm));
                    }
                }
            }
            catch { }
        }
        private PdfBookmarkViewModel? ConvertOutlineToViewModel(PdfOutline outline, PdfDocument doc, PdfBookmarkViewModel? parent)
        {
            int pIdx = 0;
            if (outline.DestinationPage != null)
            {
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    if (doc.Pages[i].Equals(outline.DestinationPage))
                    {
                        pIdx = i;
                        break;
                    }
                }
            }
            var vm = new PdfBookmarkViewModel { Title = outline.Title, PageIndex = pIdx, Parent = parent, IsExpanded = true };
            foreach (var child in outline.Outlines)
            {
                var cvm = ConvertOutlineToViewModel(child, doc, vm);
                if (cvm != null)
                    vm.Children.Add(cvm);
            }
            return vm;
        }

        // 정식 명칭 사용
        public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
        {
            public GenericPdfAnnotation(PdfDocument document) : base(document) { }
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
            public double PdfPageWidthPoint
            {
                get; set;
            }
            public double PdfPageHeightPoint
            {
                get; set;
            }
            public double CropX
            {
                get; set;
            }
            public double CropY
            {
                get; set;
            }
            public double CropHeightPoint
            {
                get; set;
            }
            public string? OriginalFilePath
            {
                get; set;
            }
            public int OriginalPageIndex
            {
                get; set;
            }
            public int Rotation
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
            public string? FontFamily
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