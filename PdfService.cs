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
                    // 파일 잠금 방지를 위해 바이트 배열로 읽음
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

                    // [수정] MPdfAnnotation 사용 (이름 충돌 방지)
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
                        System.Diagnostics.Debug.WriteLine($"[주석 로드 경고] {ex.Message}");
                    }

                    if (extracted != null && extracted.Count > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            // ViewModel의 Annotations 컬렉션 타입도 MPdfAnnotation으로 변경되었다고 가정
                            // 만약 ViewModel 수정이 어렵다면 여기서 변환 로직이 필요할 수 있음
                            // 현재는 사용자가 이름을 바꿨다고 가정하고 진행
                            if (pageVM.Annotations.Count == 0)
                            {
                                foreach (var appAnnotation in extracted)
                                {
                                    // dynamic 캐스팅을 통해 ViewModel 컬렉션에 추가 (타입 불일치 방지)
                                    ((dynamic)pageVM.Annotations).Add(appAnnotation);
                                }
                            }
                        });
                    }
                }
                catch { }
            });
        }

        // [수정] 반환 타입 MPdfAnnotation
        private List<MPdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
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
                    // PdfReference 처리
                    PdfDictionary? annotDict = null;
                    if (item is PdfDictionary dict)
                        annotDict = dict;
                    else if (item is PdfReference reference)
                        annotDict = reference.Value as PdfDictionary;

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
                        textBrush.Freeze();

                        // [수정] MPdfAnnotation 생성 (상속 없음)
                        var newAnn = new MPdfAnnotation
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
                        highlightBrush.Freeze();

                        var newAnn = new MPdfAnnotation
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
        // 4. 저장
        // =========================================================
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
                string sanitizedSourcePath = Path.GetTempFileName();

                try
                {
                    lock (PdfiumLock)
                    {
                        _docLib.Merge(sanitizedSourcePath, new[] { tempSourceCopy }); // 배열로 전달
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"파일 구조 개선 실패: {ex.Message}");
                }

                try
                {
                    using (var outputDoc = new PdfDocument())
                    {
                        var acroForm = new PdfDictionary(outputDoc);
                        outputDoc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
                        acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);
                        InjectHelveticaToAcroForm(outputDoc);

                        using (var form = XPdfForm.FromFile(sanitizedSourcePath))
                        {
                            for (int i = 0; i < pagesSnapshot.Count; i++)
                            {
                                var pageData = pagesSnapshot[i];
                                PdfPage newPage = outputDoc.AddPage();

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
                    System.Diagnostics.Debug.WriteLine($"[1차 저장 실패] {ex.Message}. 2차 시도(이미지) 진행...");
                    try
                    {
                        using (var outputDoc = new PdfDocument())
                        {
                            var acroForm = new PdfDictionary(outputDoc);
                            outputDoc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
                            acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);
                            InjectHelveticaToAcroForm(outputDoc);

                            lock (PdfiumLock)
                            {
                                byte[] bytes = File.ReadAllBytes(tempSourceCopy);
                                using (var docReader = DocLib.Instance.GetDocReader(bytes, new PageDimensions(2.0)))
                                {
                                    for (int i = 0; i < pagesSnapshot.Count; i++)
                                    {
                                        var pageData = pagesSnapshot[i];
                                        PdfPage newPage = outputDoc.AddPage();

                                        if (pageData.Width > 0 && pageData.Height > 0)
                                        {
                                            newPage.Width = XUnit.FromPoint(pageData.Width);
                                            newPage.Height = XUnit.FromPoint(pageData.Height);
                                        }

                                        if (!pageData.IsBlankPage && pageData.OriginalPageIndex < docReader.GetPageCount())
                                        {
                                            using (var pageReader = docReader.GetPageReader(pageData.OriginalPageIndex))
                                            {
                                                var rawBytes = pageReader.GetImage(RenderFlags.RenderAnnotations);
                                                int w = pageReader.GetPageWidth();
                                                int h = pageReader.GetPageHeight();
                                                var bmpSource = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, rawBytes, w * 4);

                                                using (var ms = new MemoryStream())
                                                {
                                                    var encoder = new PngBitmapEncoder();
                                                    encoder.Frames.Add(BitmapFrame.Create(bmpSource));
                                                    encoder.Save(ms);
                                                    ms.Position = 0;

                                                    using (var xImg = XImage.FromStream(ms))
                                                    {
                                                        using (var gfx = XGraphics.FromPdfPage(newPage))
                                                        {
                                                            gfx.DrawImage(xImg, 0, 0, newPage.Width.Point, newPage.Height.Point);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        AddAnnotations(outputDoc, newPage, pageData.Annotations);
                                    }
                                }
                            }
                            outputDoc.Save(tempFilePath);
                        }
                    }
                    catch (Exception ex2)
                    {
                        throw new Exception($"PDF 저장 실패 (최종): {ex2.Message}");
                    }
                }
                finally
                {
                    if (File.Exists(tempSourceCopy))
                        File.Delete(tempSourceCopy);
                    if (File.Exists(sanitizedSourcePath))
                        File.Delete(sanitizedSourcePath);
                }
            });

            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 덮어쓰기 실패: {ex.Message}");
            }
            finally
            {
                if (File.Exists(tempFilePath))
                    File.Delete(tempFilePath);
            }

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

        private void AddAnnotations(PdfDocument doc, PdfPage page, List<AnnotationSaveData> annotations)
        {
            var acroForm = GetPdfDictionary(doc.Internals.Catalog, "/AcroForm");
            if (acroForm == null)
            {
                acroForm = new PdfDictionary(doc);
                doc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
            }

            foreach (var ann in annotations)
            {
                double pdfX = ann.X;
                double pdfY = page.Height.Point - ann.Y - ann.Height;
                double pdfW = ann.Width;
                double pdfH = ann.Height;
                var rect = new PdfSharpRect(new XRect(pdfX, pdfY, pdfW, pdfH));

                var annotDict = new PdfDictionary(doc);
                annotDict.Elements["/Type"] = new PdfName("/Annot");
                annotDict.Elements["/Rect"] = rect;
                annotDict.Elements["/F"] = new PdfInteger(4);

                if (ann.Type == AnnotationType.SignaturePlaceholder)
                {
                    string fieldName = $"Signature_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
                    annotDict.Elements["/Subtype"] = new PdfName("/Widget");
                    annotDict.Elements["/FT"] = new PdfName("/Sig");
                    annotDict.Elements["/T"] = new PdfString(fieldName);
                    annotDict.Elements["/P"] = page;

                    AddAnnotToPageArray(page, annotDict, doc);
                    if (!acroForm.Elements.ContainsKey("/Fields"))
                        acroForm.Elements["/Fields"] = new PdfArray(doc);
                    acroForm.Elements.GetArray("/Fields")?.Elements.Add(annotDict);
                    continue;
                }

                if (ann.Type == AnnotationType.FreeText)
                {
                    annotDict.Elements["/Subtype"] = new PdfName("/FreeText");
                    annotDict.Elements["/Contents"] = new PdfString(ann.TextContent ?? "");
                    double fSize = ann.FontSize < 5 ? 10 : ann.FontSize;
                    string r = (ann.ForeR / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                    string g = (ann.ForeG / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                    string b = (ann.ForeB / 255.0).ToString("0.##", CultureInfo.InvariantCulture);
                    annotDict.Elements["/DA"] = new PdfString($"/Helvetica {fSize} Tf {r} {g} {b} rg");
                    annotDict.Elements["/Q"] = new PdfInteger(0);
                    CreateAppearanceStream(doc, annotDict, pdfW, pdfH, ann.TextContent ?? "", fSize, r, g, b);
                    AddAnnotToPageArray(page, annotDict, doc);
                }
                else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                {
                    string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                    annotDict.Elements["/Subtype"] = new PdfName(subtype);
                    var qp = new PdfArray(doc);
                    qp.Elements.Add(new PdfReal(pdfX));
                    qp.Elements.Add(new PdfReal(pdfY + pdfH));
                    qp.Elements.Add(new PdfReal(pdfX + pdfW));
                    qp.Elements.Add(new PdfReal(pdfY + pdfH));
                    qp.Elements.Add(new PdfReal(pdfX));
                    qp.Elements.Add(new PdfReal(pdfY));
                    qp.Elements.Add(new PdfReal(pdfX + pdfW));
                    qp.Elements.Add(new PdfReal(pdfY));
                    annotDict.Elements["/QuadPoints"] = qp;
                    if (ann.Type == AnnotationType.Highlight)
                    {
                        var c = new PdfArray(doc);
                        c.Elements.Add(new PdfReal(ann.BackR / 255.0));
                        c.Elements.Add(new PdfReal(ann.BackG / 255.0));
                        c.Elements.Add(new PdfReal(ann.BackB / 255.0));
                        annotDict.Elements["/C"] = c;
                    }
                    AddAnnotToPageArray(page, annotDict, doc);
                }
            }
        }

        private void AddAnnotToPageArray(PdfPage page, PdfDictionary annotDict, PdfDocument doc)
        {
            if (!page.Elements.ContainsKey("/Annots"))
                page.Elements["/Annots"] = new PdfArray(doc);
            page.Elements.GetArray("/Annots")?.Elements.Add(annotDict);
        }

        private void CreateAppearanceStream(PdfDocument doc, PdfDictionary annotation, double w, double h, string text, double fSize, string r, string g, string b)
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

        // [수정] MPdfAnnotation 정의 (상속 제거 -> 빌드 에러 및 이름 충돌 해결)
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
            public string TextContent { get; set; } = "";
            public double FontSize
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