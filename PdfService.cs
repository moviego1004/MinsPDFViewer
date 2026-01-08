using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

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

                double defaultW = 0, defaultH = 0;
                lock (PdfiumLock)
                {
                    if (model.DocReader != null)
                    {
                        using (var pr = model.DocReader.GetPageReader(0))
                        {
                            defaultW = pr.GetPageWidth();
                            defaultH = pr.GetPageHeight();
                        }
                    }
                }

                var tempPageList = new List<PdfPageViewModel>();
                for (int i = 0; i < pageCount; i++)
                {
                    if (model.IsDisposed)
                        return;
                    tempPageList.Add(new PdfPageViewModel
                    {
                        PageIndex = i,
                        OriginalFilePath = model.FilePath,
                        OriginalPageIndex = i,
                        Width = defaultW,
                        Height = defaultH,
                        PdfPageWidthPoint = defaultW,
                        PdfPageHeightPoint = defaultH,
                        CropWidthPoint = defaultW,
                        CropHeightPoint = defaultH
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

            _ = Task.Run(async () =>
            {
                try
                {
                    int cnt = model.Pages.Count;
                    for (int i = 0; i < cnt; i++)
                    {
                        if (model.IsDisposed || model.DocReader == null)
                            break;
                        if (i == 0)
                            continue;

                        await Task.Delay(10);
                        double realW = 0, realH = 0;
                        lock (PdfiumLock)
                        {
                            if (model.IsDisposed || model.DocReader == null)
                                break;
                            using (var pr = model.DocReader.GetPageReader(i))
                            {
                                realW = pr.GetPageWidth();
                                realH = pr.GetPageHeight();
                            }
                        }
                        if (!model.IsDisposed)
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                if (i < model.Pages.Count && !model.IsDisposed)
                                {
                                    var vm = model.Pages[i];
                                    if (Math.Abs(vm.Width - realW) > 0.1)
                                    {
                                        vm.Width = realW;
                                        vm.Height = realH;
                                        vm.PdfPageWidthPoint = realW;
                                        vm.PdfPageHeightPoint = realH;
                                        vm.CropWidthPoint = realW;
                                        vm.CropHeightPoint = realH;
                                    }
                                }
                            });
                        }
                    }
                }
                catch { }
            });
        }

        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage)
                return;
            if (pageVM.ImageSource != null)
            {
                LoadAnnotationsLazy(model, pageVM);
                return;
            }

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
                    var bmp = BitmapSource.Create(w, h, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null, imgBytes, w * 4);
                    bmp.Freeze();
                    pageVM.ImageSource = bmp;
                });
            }

            LoadAnnotationsLazy(model, pageVM);
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

        private void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (pageVM.Annotations.Count > 0 || pageVM.IsBlankPage)
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
                    using (var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                    {
                        if (pageVM.OriginalPageIndex < doc.PageCount)
                        {
                            var p = doc.Pages[pageVM.OriginalPageIndex];
                            extracted = ExtractAnnotationsFromPage(p);
                        }
                    }

                    if (extracted != null && extracted.Count > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                                foreach (var ann in extracted)
                                    pageVM.Annotations.Add(ann);
                        });
                    }
                }
                catch { }
            });
        }

        // =========================================================
        // 4. 저장 (Scale 보정 + 폰트 주입 + 모든 에러 해결)
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
                PdfPageWidthPoint = p.PdfPageWidthPoint,
                PdfPageHeightPoint = p.PdfPageHeightPoint,
                CropX = p.CropX,
                CropY = p.CropY,
                CropHeightPoint = p.CropHeightPoint,
                OriginalFilePath = p.OriginalFilePath,
                OriginalPageIndex = p.OriginalPageIndex,
                Rotation = p.Rotation,
                Annotations = p.Annotations.Select(a => new AnnotationSaveData
                {
                    Type = a.Type,
                    X = a.X,
                    Y = a.Y,
                    Width = a.Width,
                    Height = a.Height,
                    TextContent = a.TextContent,
                    FontSize = a.FontSize,
                    FontFamily = a.FontFamily,
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
                File.Copy(model.FilePath, tempFilePath, true);

                using (var doc = PdfReader.Open(tempFilePath, PdfDocumentOpenMode.Modify))
                {
                    // 1. AcroForm 폰트 리소스
                    var helveticaDict = InjectHelveticaToAcroForm(doc);

                    foreach (var pageData in pagesSnapshot)
                    {
                        if (pageData.OriginalPageIndex >= 0 && pageData.OriginalPageIndex < doc.PageCount)
                        {
                            var pdfPage = doc.Pages[pageData.OriginalPageIndex];

                            // 2. 페이지 폰트 리소스
                            if (helveticaDict != null)
                                InjectHelveticaToPage(pdfPage, helveticaDict);

                            // 3. 스케일 계산 (UI -> PDF)
                            double actualPdfWidth = pdfPage.Width.Point;
                            double actualPdfHeight = pdfPage.Height.Point;
                            double scaleX = (pageData.Width > 0) ? actualPdfWidth / pageData.Width : 1.0;
                            double scaleY = (pageData.Height > 0) ? actualPdfHeight / pageData.Height : 1.0;

                            // 4. 삭제 로직 (CS8121 에러 해결: 패턴 매칭 제거)
                            if (pdfPage.Annotations != null)
                            {
                                for (int k = pdfPage.Annotations.Count - 1; k >= 0; k--)
                                {
                                    // Indexer returns PdfAnnotation directly
                                    var annot = pdfPage.Annotations[k];
                                    if (annot == null)
                                        continue;

                                    string subtype = "";
                                    if (annot.Elements.ContainsKey("/Subtype"))
                                        subtype = annot.Elements.GetString("/Subtype");

                                    if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                    {
                                        pdfPage.Annotations.Elements.RemoveAt(k);
                                    }
                                }
                            }

                            // 5. 새 주석 추가
                            foreach (var ann in pageData.Annotations)
                            {
                                if (ann.Type == AnnotationType.SignaturePlaceholder)
                                    continue;

                                // 스케일 적용
                                double pdfW = ann.Width * scaleX;
                                double pdfH = ann.Height * scaleY;
                                double pdfX = ann.X * scaleX;
                                double pdfY = actualPdfHeight - (ann.Y * scaleY) - pdfH;

                                var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));
                                var newPdfAnn = new GenericPdfAnnotation(doc);
                                newPdfAnn.Rectangle = rect;
                                newPdfAnn.Elements["/F"] = new PdfInteger(4);

                                if (ann.Type == AnnotationType.FreeText)
                                {
                                    newPdfAnn.Elements["/Subtype"] = new PdfName("/FreeText");
                                    newPdfAnn.Contents = ann.TextContent;

                                    string r = (ann.ForeR / 255.0).ToString("0.##");
                                    string g = (ann.ForeG / 255.0).ToString("0.##");
                                    string b = (ann.ForeB / 255.0).ToString("0.##");

                                    double scaledFontSize = ann.FontSize * scaleY;
                                    newPdfAnn.Elements["/DA"] = new PdfString($"/Helvetica {scaledFontSize} Tf {r} {g} {b} rg");

                                    // 디버그용 테두리 (확인 후 삭제 가능)
                                    var bs = new PdfDictionary(doc);
                                    bs.Elements["/W"] = new PdfInteger(1);
                                    bs.Elements["/S"] = new PdfName("/S");
                                    newPdfAnn.Elements["/BS"] = bs;

                                    pdfPage.Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                                {
                                    var quadPoints = new PdfArray(doc);
                                    quadPoints.Elements.Add(new PdfReal(pdfX));
                                    quadPoints.Elements.Add(new PdfReal(pdfY + pdfH));
                                    quadPoints.Elements.Add(new PdfReal(pdfX + pdfW));
                                    quadPoints.Elements.Add(new PdfReal(pdfY + pdfH));
                                    quadPoints.Elements.Add(new PdfReal(pdfX));
                                    quadPoints.Elements.Add(new PdfReal(pdfY));
                                    quadPoints.Elements.Add(new PdfReal(pdfX + pdfW));
                                    quadPoints.Elements.Add(new PdfReal(pdfY));

                                    newPdfAnn.Elements["/QuadPoints"] = quadPoints;

                                    if (ann.Type == AnnotationType.Highlight)
                                    {
                                        newPdfAnn.Elements["/Subtype"] = new PdfName("/Highlight");
                                        var colorArr = new PdfArray(doc);
                                        colorArr.Elements.Add(new PdfReal(ann.BackR / 255.0));
                                        colorArr.Elements.Add(new PdfReal(ann.BackG / 255.0));
                                        colorArr.Elements.Add(new PdfReal(ann.BackB / 255.0));
                                        newPdfAnn.Elements["/C"] = colorArr;
                                    }
                                    else
                                    {
                                        newPdfAnn.Elements["/Subtype"] = new PdfName("/Underline");
                                    }

                                    pdfPage.Annotations.Add(newPdfAnn);
                                }
                            }
                        }
                    }

                    var acroForm = GetPdfDictionary(doc.Internals.Catalog, "/AcroForm");
                    if (acroForm == null)
                    {
                        acroForm = new PdfDictionary(doc);
                        doc.Internals.Catalog.Elements["/AcroForm"] = acroForm;
                    }
                    acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);

                    doc.Save(tempFilePath);
                }
            });

            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);

                DebugInspectPdf(outputPath);
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
                model.IsDisposed = false;
            }
        }

        // [Helper] PdfDictionary 안전 추출 (리턴 타입 Nullable)
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
            else
            {
                return GetPdfDictionary(fontDict, "/Helvetica");
            }
        }

        private void InjectHelveticaToPage(PdfPage page, PdfDictionary helveticaDict)
        {
            if (helveticaDict == null)
                return;

            // [수정] page 자체를 전달 (page.Elements 아님) -> CS1503 해결
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

            if (!fontDict.Elements.ContainsKey("/Helvetica"))
            {
                fontDict.Elements["/Helvetica"] = helveticaDict;
            }
        }

        public void DebugInspectPdf(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"=== [PDF 정밀 진단] 파일: {Path.GetFileName(filePath)} ===");
            if (!File.Exists(filePath))
                return;

            try
            {
                using (var doc = PdfReader.Open(filePath, PdfDocumentOpenMode.Import))
                {
                    var acroForm = GetPdfDictionary(doc.Internals.Catalog, "/AcroForm");
                    if (acroForm != null)
                    {
                        var needApp = acroForm.Elements["/NeedAppearances"];
                        System.Diagnostics.Debug.WriteLine($"✅ AcroForm 존재. NeedApp: {needApp}");
                    }

                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        if (page.Annotations != null)
                        {
                            System.Diagnostics.Debug.WriteLine($"📄 [Page {i}] 주석 {page.Annotations.Count}개");

                            // [수정] foreach loop: PdfItem -> PdfDictionary로 안전 변환 후 Elements 접근
                            // 이 부분이 CS1061 에러의 주범이었습니다.
                            foreach (var item in page.Annotations)
                            {

                                // 1. 안전하게 Dictionary로 변환
                                PdfDictionary? ann = item as PdfDictionary;
                                if (ann == null && item is PdfReference r)
                                    ann = r.Value as PdfDictionary;

                                // 2. 변환 성공 시에만 Elements 사용
                                if (ann != null)
                                {
                                    string subtype = "";
                                    if (ann.Elements.ContainsKey("/Subtype"))
                                        subtype = ann.Elements.GetString("/Subtype");

                                    string rect = "N/A";
                                    if (ann.Elements.ContainsKey("/Rect"))
                                        rect = ann.Elements["/Rect"].ToString();

                                    System.Diagnostics.Debug.WriteLine($"   - {subtype} | Rect: {rect}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"❌ 진단 에러: {ex.Message}");
            }
            System.Diagnostics.Debug.WriteLine("=== [진단 종료] ===");
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

        private List<MinsPDFViewer.PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            return new List<MinsPDFViewer.PdfAnnotation>();
        }

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
            public string TextContent { get; set; } = "";
            public double FontSize
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