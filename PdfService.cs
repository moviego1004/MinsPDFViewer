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
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private readonly IDocLib _docLib;

        // [설정] 단일 리더 품질 배율 (2.0 = 속도와 화질의 균형)
        private const double RENDER_SCALE = 2.0;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        // =========================================================
        // 1. PDF 로드 (단일 리더, 메모리 모드, 파일 잠금 없음)
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
                double defaultW = 0, defaultH = 0;

                lock (PdfiumLock)
                {
                    if (model.DocReader == null)
                        return;
                    pageCount = model.DocReader.GetPageCount();
                    if (pageCount > 0)
                    {
                        using (var pr = model.DocReader.GetPageReader(0))
                        {
                            defaultW = pr.GetPageWidth();
                            defaultH = pr.GetPageHeight();
                        }
                    }
                }

                if (pageCount == 0)
                    return;

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

        // =========================================================
        // 3. 렌더링
        // =========================================================
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
        // 4. 저장 (NeedAppearances 수정 및 안정화)
        // =========================================================
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            // [Step 1] 데이터 스냅샷
            var bookmarkList = model.Bookmarks.ToList();
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
                    ForeR = (a.Foreground as SolidColorBrush)?.Color.R ?? 0,
                    ForeG = (a.Foreground as SolidColorBrush)?.Color.G ?? 0,
                    ForeB = (a.Foreground as SolidColorBrush)?.Color.B ?? 0,
                    BackR = (a.Background as SolidColorBrush)?.Color.R ?? 255,
                    BackG = (a.Background as SolidColorBrush)?.Color.G ?? 255,
                    BackB = (a.Background as SolidColorBrush)?.Color.B ?? 255,
                    IsHighlight = (a.Type == AnnotationType.Highlight)
                }).ToList()
            }).ToList();

            bool isOverwriting = string.Equals(model.FilePath, outputPath, StringComparison.OrdinalIgnoreCase);

            // [Step 2] 원본 읽기
            byte[] sourceFileBytes = ReadFileSafely(model.FilePath);

            // [Step 3] PDF 생성 및 저장
            byte[] resultBytes = await Task.Run(() =>
            {
                using (var sourceStream = new MemoryStream(sourceFileBytes))
                using (var outputStream = new MemoryStream())
                using (var outputDoc = new PdfDocument())
                {
                    var sourceDocs = new Dictionary<string, PdfDocument>();
                    sourceDocs[model.FilePath] = PdfReader.Open(sourceStream, PdfDocumentOpenMode.Import);

                    try
                    {
                        foreach (var pageData in pagesSnapshot)
                        {
                            PdfPage newPage;
                            if (pageData.IsBlankPage)
                            {
                                newPage = outputDoc.AddPage();
                                newPage.Width = XUnit.FromPoint(pageData.Width);
                                newPage.Height = XUnit.FromPoint(pageData.Height);
                            }
                            else
                            {
                                string srcPath = pageData.OriginalFilePath ?? model.FilePath;
                                int srcIndex = pageData.OriginalPageIndex;

                                if (!sourceDocs.ContainsKey(srcPath))
                                    sourceDocs[srcPath] = PdfReader.Open(srcPath, PdfDocumentOpenMode.Import);

                                var srcDoc = sourceDocs[srcPath];
                                if (srcIndex >= 0 && srcIndex < srcDoc.PageCount)
                                    newPage = outputDoc.AddPage(srcDoc.Pages[srcIndex]);
                                else
                                    continue;
                            }

                            newPage.Rotate = (newPage.Rotate + pageData.Rotation) % 360;
                            if (newPage.Rotate < 0)
                                newPage.Rotate += 360;

                            foreach (var ann in pageData.Annotations)
                            {
                                if (ann.Type == AnnotationType.SignaturePlaceholder)
                                    continue;

                                double scaleX = (pageData.PdfPageWidthPoint > 0) ? pageData.PdfPageWidthPoint / pageData.Width : 1.0;
                                double scaleY = (pageData.PdfPageHeightPoint > 0) ? pageData.PdfPageHeightPoint / pageData.Height : 1.0;

                                double pdfW = ann.Width * scaleX;
                                double pdfH = ann.Height * scaleY;
                                double pdfX = pageData.CropX + (ann.X * scaleX);

                                double cropHeight = pageData.CropHeightPoint > 0 ? pageData.CropHeightPoint : pageData.PdfPageHeightPoint;
                                if (cropHeight == 0)
                                    cropHeight = pageData.Height;
                                double pdfY = (pageData.CropY + cropHeight) - (ann.Y * scaleY) - pdfH;

                                var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));
                                var newPdfAnn = new GenericPdfAnnotation(outputDoc);
                                newPdfAnn.Rectangle = rect;

                                if (ann.Type == AnnotationType.FreeText)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/FreeText");
                                    newPdfAnn.Contents = ann.TextContent;
                                    string r = (ann.ForeR / 255.0).ToString("0.##");
                                    string g = (ann.ForeG / 255.0).ToString("0.##");
                                    string b = (ann.ForeB / 255.0).ToString("0.##");
                                    newPdfAnn.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r} {g} {b} rg");
                                    outputDoc.Pages[outputDoc.PageCount - 1].Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Highlight)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
                                    var colorArr = new PdfArray(outputDoc);
                                    colorArr.Elements.Add(new PdfReal(ann.BackR / 255.0));
                                    colorArr.Elements.Add(new PdfReal(ann.BackG / 255.0));
                                    colorArr.Elements.Add(new PdfReal(ann.BackB / 255.0));
                                    newPdfAnn.Elements.SetObject(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.C, colorArr);
                                    outputDoc.Pages[outputDoc.PageCount - 1].Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Underline)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Underline");
                                    outputDoc.Pages[outputDoc.PageCount - 1].Annotations.Add(newPdfAnn);
                                }
                            }
                        }

                        if (bookmarkList.Count > 0)
                        {
                            foreach (var bm in bookmarkList)
                                SaveBookmarkToPdfSharp(bm, outputDoc.Outlines, outputDoc);
                        }

                        // [수정] AcroForm 프로퍼티 접근 에러 해결
                        // 프로퍼티(outputDoc.AcroForm) 대신 내부 Catalog에 직접 접근하여 사전(Dictionary)을 생성합니다.
                        var catalog = outputDoc.Internals.Catalog;
                        PdfDictionary acroFormDict;

                        if (catalog.Elements.ContainsKey("/AcroForm"))
                        {
                            acroFormDict = catalog.Elements.GetDictionary("/AcroForm");
                        }
                        else
                        {
                            // 없으면 수동으로 만들어서 집어넣습니다. (이러면 에러 안 남)
                            acroFormDict = new PdfDictionary(outputDoc);
                            catalog.Elements["/AcroForm"] = acroFormDict;
                        }

                        // NeedAppearances 플래그 설정 (이게 있어야 텍스트 박스가 보임)
                        if (acroFormDict.Elements.ContainsKey("/NeedAppearances"))
                        {
                            acroFormDict.Elements["/NeedAppearances"] = new PdfBoolean(true);
                        }
                        else
                        {
                            acroFormDict.Elements.Add("/NeedAppearances", new PdfBoolean(true));
                        }

                        outputDoc.Save(outputStream);
                        return outputStream.ToArray();
                    }
                    finally
                    {
                        foreach (var d in sourceDocs.Values)
                            d.Dispose();
                    }
                }
            });

            // [Step 4] 파일 덮어쓰기
            try
            {
                File.WriteAllBytes(outputPath, resultBytes);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}");
            }

            // [Step 5] 재로드
            if (isOverwriting)
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

        // 헬퍼 메서드들
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

        private void SaveBookmarkToPdfSharp(PdfBookmarkViewModel vm, PdfOutlineCollection parentCollection, PdfDocument doc)
        {
            if (vm.PageIndex >= 0 && vm.PageIndex < doc.PageCount)
            {
                var target = doc.Pages[vm.PageIndex];
                var added = parentCollection.Add(vm.Title, target, true);
                foreach (var child in vm.Children)
                {
                    SaveBookmarkToPdfSharp(child, added.Outlines, doc);
                }
            }
        }

        private List<MinsPDFViewer.PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            return new List<MinsPDFViewer.PdfAnnotation>();
        }

        // DTO & Annotation Classes
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