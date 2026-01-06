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

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        // =========================================================
        // 1. PDF 로드 (핵심: 메인 리더도 메모리로 로드 -> 파일 잠금 없음)
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
                    IDocReader? cleanReader = null;

                    // [핵심] 파일을 메모리로 미리 다 읽어옵니다.
                    byte[] fileBytes = ReadFileSafely(filePath);

                    lock (PdfiumLock)
                    {
                        // UI 스레드에서 생성 (안전성 확보)
                        // 파일 경로 대신 'fileBytes'를 사용하여 생성 -> 파일 Lock이 걸리지 않음!
                        docReader = Application.Current.Dispatcher.Invoke(() =>
                            _docLib.GetDocReader(fileBytes, new PageDimensions(1.0)));

                        cleanReader = Application.Current.Dispatcher.Invoke(() =>
                            _docLib.GetDocReader(fileBytes, new PageDimensions(3.0)));
                    }

                    var model = new PdfDocumentModel
                    {
                        FilePath = filePath,
                        FileName = Path.GetFileName(filePath),
                        DocLib = _docLib,
                        DocReader = docReader,
                        CleanDocReader = cleanReader
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
        // 2. 초기화 (변경 없음)
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
        // 3. 렌더링 (변경 없음)
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
            bool needHighQuality = false;

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
                            needHighQuality = true;
                        }
                    }
                    catch { }
                }
            }

            if (imgBytes != null && !model.IsDisposed)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    // PixelFormats 명시
                    var bmp = BitmapSource.Create(w, h, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null, imgBytes, w * 4);
                    bmp.Freeze();
                    pageVM.ImageSource = bmp;
                });
            }

            if (needHighQuality && !model.IsDisposed)
            {
                Task.Run(() =>
                {
                    try
                    {
                        if (model.IsDisposed)
                            return;
                        Task.Delay(50).Wait();

                        byte[]? hqBytes = null;
                        int hw = 0, hh = 0;

                        lock (PdfiumLock)
                        {
                            if (model.IsDisposed)
                                return;
                            var reader = model.CleanDocReader ?? model.DocReader;
                            if (reader != null)
                            {
                                using (var pr = reader.GetPageReader(pageVM.OriginalPageIndex))
                                {
                                    hqBytes = pr.GetImage(RenderFlags.RenderAnnotations);
                                    hw = pr.GetPageWidth();
                                    hh = pr.GetPageHeight();
                                }
                            }
                        }

                        if (hqBytes != null && !model.IsDisposed)
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                var bmp = BitmapSource.Create(hw, hh, 96, 96, System.Windows.Media.PixelFormats.Bgra32, null, hqBytes, hw * 4);
                                bmp.Freeze();
                                pageVM.ImageSource = bmp;
                            });
                        }
                    }
                    catch { }
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
        // 4. 저장 (크래시 해결의 핵심: Dispose 제거)
        // =========================================================
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            // 1. 데이터 스냅샷 (UI 스레드)
            var bookmarkList = model.Bookmarks.ToList();
            var pagesSnapshot = model.Pages.Select(p => new
            {
                p.IsBlankPage,
                p.Width,
                p.Height,
                p.OriginalFilePath,
                p.OriginalPageIndex,
                p.Rotation,
                p.Annotations
            }).ToList();

            bool isOverwriting = string.Equals(model.FilePath, outputPath, StringComparison.OrdinalIgnoreCase);

            // 2. 덮어쓰기 전 준비
            if (isOverwriting)
            {
                model.IsDisposed = true; // 백그라운드 작업(렌더링 등) 중단 요청
                await Task.Delay(200);   // 작업들이 멈출 때까지 잠시 대기

                lock (PdfiumLock)
                {
                    // [대박 중요] 여기서 Dispose()를 호출하지 않습니다! 
                    // 우리는 메모리 로드 방식을 쓰므로 파일을 잡고 있지 않습니다.
                    // 따라서 Dispose 없이 그냥 참조만 끊어도 파일 덮어쓰기가 가능합니다.
                    // 이렇게 하면 크래시가 발생할 여지 자체가 사라집니다.
                    model.DocReader = null;
                    model.CleanDocReader = null;
                }
            }

            string tempOutputPath = Path.GetTempFileName();

            // 3. 파일 생성 (PdfSharp)
            await Task.Run(() =>
            {
                using (var outputDoc = new PdfDocument())
                {
                    var sourceDocs = new Dictionary<string, PdfDocument>();
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

                                var rect = new PdfRectangle(new XRect(ann.X, ann.Y, ann.Width, ann.Height)); // 단순화된 좌표
                                var newPdfAnn = new GenericPdfAnnotation(outputDoc);
                                newPdfAnn.Rectangle = rect;

                                if (ann.Type == AnnotationType.FreeText)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/FreeText");
                                    newPdfAnn.Contents = ann.TextContent;

                                    string r = "0", g = "0", b = "0";
                                    if (ann.Foreground is SolidColorBrush br)
                                    {
                                        r = (br.Color.R / 255.0).ToString("0.##");
                                        g = (br.Color.G / 255.0).ToString("0.##");
                                        b = (br.Color.B / 255.0).ToString("0.##");
                                    }
                                    newPdfAnn.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r} {g} {b} rg");
                                    outputDoc.Pages[outputDoc.PageCount - 1].Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Highlight)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
                                    var colorArr = new PdfArray(outputDoc);
                                    if (ann.Background is SolidColorBrush br)
                                    {
                                        colorArr.Elements.Add(new PdfReal(br.Color.R / 255.0));
                                        colorArr.Elements.Add(new PdfReal(br.Color.G / 255.0));
                                        colorArr.Elements.Add(new PdfReal(br.Color.B / 255.0));
                                    }
                                    else
                                    {
                                        colorArr.Elements.Add(new PdfReal(1));
                                        colorArr.Elements.Add(new PdfReal(1));
                                        colorArr.Elements.Add(new PdfReal(0));
                                    }
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

                        outputDoc.Save(tempOutputPath);
                    }
                    finally
                    {
                        foreach (var d in sourceDocs.Values)
                            d.Dispose();
                    }
                }
            });

            // 4. 파일 이동 (이제 안전하게 덮어쓰기 가능)
            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempOutputPath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}");
            }

            // 5. 재시작 (덮어쓰기 했다면 다시 로드)
            if (isOverwriting)
            {
                model.IsDisposed = false;

                // 다시 메모리로 읽어서 리더 생성
                byte[] fileBytes = ReadFileSafely(outputPath);

                lock (PdfiumLock)
                {
                    model.DocReader = Application.Current.Dispatcher.Invoke(() =>
                        _docLib.GetDocReader(fileBytes, new PageDimensions(1.0)));

                    model.CleanDocReader = Application.Current.Dispatcher.Invoke(() =>
                        _docLib.GetDocReader(fileBytes, new PageDimensions(3.0)));
                }
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

        // GenericPdfAnnotation 정의
        public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
        {
            public GenericPdfAnnotation(PdfDocument document) : base(document) { }
        }
    }
}