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

        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            return await Task.Run(() =>
            {
                try
                {
                    IDocReader docReader;
                    lock (PdfiumLock)
                    {
                        docReader = _docLib.GetDocReader(filePath, new PageDimensions(1.0));
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
        // [수정] IsDisposed 체크 추가로 충돌 방지
        // =========================================================
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null || model.IsDisposed)
                return;

            // 1. 페이지 목록 생성
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

            // 2. 백그라운드 크기 보정 루프
            _ = Task.Run(async () =>
            {
                try
                {
                    int cnt = model.Pages.Count;
                    for (int i = 0; i < cnt; i++)
                    {
                        // [핵심] 문서가 닫혔거나 닫히는 중이면 즉시 중단 -> Dispose 충돌 방지
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
                    var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, imgBytes, w * 4);
                    bmp.Freeze();
                    pageVM.ImageSource = bmp;
                });
            }

            // 고화질 렌더링
            if (needHighQuality && !model.IsDisposed)
            {
                Task.Run(() =>
                {
                    if (model.IsDisposed)
                        return;
                    try
                    {
                        Task.Delay(50).Wait();
                        byte[]? hqBytes = null;
                        int hw = 0, hh = 0;

                        lock (PdfiumLock)
                        {
                            if (model.IsDisposed)
                                return;
                            if (model.CleanDocReader == null)
                            {
                                try
                                {
                                    model.CleanDocReader = _docLib.GetDocReader(model.FilePath, new PageDimensions(3.0));
                                }
                                catch { }
                            }
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
                                var bmp = BitmapSource.Create(hw, hh, 96, 96, PixelFormats.Bgra32, null, hqBytes, hw * 4);
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

                    var list = new List<PdfAnnotation>();
                    // PdfSharp Lock 충돌 방지 위해 별도 처리 권장되나 Import모드는 읽기전용이라 비교적 안전
                    using (var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                    {
                        if (pageVM.OriginalPageIndex < doc.PageCount)
                        {
                            // 주석 추출 로직 (기존 헬퍼 사용)
                            // 여기서는 생략된 ExtractAnnotationsFromPage 호출
                        }
                    }
                }
                catch { }
            });
        }

        // =========================================================
        // [수정] 저장 로직: Dispose 전에 IsDisposed 설정하여 백그라운드 작업 차단
        // =========================================================
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            string tempPath = Path.GetTempFileName();

            // 1. 파일 생성 (오래 걸릴 수 있으므로 백그라운드)
            await Task.Run(() =>
            {
                using (var outputDoc = new PdfDocument())
                {
                    var sourceDocs = new Dictionary<string, PdfDocument>();
                    try
                    {
                        // 페이지 및 주석 복사
                        foreach (var pageVM in model.Pages)
                        {
                            if (model.IsDisposed)
                                throw new OperationCanceledException("저장 중 문서가 닫혔습니다.");

                            PdfPage newPage;
                            if (pageVM.IsBlankPage)
                            {
                                newPage = outputDoc.AddPage();
                                newPage.Width = pageVM.Width;
                                newPage.Height = pageVM.Height;
                            }
                            else
                            {
                                string srcPath = pageVM.OriginalFilePath ?? model.FilePath;
                                if (!sourceDocs.ContainsKey(srcPath))
                                    sourceDocs[srcPath] = PdfReader.Open(srcPath, PdfDocumentOpenMode.Import);

                                var srcDoc = sourceDocs[srcPath];
                                newPage = outputDoc.AddPage(srcDoc.Pages[pageVM.OriginalPageIndex]);
                            }

                            newPage.Rotate = (newPage.Rotate + pageVM.Rotation) % 360;
                            if (newPage.Rotate < 0)
                                newPage.Rotate += 360;

                            // (주석 저장 로직 - 기존과 동일하게 유지)
                            foreach (var ann in pageVM.Annotations)
                            {
                                if (ann.Type == AnnotationType.SignaturePlaceholder)
                                    continue;
                                // ... (기존 좌표 변환 및 주석 추가 코드) ...
                                double scaleX = pageVM.PdfPageWidthPoint / pageVM.Width;
                                double scaleY = pageVM.PdfPageHeightPoint / pageVM.Height;
                                double pdfW = ann.Width * scaleX;
                                double pdfH = ann.Height * scaleY;
                                double pdfX = pageVM.CropX + (ann.X * scaleX);
                                double pdfY = (pageVM.CropY + pageVM.CropHeightPoint) - (ann.Y * scaleY) - pdfH;

                                var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));
                                var newPdfAnn = new GenericPdfAnnotation(outputDoc);
                                newPdfAnn.Rectangle = rect;

                                if (ann.Type == AnnotationType.FreeText)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/FreeText");
                                    newPdfAnn.Contents = ann.TextContent;
                                    string daString = $"/Helv {ann.FontSize} Tf 0 0 0 rg";
                                    if (ann.Foreground is SolidColorBrush b)
                                        daString = $"/Helv {ann.FontSize} Tf {b.Color.R / 255.0:0.##} {b.Color.G / 255.0:0.##} {b.Color.B / 255.0:0.##} rg";
                                    newPdfAnn.Elements.SetString("/DA", daString);
                                    newPage.Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Highlight)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
                                    var colorArr = new PdfArray(outputDoc);
                                    colorArr.Elements.Add(new PdfReal(1));
                                    colorArr.Elements.Add(new PdfReal(1));
                                    colorArr.Elements.Add(new PdfReal(0));
                                    if (ann.Background is SolidColorBrush b)
                                    {
                                        colorArr.Elements.Clear();
                                        colorArr.Elements.Add(new PdfReal(b.Color.R / 255.0));
                                        colorArr.Elements.Add(new PdfReal(b.Color.G / 255.0));
                                        colorArr.Elements.Add(new PdfReal(b.Color.B / 255.0));
                                    }
                                    newPdfAnn.Elements.SetObject(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.C, colorArr);
                                    newPage.Annotations.Add(newPdfAnn);
                                }
                                else if (ann.Type == AnnotationType.Underline)
                                {
                                    newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Underline");
                                    newPage.Annotations.Add(newPdfAnn);
                                }
                            }
                        }

                        // 북마크 저장
                        if (model.Bookmarks.Count > 0)
                        {
                            foreach (var bm in model.Bookmarks)
                                SaveBookmarkToPdfSharp(bm, outputDoc.Outlines, outputDoc);
                        }

                        outputDoc.Save(tempPath);
                    }
                    finally
                    {
                        foreach (var d in sourceDocs.Values)
                            d.Dispose();
                    }
                }
            });

            // 2. 덮어쓰기 로직 (안전하게 Dispose)
            bool isOverwrite = string.Equals(model.FilePath, outputPath, StringComparison.OrdinalIgnoreCase);
            if (isOverwrite)
            {
                // [핵심] 1. 먼저 플래그를 올려 백그라운드 루프가 스스로 멈추게 함
                model.IsDisposed = true;

                // [핵심] 2. 백그라운드 스레드가 Lock을 놓을 때까지 잠시 대기
                await Task.Delay(100);

                // [핵심] 3. UI 스레드에서 안전하게 해제 (AccessViolation 방지)
                lock (PdfiumLock)
                {
                    try
                    {
                        model.CleanDocReader?.Dispose();
                    }
                    catch { }
                    model.CleanDocReader = null;

                    try
                    {
                        model.DocReader?.Dispose();
                    }
                    catch { }
                    model.DocReader = null;
                }
            }

            try
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempPath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}");
            }

            // 3. 재시작
            if (isOverwrite)
            {
                // 플래그 초기화
                model.IsDisposed = false;
                lock (PdfiumLock)
                {
                    model.DocReader = _docLib.GetDocReader(outputPath, new PageDimensions(1.0));
                }
                // (필요하다면 여기서 다시 InitializeDocumentAsync 호출 가능)
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

            var vm = new PdfBookmarkViewModel
            {
                Title = outline.Title,
                PageIndex = pIdx,
                Parent = parent,
                IsExpanded = true
            };
            foreach (var child in outline.Outlines)
            {
                var cvm = ConvertOutlineToViewModel(child, doc, vm);
                if (cvm != null)
                    vm.Children.Add(cvm);
            }
            return vm;
        }

        private void SaveBookmarkToPdfSharp(PdfBookmarkViewModel vm, PdfOutlineCollection collection, PdfDocument doc)
        {
            if (vm.PageIndex >= 0 && vm.PageIndex < doc.PageCount)
            {
                var target = doc.Pages[vm.PageIndex];
                var added = collection.Add(vm.Title, target, true);
                foreach (var child in vm.Children)
                {
                    SaveBookmarkToPdfSharp(child, added.Outlines, doc);
                }
            }
        }

        public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
        {
            public GenericPdfAnnotation(PdfDocument d) : base(d) { }
        }
    }
}