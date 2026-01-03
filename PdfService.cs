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
                    lock (PdfiumLock)
                    {
                        docReader = _docLib.GetDocReader(filePath, new PageDimensions(1.0));
                    }

                    var model = new PdfDocumentModel
                    {
                        FilePath = filePath,
                        FileName = Path.GetFileName(filePath),
                        DocLib = _docLib,
                        DocReader = docReader,
                        CleanDocReader = null
                    };

                    // 책갈피 로드 (별도 메서드 호출 필요)
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
        // 2. 초기화 (전략 2: 선 로딩, 후 보정)
        // =========================================================
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null)
                return;

            await Task.Run(() =>
            {
                IDocReader reader = model.DocReader;
                int pageCount = reader.GetPageCount();
                if (pageCount == 0)
                    return;

                double defaultWidth = 0;
                double defaultHeight = 0;

                lock (PdfiumLock)
                {
                    using (var pr = reader.GetPageReader(0))
                    {
                        defaultWidth = pr.GetPageWidth();
                        defaultHeight = pr.GetPageHeight();
                    }
                }

                var tempPageList = new List<PdfPageViewModel>();

                for (int i = 0; i < pageCount; i++)
                {
                    var pageVM = new PdfPageViewModel
                    {
                        PageIndex = i,
                        // [중요] 편집을 위해 출처 기록
                        OriginalFilePath = model.FilePath,
                        OriginalPageIndex = i,
                        IsBlankPage = false,
                        Rotation = 0,

                        Width = defaultWidth,
                        Height = defaultHeight,
                        ImageSource = null,

                        PdfPageWidthPoint = defaultWidth,
                        PdfPageHeightPoint = defaultHeight,
                        CropX = 0,
                        CropY = 0,
                        CropWidthPoint = defaultWidth,
                        CropHeightPoint = defaultHeight,
                        HasSignature = false
                    };
                    tempPageList.Add(pageVM);
                }

                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var page in tempPageList)
                        model.Pages.Add(page);
                });
            });

            // 백그라운드 크기 보정
            _ = Task.Run(() =>
            {
                try
                {
                    if (model.DocReader == null)
                        return;
                    int cnt = model.Pages.Count;
                    for (int i = 0; i < cnt; i++)
                    {
                        if (model.DocReader == null)
                            break;
                        if (i == 0)
                            continue;

                        double realWidth = 0, realHeight = 0;
                        lock (PdfiumLock)
                        {
                            using (var pr = model.DocReader.GetPageReader(i))
                            {
                                realWidth = pr.GetPageWidth();
                                realHeight = pr.GetPageHeight();
                            }
                        }

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (i < model.Pages.Count)
                            {
                                var vm = model.Pages[i];
                                if (Math.Abs(vm.Width - realWidth) > 0.1 || Math.Abs(vm.Height - realHeight) > 0.1)
                                {
                                    vm.Width = realWidth;
                                    vm.Height = realHeight;
                                    vm.PdfPageWidthPoint = realWidth;
                                    vm.PdfPageHeightPoint = realHeight;
                                    vm.CropWidthPoint = realWidth;
                                    vm.CropHeightPoint = realHeight;
                                }
                            }
                        });
                        if (i % 10 == 0)
                            Task.Delay(1).Wait();
                    }
                }
                catch { }
            });
        }

        // =========================================================
        // 3. 페이지 렌더링 (고화질 독립 리더 방식)
        // =========================================================
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            // 빈 페이지는 렌더링 안 함 (하얀색)
            if (pageVM.IsBlankPage)
                return;

            double renderScale = 3.0;

            if (pageVM.ImageSource == null)
            {
                lock (PdfiumLock)
                {
                    // 원본 소스가 있는 경우에만
                    if (!string.IsNullOrEmpty(pageVM.OriginalFilePath))
                    {
                        // CleanDocReader 관리 로직 (생략 가능하나 성능 위해 유지)
                        // 여기서는 단순화를 위해 모델의 리더를 우선 사용하되, 
                        // 페이지가 다른 파일에서 왔다면(추후 구현) 별도 처리 필요.
                        // 현재는 같은 파일이므로 model.DocReader 사용.

                        var readerToUse = model.CleanDocReader ?? model.DocReader;
                        if (model.CleanDocReader == null && model.FilePath == pageVM.OriginalFilePath)
                        {
                            try
                            {
                                model.CleanDocReader = _docLib.GetDocReader(model.FilePath, new PageDimensions(renderScale));
                            }
                            catch { }
                            readerToUse = model.CleanDocReader ?? model.DocReader;
                        }

                        if (readerToUse != null)
                        {
                            try
                            {
                                using (var pageReader = readerToUse.GetPageReader(pageVM.OriginalPageIndex))
                                {
                                    var w = pageReader.GetPageWidth();
                                    var h = pageReader.GetPageHeight();
                                    var bytes = pageReader.GetImage(RenderFlags.RenderAnnotations);

                                    Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, bytes, w * 4);
                                        bmp.Freeze();
                                        pageVM.ImageSource = bmp;
                                    });
                                }
                            }
                            catch { }
                        }
                    }
                }
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
                    string path = pageVM.OriginalFilePath ?? model.FilePath;
                    if (!File.Exists(path))
                        return;

                    using (var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import))
                    {
                        if (pageVM.OriginalPageIndex < doc.PageCount)
                        {
                            var p = doc.Pages[pageVM.OriginalPageIndex];
                            var extractedAnns = ExtractAnnotationsFromPage(p);

                            if (extractedAnns.Count > 0)
                            {
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    if (pageVM.Annotations.Count > 0)
                                        return;
                                    foreach (var ann in extractedAnns)
                                    {
                                        // 좌표 보정 로직 (생략)
                                        pageVM.Annotations.Add(ann);
                                    }
                                });
                            }
                        }
                    }
                }
                catch { }
            });
        }

        // =========================================================
        // 4. 저장 (기존과 동일)
        // =========================================================
        // public void SavePdf(PdfDocumentModel model, string outputPath)
        // {
        //     if (model == null || string.IsNullOrEmpty(model.FilePath))
        //         return;

        //     string tempPath = Path.GetTempFileName();
        //     File.Copy(model.FilePath, tempPath, true);

        //     try
        //     {
        //         using (var doc = PdfReader.Open(tempPath, PdfDocumentOpenMode.Modify))
        //         {
        //             foreach (var pageVM in model.Pages)
        //             {
        //                 if (pageVM.PageIndex >= doc.PageCount)
        //                     continue;
        //                 var pdfPage = doc.Pages[pageVM.PageIndex];

        //                 if (pdfPage.Annotations != null)
        //                 {
        //                     for (int k = pdfPage.Annotations.Count - 1; k >= 0; k--)
        //                     {
        //                         var annot = pdfPage.Annotations[k];
        //                         if (annot == null)
        //                             continue;
        //                         var subtype = annot.Elements.GetString("/Subtype");
        //                         if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
        //                         {
        //                             pdfPage.Annotations.Elements.RemoveAt(k);
        //                         }
        //                     }
        //                 }

        //                 foreach (MinsPDFViewer.PdfAnnotation ann in pageVM.Annotations)
        //                 {
        //                     if (ann.Type == AnnotationType.SignaturePlaceholder || ann.Type == AnnotationType.SignatureField)
        //                         continue;

        //                     double scaleX = pageVM.PdfPageWidthPoint / pageVM.Width;
        //                     double scaleY = pageVM.PdfPageHeightPoint / pageVM.Height;

        //                     double pdfW = ann.Width * scaleX;
        //                     double pdfH = ann.Height * scaleY;

        //                     double cropX = pageVM.CropX;
        //                     double cropY = pageVM.CropY;
        //                     double cropH = pageVM.CropHeightPoint;

        //                     double pdfX = cropX + (ann.X * scaleX);
        //                     double pdfY = (cropY + cropH) - (ann.Y * scaleY) - pdfH;

        //                     var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));
        //                     var newPdfAnn = new GenericPdfAnnotation(doc);
        //                     newPdfAnn.Rectangle = rect;

        //                     if (ann.Type == AnnotationType.FreeText)
        //                     {
        //                         newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/FreeText");
        //                         newPdfAnn.Contents = ann.TextContent;
        //                         string daString = $"/Helv {ann.FontSize} Tf 0 0 0 rg";
        //                         if (ann.Foreground is SolidColorBrush b)
        //                             daString = $"/Helv {ann.FontSize} Tf {b.Color.R / 255.0:0.##} {b.Color.G / 255.0:0.##} {b.Color.B / 255.0:0.##} rg";
        //                         newPdfAnn.Elements.SetString("/DA", daString);
        //                     }
        //                     else if (ann.Type == AnnotationType.Highlight)
        //                     {
        //                         newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
        //                         var colorArr = new PdfArray(doc);
        //                         if (ann.Background is SolidColorBrush b)
        //                         {
        //                             colorArr.Elements.Add(new PdfReal(b.Color.R / 255.0));
        //                             colorArr.Elements.Add(new PdfReal(b.Color.G / 255.0));
        //                             colorArr.Elements.Add(new PdfReal(b.Color.B / 255.0));
        //                         }
        //                         else
        //                         {
        //                             colorArr.Elements.Add(new PdfReal(1));
        //                             colorArr.Elements.Add(new PdfReal(1));
        //                             colorArr.Elements.Add(new PdfReal(0));
        //                         }
        //                         newPdfAnn.Elements.SetObject(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.C, colorArr);
        //                     }
        //                     else if (ann.Type == AnnotationType.Underline)
        //                     {
        //                         newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Underline");
        //                     }

        //                     pdfPage.Annotations.Add(newPdfAnn);
        //                 }
        //             }

        //             doc.Outlines.Clear();
        //             foreach (var bm in model.Bookmarks)
        //             {
        //                 SaveBookmarkToDocument(bm, doc.Outlines, doc);
        //             }

        //             doc.Save(outputPath);
        //         }
        //     }
        //     catch (Exception ex)
        //     {
        //         if (File.Exists(outputPath))
        //             File.Delete(outputPath);
        //         throw new Exception($"저장 오류: {ex.Message}");
        //     }
        //     finally { if (File.Exists(tempPath)) File.Delete(tempPath); }
        // }

        // =========================================================
        // 4. 저장 (재조립 방식: 회전, 삭제, 추가 반영)
        // =========================================================
        // 4. 저장 (재조립 방식)
        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            string tempOutputPath = Path.GetTempFileName();

            using (var outputDoc = new PdfDocument())
            {
                var sourceDocs = new Dictionary<string, PdfDocument>();

                try
                {
                    foreach (var pageVM in model.Pages)
                    {
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
                            int srcIndex = pageVM.OriginalPageIndex;

                            if (!sourceDocs.ContainsKey(srcPath))
                            {
                                sourceDocs[srcPath] = PdfReader.Open(srcPath, PdfDocumentOpenMode.Import);
                            }

                            var srcDoc = sourceDocs[srcPath];
                            if (srcIndex < 0 || srcIndex >= srcDoc.PageCount)
                                continue;

                            newPage = outputDoc.AddPage(srcDoc.Pages[srcIndex]);
                        }

                        // 회전 적용
                        int currentRotation = newPage.Rotate;
                        int finalRotation = (currentRotation + pageVM.Rotation) % 360;
                        if (finalRotation < 0)
                            finalRotation += 360;
                        newPage.Rotate = finalRotation;

                        // 주석 저장 (약식 구현)
                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SignaturePlaceholder)
                                continue;
                            // ... (기존 SavePdf의 주석 추가 로직 복사해서 여기에 넣으세요) ...
                            // 주의: doc.Pages[...] 대신 newPage.Annotations.Add(...) 사용

                            // [간단 예시 - Highlight]
                            if (ann.Type == AnnotationType.Highlight)
                            {
                                var rect = new PdfRectangle(new XRect(ann.X, ann.Y, ann.Width, ann.Height)); // 좌표 보정 필요
                                var newPdfAnn = new GenericPdfAnnotation(outputDoc);
                                newPdfAnn.Rectangle = rect;
                                newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
                                newPage.Annotations.Add(newPdfAnn);
                            }
                        }
                    }
                    outputDoc.Save(tempOutputPath);
                }
                finally
                {
                    foreach (var doc in sourceDocs.Values)
                        doc.Dispose();
                }
            }

            // 임시 파일을 최종 경로로 이동
            if (File.Exists(outputPath))
                File.Delete(outputPath);
            File.Move(tempOutputPath, outputPath);
        }

        // =========================================================
        // 5. 주석 추출
        // =========================================================
        private List<MinsPDFViewer.PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            var list = new List<MinsPDFViewer.PdfAnnotation>();
            if (page.Annotations == null)
                return list;

            double cropX = (page.CropBox.Width > 0) ? page.CropBox.X1 : 0;
            double cropY = (page.CropBox.Width > 0) ? page.CropBox.Y1 : 0;
            double cropH = (page.CropBox.Width > 0) ? page.CropBox.Height : page.Height.Point;

            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var pdfAnnot = page.Annotations[k] as PdfSharp.Pdf.Annotations.PdfAnnotation;
                if (pdfAnnot == null)
                    continue;
                var subtype = pdfAnnot.Elements.GetString("/Subtype");
                var rect = pdfAnnot.Rectangle.ToXRect();

                double finalX = rect.X - cropX;
                double finalY = (cropY + cropH) - (rect.Y + rect.Height);
                double finalW = rect.Width;
                double finalH = rect.Height;

                MinsPDFViewer.PdfAnnotation? newAnnot = null;
                string ft = pdfAnnot.Elements.GetString("/FT");
                string fieldName = pdfAnnot.Elements.GetString("/T");

                if (subtype == "/Widget" && (ft == "/Sig" || (!string.IsNullOrEmpty(fieldName) && string.IsNullOrEmpty(ft))))
                {
                    newAnnot = new MinsPDFViewer.PdfAnnotation { Type = AnnotationType.SignatureField, FieldName = fieldName };
                }
                else if (subtype == "/FreeText")
                {
                    double fontSize = 12;
                    var brush = Brushes.Black;
                    var da = pdfAnnot.Elements.GetString("/DA");
                    if (!string.IsNullOrEmpty(da))
                    {
                        try
                        {
                            var sizeMatch = System.Text.RegularExpressions.Regex.Match(da, @"(\d+(\.\d+)?) Tf");
                            if (sizeMatch.Success)
                                double.TryParse(sizeMatch.Groups[1].Value, out fontSize);
                            var colorMatch = System.Text.RegularExpressions.Regex.Match(da, @"(\d+(\.\d+)?) (\d+(\.\d+)?) (\d+(\.\d+)?) rg");
                            if (colorMatch.Success)
                            {
                                byte r = (byte)(double.Parse(colorMatch.Groups[1].Value) * 255);
                                byte g = (byte)(double.Parse(colorMatch.Groups[3].Value) * 255);
                                byte b = (byte)(double.Parse(colorMatch.Groups[5].Value) * 255);
                                brush = new SolidColorBrush(Color.FromRgb(r, g, b));
                            }
                        }
                        catch { }
                    }
                    var finalBrush = brush.Clone();
                    finalBrush.Freeze();
                    var transBrush = Brushes.Transparent.Clone();
                    transBrush.Freeze();
                    newAnnot = new MinsPDFViewer.PdfAnnotation { Type = AnnotationType.FreeText, TextContent = pdfAnnot.Contents, FontSize = fontSize, Foreground = finalBrush, Background = transBrush };
                }
                else if (subtype == "/Highlight")
                {
                    newAnnot = new MinsPDFViewer.PdfAnnotation { Type = AnnotationType.Highlight };
                    var colorArray = pdfAnnot.Elements.GetArray("/C");
                    if (colorArray != null && colorArray.Elements.Count >= 3)
                    {
                        byte r = (byte)(colorArray.Elements.GetReal(0) * 255);
                        byte g = (byte)(colorArray.Elements.GetReal(1) * 255);
                        byte b = (byte)(colorArray.Elements.GetReal(2) * 255);
                        var brush = new SolidColorBrush(Color.FromArgb(80, r, g, b));
                        brush.Freeze();
                        newAnnot.Background = brush;
                    }
                    else
                    {
                        var brush = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0));
                        brush.Freeze();
                        newAnnot.Background = brush;
                    }
                }
                else if (subtype == "/Underline")
                {
                    newAnnot = new MinsPDFViewer.PdfAnnotation { Type = AnnotationType.Underline, Background = Brushes.Black };
                }

                if (newAnnot != null)
                {
                    newAnnot.X = finalX;
                    newAnnot.Y = finalY;
                    newAnnot.Width = finalW;
                    newAnnot.Height = finalH;
                    list.Add(newAnnot);
                }
            }
            return list;
        }

        public void LoadBookmarks(PdfDocumentModel model)
        {
            try
            {
                if (!File.Exists(model.FilePath))
                    return;
                using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Import))
                {
                    model.Bookmarks.Clear();
                    foreach (var outline in doc.Outlines)
                    {
                        var vm = ConvertOutlineToViewModel(outline, doc, null);
                        if (vm != null)
                            model.Bookmarks.Add(vm);
                    }
                }
            }
            catch { }
        }

        private PdfBookmarkViewModel? ConvertOutlineToViewModel(PdfOutline outline, PdfDocument doc, PdfBookmarkViewModel? parent)
        {
            int pageIndex = -1;
            if (outline.DestinationPage != null)
            {
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    if (doc.Pages[i].Equals(outline.DestinationPage))
                    {
                        pageIndex = i;
                        break;
                    }
                }
            }
            if (pageIndex == -1)
                pageIndex = 0;
            var vm = new PdfBookmarkViewModel { Title = outline.Title, PageIndex = pageIndex, Parent = parent };
            foreach (var child in outline.Outlines)
            {
                var childVm = ConvertOutlineToViewModel(child, doc, vm);
                if (childVm != null)
                    vm.Children.Add(childVm);
            }
            return vm;
        }

        private void SaveBookmarkToDocument(PdfBookmarkViewModel vm, PdfOutlineCollection parentCollection, PdfDocument doc)
        {
            if (vm.PageIndex >= 0 && vm.PageIndex < doc.PageCount)
            {
                var page = doc.Pages[vm.PageIndex];
                var outline = parentCollection.Add(vm.Title, page, true);
                foreach (var child in vm.Children)
                    SaveBookmarkToDocument(child, outline.Outlines, doc);
            }
        }

        public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
        {
            public GenericPdfAnnotation(PdfDocument document) : base(document) { }
        }
    }
}