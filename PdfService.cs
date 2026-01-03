using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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
                    var bytes = File.ReadAllBytes(filePath);
                    IDocReader docReader;
                    lock (PdfiumLock)
                    {
                        // 메타데이터용 (1.0배)
                        docReader = _docLib.GetDocReader(bytes, new PageDimensions(1.0));
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
        // 2. 초기화 (렌더링 품질 대폭 향상)
        // =========================================================
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null)
                return;

            // [화질 개선 핵심] 1.5 -> 3.0 (300% 확대 렌더링)
            // Adobe/Foxit 수준의 선명함을 위해 3.0 권장 (FHD~4K 대응)
            double renderScale = 3.0;

            await Task.Run(() =>
            {
                byte[]? originalBytes = null;
                try
                {
                    originalBytes = File.ReadAllBytes(model.FilePath);
                }
                catch { return; }

                // 2-1. Clean PDF 생성
                byte[] cleanPdfBytes = originalBytes;
                try
                {
                    using (var ms = new MemoryStream(originalBytes))
                    using (var sharpDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                    {
                        bool modified = false;
                        for (int i = 0; i < sharpDoc.PageCount; i++)
                        {
                            var page = sharpDoc.Pages[i];
                            if (page.Annotations != null)
                            {
                                for (int k = page.Annotations.Count - 1; k >= 0; k--)
                                {
                                    var annot = page.Annotations[k];
                                    if (annot == null)
                                        continue;

                                    var subtype = annot.Elements.GetString("/Subtype");
                                    if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                    {
                                        page.Annotations.Elements.RemoveAt(k);
                                        modified = true;
                                    }
                                }
                            }
                        }
                        if (modified)
                        {
                            using (var outMs = new MemoryStream())
                            {
                                sharpDoc.Save(outMs);
                                cleanPdfBytes = outMs.ToArray();
                            }
                        }
                    }
                }
                catch { }

                // 2-2. Clean Reader (고해상도 설정)
                IDocReader cleanReader;
                lock (PdfiumLock)
                {
                    cleanReader = _docLib.GetDocReader(cleanPdfBytes, new PageDimensions(renderScale));
                }
                model.CleanDocReader = cleanReader;

                // 2-3. ViewModel 생성
                using (var msOriginal = new MemoryStream(originalBytes))
                using (var sharpDocOriginal = PdfReader.Open(msOriginal, PdfDocumentOpenMode.Import))
                {
                    int pageCount = cleanReader.GetPageCount();
                    for (int i = 0; i < pageCount; i++)
                    {
                        double rawWidth = 0, rawHeight = 0;
                        lock (PdfiumLock)
                        {
                            using (var pr = cleanReader.GetPageReader(i))
                            {
                                rawWidth = pr.GetPageWidth();
                                rawHeight = pr.GetPageHeight();
                            }
                        }

                        // UI 표시용 크기는 다시 원래대로(1.0 기준) 돌려서 계산
                        // (이미지는 3배 크지만, 화면 공간은 1배로 잡아야 3배 밀도로 선명하게 보임)
                        double uiWidth = rawWidth / renderScale;
                        double uiHeight = rawHeight / renderScale;

                        var pageVM = new PdfPageViewModel
                        {
                            PageIndex = i,
                            Width = uiWidth,
                            Height = uiHeight,
                            ImageSource = null
                        };

                        if (i < sharpDocOriginal.PageCount)
                        {
                            var p = sharpDocOriginal.Pages[i];
                            pageVM.PdfPageWidthPoint = p.Width.Point;
                            pageVM.PdfPageHeightPoint = p.Height.Point;

                            if (p.CropBox.Width > 0 && p.CropBox.Height > 0)
                            {
                                pageVM.CropX = p.CropBox.X1;
                                pageVM.CropY = p.CropBox.Y1;
                                pageVM.CropWidthPoint = p.CropBox.Width;
                                pageVM.CropHeightPoint = p.CropBox.Height;
                            }
                            else
                            {
                                pageVM.CropX = 0;
                                pageVM.CropY = 0;
                                pageVM.CropWidthPoint = p.Width.Point;
                                pageVM.CropHeightPoint = p.Height.Point;
                            }

                            pageVM.HasSignature = CheckIfPageHasSignature(p);
                            pageVM.Rotation = p.Rotate;
                        }

                        Application.Current.Dispatcher.Invoke(() => model.Pages.Add(pageVM));
                    }
                }
            });
        }

        // =========================================================
        // 3. 렌더링 & 지연 로딩
        // =========================================================
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (pageVM.ImageSource == null && model.CleanDocReader != null)
            {
                lock (PdfiumLock)
                {
                    using (var pageReader = model.CleanDocReader.GetPageReader(pageVM.PageIndex))
                    {
                        var w = pageReader.GetPageWidth();
                        var h = pageReader.GetPageHeight();

                        // [화질 개선] 
                        // 네온사인 현상(Color Fringing) 방지를 위해 LCD 텍스트 플래그 제거
                        // 오직 RenderAnnotations만 사용하여 순수 안티앨리어싱 적용
                        var bytes = pageReader.GetImage(RenderFlags.RenderAnnotations);

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Bgra32, null, bytes, w * 4);
                            bmp.Freeze();
                            pageVM.ImageSource = bmp;
                        });
                    }
                }
            }

            // [주석 지연 로딩]
            if (pageVM.Annotations.Count == 0)
            {
                Task.Run(() =>
                {
                    try
                    {
                        using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Import))
                        {
                            if (pageVM.PageIndex < doc.PageCount)
                            {
                                var p = doc.Pages[pageVM.PageIndex];
                                var extractedAnns = ExtractAnnotationsFromPage(p);

                                if (extractedAnns.Count > 0)
                                {
                                    Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        if (pageVM.Annotations.Count > 0)
                                            return;
                                        foreach (var ann in extractedAnns)
                                        {
                                            double scaleX = pageVM.Width / pageVM.PdfPageWidthPoint;
                                            double scaleY = pageVM.Height / pageVM.PdfPageHeightPoint;

                                            ann.X *= scaleX;
                                            ann.Y *= scaleY;
                                            ann.Width *= scaleX;
                                            if (ann.Type != AnnotationType.Underline)
                                                ann.Height *= scaleY;

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
        }

        // =========================================================
        // 4. 저장 (좌표 변환 수정됨)
        // =========================================================
        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || string.IsNullOrEmpty(model.FilePath))
                return;

            string tempPath = Path.GetTempFileName();
            File.Copy(model.FilePath, tempPath, true);

            try
            {
                using (var doc = PdfReader.Open(tempPath, PdfDocumentOpenMode.Modify))
                {
                    foreach (var pageVM in model.Pages)
                    {
                        if (pageVM.PageIndex >= doc.PageCount)
                            continue;
                        var pdfPage = doc.Pages[pageVM.PageIndex];

                        if (pdfPage.Annotations != null)
                        {
                            for (int k = pdfPage.Annotations.Count - 1; k >= 0; k--)
                            {
                                var annot = pdfPage.Annotations[k];
                                if (annot == null)
                                    continue;
                                var subtype = annot.Elements.GetString("/Subtype");
                                if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                {
                                    pdfPage.Annotations.Elements.RemoveAt(k);
                                }
                            }
                        }

                        foreach (MinsPDFViewer.PdfAnnotation ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SignaturePlaceholder || ann.Type == AnnotationType.SignatureField)
                                continue;

                            double scaleX = pageVM.PdfPageWidthPoint / pageVM.Width;
                            double scaleY = pageVM.PdfPageHeightPoint / pageVM.Height;

                            double pdfW = ann.Width * scaleX;
                            double pdfH = ann.Height * scaleY;

                            double pdfX = (ann.X + pageVM.CropX) * scaleX;
                            double visibleTop = pageVM.CropY + pageVM.CropHeightPoint;
                            double pdfY = visibleTop - (ann.Y * scaleY) - pdfH;

                            var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));
                            var newPdfAnn = new GenericPdfAnnotation(doc);
                            newPdfAnn.Rectangle = rect;

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/FreeText");
                                newPdfAnn.Contents = ann.TextContent;

                                string fontName = "/Helv";
                                double fontSize = ann.FontSize > 0 ? ann.FontSize : 12;
                                double r = 0, g = 0, b = 0;
                                if (ann.Foreground is SolidColorBrush brush)
                                {
                                    r = brush.Color.R / 255.0;
                                    g = brush.Color.G / 255.0;
                                    b = brush.Color.B / 255.0;
                                }
                                string daString = $"{fontName} {fontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg";
                                newPdfAnn.Elements.SetString("/DA", daString);
                            }
                            else if (ann.Type == AnnotationType.Highlight)
                            {
                                newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Highlight");
                                var colorArr = new PdfArray(doc);
                                if (ann.Background is SolidColorBrush b)
                                {
                                    colorArr.Elements.Add(new PdfReal(b.Color.R / 255.0));
                                    colorArr.Elements.Add(new PdfReal(b.Color.G / 255.0));
                                    colorArr.Elements.Add(new PdfReal(b.Color.B / 255.0));
                                }
                                else
                                {
                                    colorArr.Elements.Add(new PdfReal(1));
                                    colorArr.Elements.Add(new PdfReal(1));
                                    colorArr.Elements.Add(new PdfReal(0));
                                }
                                newPdfAnn.Elements.SetObject(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.C, colorArr);
                            }
                            else if (ann.Type == AnnotationType.Underline)
                            {
                                newPdfAnn.Elements.SetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype, "/Underline");
                            }

                            pdfPage.Annotations.Add(newPdfAnn);
                        }
                    }

                    doc.Outlines.Clear();
                    foreach (var bm in model.Bookmarks)
                    {
                        SaveBookmarkToDocument(bm, doc.Outlines, doc);
                    }

                    doc.Save(outputPath);
                }
            }
            catch (Exception ex)
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                throw new Exception($"저장 오류: {ex.Message}");
            }
            finally { if (File.Exists(tempPath)) File.Delete(tempPath); }
        }

        // =========================================================
        // 5. 주석 추출
        // =========================================================
        private List<MinsPDFViewer.PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            var list = new List<MinsPDFViewer.PdfAnnotation>();
            if (page.Annotations == null)
                return list;

            double cropX = 0;
            double cropY = 0;
            double cropH = page.Height.Point;

            if (page.CropBox.Width > 0 && page.CropBox.Height > 0)
            {
                cropX = page.CropBox.X1;
                cropY = page.CropBox.Y1;
                cropH = page.CropBox.Height;
            }

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
                            var sizeMatch = Regex.Match(da, @"(\d+(\.\d+)?) Tf");
                            if (sizeMatch.Success)
                                double.TryParse(sizeMatch.Groups[1].Value, out fontSize);

                            var colorMatch = Regex.Match(da, @"(\d+(\.\d+)?) (\d+(\.\d+)?) (\d+(\.\d+)?) rg");
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

                    newAnnot = new MinsPDFViewer.PdfAnnotation
                    {
                        Type = AnnotationType.FreeText,
                        TextContent = pdfAnnot.Contents,
                        FontSize = fontSize,
                        Foreground = finalBrush,
                        Background = transBrush
                    };
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

        private bool CheckIfPageHasSignature(PdfPage page)
        {
            if (page.Annotations == null)
                return false;
            for (int i = 0; i < page.Annotations.Count; i++)
            {
                var annot = page.Annotations[i];
                if (annot != null && annot.Elements.GetString("/Subtype") == "/Widget")
                    return true;
            }
            return false;
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
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"책갈피 로드 실패: {ex.Message}"); }
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

            var vm = new PdfBookmarkViewModel
            {
                Title = outline.Title,
                PageIndex = pageIndex,
                Parent = parent
            };

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
                {
                    SaveBookmarkToDocument(child, outline.Outlines, doc);
                }
            }
        }

        public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
        {
            public GenericPdfAnnotation(PdfDocument document) : base(document) { }
        }
    }
}