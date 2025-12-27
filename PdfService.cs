using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace MinsPDFViewer
{
    public class PdfService
    {
        private readonly IDocLib _docLib;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        public PdfDocumentModel? LoadPdf(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                var bytes = File.ReadAllBytes(filePath);
                var docReader = _docLib.GetDocReader(bytes, new PageDimensions(1.0));

                var model = new PdfDocumentModel
                {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath),
                    DocLib = _docLib,
                    DocReader = docReader
                };

                return model;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                return null;
            }
        }

        public async Task RenderPagesAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null)
                return;

            double renderScale = 3.0;

            await Task.Run(() =>
            {
                byte[]? originalBytes = null;
                try
                {
                    originalBytes = File.ReadAllBytes(model.FilePath);
                }
                catch { return; }

                // [Step 1] Clean PDF 생성 (주석 제거 버전)
                byte[]? cleanPdfBytes = null;
                try
                {
                    using (var ms = new MemoryStream(originalBytes))
                    using (var sharpDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                    {
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
                                    }
                                }
                            }
                        }
                        using (var outMs = new MemoryStream())
                        {
                            sharpDoc.Save(outMs);
                            cleanPdfBytes = outMs.ToArray();
                        }
                    }
                }
                catch
                {
                    cleanPdfBytes = originalBytes;
                }

                // [Step 2] 이미지 렌더링
                using (var renderReader = _docLib.GetDocReader(cleanPdfBytes, new PageDimensions(renderScale)))
                {
                    // [Step 3] 주석 추출
                    using (var msOriginal = new MemoryStream(originalBytes))
                    using (var sharpDocOriginal = PdfReader.Open(msOriginal, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < renderReader.GetPageCount(); i++)
                        {
                            // [핵심 수정] 개별 페이지 렌더링 오류가 전체 앱을 죽이지 않도록 방어
                            try
                            {
                                using (var pageReader = renderReader.GetPageReader(i))
                                {
                                    var rawWidth = pageReader.GetPageWidth();
                                    var rawHeight = pageReader.GetPageHeight();

                                    var rawBytes = pageReader.GetImage(RenderFlags.RenderAnnotations);

                                    BitmapSource? source = null;
                                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        source = BitmapSource.Create(rawWidth, rawHeight, 96, 96,
                                            System.Windows.Media.PixelFormats.Bgra32, null,
                                            rawBytes, rawWidth * 4);
                                        source.Freeze();
                                    });

                                    double uiWidth = rawWidth / renderScale;
                                    double uiHeight = rawHeight / renderScale;

                                    var pageVM = new PdfPageViewModel
                                    {
                                        PageIndex = i,
                                        Width = uiWidth,
                                        Height = uiHeight,
                                        ImageSource = source
                                    };

                                    if (i < sharpDocOriginal.PageCount)
                                    {
                                        var p = sharpDocOriginal.Pages[i];
                                        pageVM.PdfPageWidthPoint = p.Width.Point;
                                        pageVM.PdfPageHeightPoint = p.Height.Point;
                                        pageVM.CropX = p.CropBox.X1;
                                        pageVM.CropY = p.CropBox.Y1;
                                        pageVM.CropWidthPoint = p.CropBox.Width;
                                        pageVM.CropHeightPoint = p.CropBox.Height;

                                        // 서명 여부 확인 (UI 표시용)
                                        bool hasSignature = CheckIfPageHasSignature(p);
                                        pageVM.HasSignature = hasSignature;

                                        // 주석 추출
                                        var extractedAnns = ExtractAnnotationsFromPage(p);

                                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                        {
                                            foreach (var ann in extractedAnns)
                                            {
                                                double finalX = ann.X * (uiWidth / p.Width.Point);
                                                double finalY = ann.Y * (uiHeight / p.Height.Point);
                                                double finalW = ann.Width * (uiWidth / p.Width.Point);
                                                double finalH = ann.Height * (uiHeight / p.Height.Point);

                                                ann.X = finalX;
                                                ann.Y = finalY;
                                                ann.Width = finalW;
                                                if (ann.Type != AnnotationType.Underline)
                                                    ann.Height = finalH;

                                                pageVM.Annotations.Add(ann);
                                            }
                                        });
                                    }

                                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        model.Pages.Add(pageVM);
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                // 페이지 하나가 실패해도 죽지 않고 로그를 남기고 넘어감
                                System.Diagnostics.Debug.WriteLine($"[Error] Page {i} render failed: {ex.Message}");
                            }
                        }
                    }
                }
            });
        }

        private bool CheckIfPageHasSignature(PdfPage page)
        {
            if (page.Annotations == null)
                return false;
            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k];
                if (annot == null)
                    continue;
                if (annot.Elements.GetString("/Subtype") == "/Widget" && annot.Elements.GetString("/FT") == "/Sig")
                    return true;
            }
            return false;
        }

        private List<PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            var list = new List<PdfAnnotation>();
            if (page.Annotations == null)
                return list;

            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k];
                if (annot == null)
                    continue;

                var subtype = annot.Elements.GetString("/Subtype") ?? "";
                var rect = annot.Rectangle.ToXRect();

                double finalX = rect.X;
                double finalY = page.Height.Point - (rect.Y + rect.Height);
                double finalW = rect.Width;
                double finalH = rect.Height;

                PdfAnnotation? newAnnot = null;

                if (subtype == "/Widget" && annot.Elements.GetString("/FT") == "/Sig")
                {
                    string fieldName = annot.Elements.GetString("/T");

                    newAnnot = new PdfAnnotation
                    {
                        Type = AnnotationType.SignatureField,
                        FieldName = fieldName,
                        SignatureData = null
                    };
                }
                else if (subtype == "/FreeText")
                {
                    string da = annot.Elements.GetString("/DA");
                    Brush textColor = ParseAnnotationColor(da);
                    textColor.Freeze();
                    (double fontSize, bool isBold) = ParseAnnotationFont(da);

                    newAnnot = new PdfAnnotation
                    {
                        Type = AnnotationType.FreeText,
                        TextContent = annot.Contents,
                        FontSize = fontSize > 0 ? fontSize : 12,
                        IsBold = isBold,
                        Foreground = textColor,
                        Background = Brushes.Transparent
                    };
                }
                else if (subtype == "/Highlight")
                {
                    var highlightBrush = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0));
                    highlightBrush.Freeze();

                    newAnnot = new PdfAnnotation
                    {
                        Type = AnnotationType.Highlight,
                        AnnotationColor = Colors.Yellow,
                        Background = highlightBrush
                    };
                }
                else if (subtype == "/Underline")
                {
                    newAnnot = new PdfAnnotation
                    {
                        Type = AnnotationType.Underline,
                        AnnotationColor = Colors.Black,
                        Background = Brushes.Black,
                        Height = 2
                    };
                    finalY = finalY + finalH - 2;
                }

                if (newAnnot != null)
                {
                    newAnnot.X = finalX;
                    newAnnot.Y = finalY;
                    newAnnot.Width = finalW;
                    if (newAnnot.Type != AnnotationType.Underline)
                        newAnnot.Height = finalH;
                    list.Add(newAnnot);
                }
            }
            return list;
        }

        private (double size, bool bold) ParseAnnotationFont(string? da)
        {
            double size = 12;
            bool bold = false;
            if (string.IsNullOrEmpty(da))
                return (size, bold);
            try
            {
                var parts = da.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i] == "Tf" && i >= 2)
                    {
                        if (double.TryParse(parts[i - 1], NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedSize))
                            size = parsedSize;
                        if (parts[i - 2].IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0)
                            bold = true;
                    }
                }
            }
            catch { }
            return (size, bold);
        }

        private Brush ParseAnnotationColor(string? da)
        {
            if (string.IsNullOrEmpty(da))
                return Brushes.Black;
            try
            {
                var parts = da.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < parts.Length; i++)
                {
                    if (parts[i] == "rg" && i >= 3)
                    {
                        double r = double.Parse(parts[i - 3], CultureInfo.InvariantCulture);
                        double g = double.Parse(parts[i - 2], CultureInfo.InvariantCulture);
                        double b = double.Parse(parts[i - 1], CultureInfo.InvariantCulture);
                        return new SolidColorBrush(Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255)));
                    }
                    else if (parts[i] == "g" && i >= 1)
                    {
                        double gray = double.Parse(parts[i - 1], CultureInfo.InvariantCulture);
                        byte val = (byte)(gray * 255);
                        return new SolidColorBrush(Color.FromRgb(val, val, val));
                    }
                }
            }
            catch { }
            return Brushes.Black;
        }

        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || string.IsNullOrEmpty(model.FilePath))
                return;

            string tempOutputPath = Path.GetTempFileName();

            try
            {
                using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Modify))
                {
                    foreach (var pageVM in model.Pages)
                    {
                        if (pageVM.PageIndex >= doc.PageCount)
                            continue;
                        var pdfPage = doc.Pages[pageVM.PageIndex];

                        if (pdfPage.Annotations != null)
                        {
                            for (int i = pdfPage.Annotations.Count - 1; i >= 0; i--)
                            {
                                var annot = pdfPage.Annotations[i];
                                if (annot == null)
                                    continue;

                                var subtype = annot.Elements.GetString("/Subtype");
                                if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                {
                                    pdfPage.Annotations.Elements.RemoveAt(i);
                                }
                            }
                        }

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SignatureField || ann.Type == AnnotationType.SignaturePlaceholder)
                                continue;

                            double effectivePdfWidth = (pageVM.CropWidthPoint > 0) ? pageVM.CropWidthPoint : pageVM.PdfPageWidthPoint;
                            double effectivePdfHeight = (pageVM.CropHeightPoint > 0) ? pageVM.CropHeightPoint : pageVM.PdfPageHeightPoint;
                            double pdfOriginX = pageVM.CropX;
                            double pdfOriginY = pageVM.CropY;
                            double scaleX = effectivePdfWidth / pageVM.Width;
                            double scaleY = effectivePdfHeight / pageVM.Height;

                            double pdfX = pdfOriginX + (ann.X * scaleX);
                            double pdfY = (pdfOriginY + effectivePdfHeight) - ((ann.Y + ann.Height) * scaleY);
                            double pdfW = ann.Width * scaleX;
                            double pdfH = ann.Height * scaleY;

                            var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/FreeText");
                                pdfAnnot.Elements["/Contents"] = new PdfString(ann.TextContent);

                                double r = ann.Foreground is SolidColorBrush b ? b.Color.R / 255.0 : 0;
                                double g = ann.Foreground is SolidColorBrush b2 ? b2.Color.G / 255.0 : 0;
                                double b_ = ann.Foreground is SolidColorBrush b3 ? b3.Color.B / 255.0 : 0;
                                string fontName = ann.IsBold ? "/Helv-Bold" : "/Helv";
                                string da = $"{r.ToString("0.###", CultureInfo.InvariantCulture)} {g.ToString("0.###", CultureInfo.InvariantCulture)} {b_.ToString("0.###", CultureInfo.InvariantCulture)} rg {fontName} {ann.FontSize.ToString(CultureInfo.InvariantCulture)} Tf";
                                pdfAnnot.Elements["/DA"] = new PdfString(da);
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else if (ann.Type == AnnotationType.Highlight)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/Highlight");
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(1), new PdfReal(1), new PdfReal(0));
                                pdfAnnot.Elements["/CA"] = new PdfReal(0.5);
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else if (ann.Type == AnnotationType.Underline)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/Underline");
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(0), new PdfReal(0), new PdfReal(0));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }
                    }
                    doc.Save(tempOutputPath);
                }

                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempOutputPath, outputPath);
            }
            catch
            {
                if (File.Exists(tempOutputPath))
                    File.Delete(tempOutputPath);
                throw;
            }
        }
    }
}