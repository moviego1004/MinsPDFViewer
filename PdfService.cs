using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf.Annotations; 
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Advanced;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public PdfDocumentModel? LoadPdf(string path)
        {
            if (!File.Exists(path)) return null;

            try
            {
                byte[] fileBytes = File.ReadAllBytes(path);

                var docReader = DocLib.Instance.GetDocReader(fileBytes, new PageDimensions(1.0));
                
                var model = new PdfDocumentModel
                {
                    FilePath = path,
                    FileName = System.IO.Path.GetFileName(path),
                    DocReader = docReader
                };

                using (var ms = new MemoryStream(fileBytes))
                using (var tempDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Import))
                {
                    for (int i = 0; i < docReader.GetPageCount(); i++)
                    {
                        using (var pageReader = docReader.GetPageReader(i))
                        {
                            var width = pageReader.GetPageWidth();
                            var height = pageReader.GetPageHeight();

                            var pageVm = new PdfPageViewModel
                            {
                                PageIndex = i,
                                Width = width,
                                Height = height,
                            };

                            if (tempDoc != null && i < tempDoc.PageCount)
                            {
                                var pdfPage = tempDoc.Pages[i];
                                
                                // [수정] CropBox 정보 상세 저장
                                pageVm.PdfPageWidthPoint = pdfPage.Width.Point;
                                pageVm.PdfPageHeightPoint = pdfPage.Height.Point;
                                pageVm.CropX = pdfPage.CropBox.X1;
                                pageVm.CropY = pdfPage.CropBox.Y1;
                                pageVm.CropWidthPoint = pdfPage.CropBox.Width;
                                pageVm.CropHeightPoint = pdfPage.CropBox.Height;

                                // Scale: (보이는 영역 Point) / (이미지 Pixel)
                                double scaleX = pageVm.CropWidthPoint / width;
                                double scaleY = pageVm.CropHeightPoint / height;

                                if (pdfPage.Annotations != null)
                                {
                                    foreach (var item in pdfPage.Annotations)
                                    {
                                        var pdfAnn = item as PdfSharp.Pdf.Annotations.PdfAnnotation;
                                        if (pdfAnn == null) continue;

                                        string subtype = pdfAnn.Elements.GetString("/Subtype");
                                        XRect rect = pdfAnn.Rectangle.ToXRect(); 

                                        // [수정] 주석 좌표 변환 로직 (나누기 적용)
                                        // 1. Point 좌표에서 여백(CropX, Y)을 뺌
                                        // 2. Scale로 나눔 (Pixel = Point / Scale)
                                        double annW = rect.Width / scaleX;
                                        double annH = rect.Height / scaleY;
                                        double annX = (rect.X - pageVm.CropX) / scaleX;
                                        
                                        // Y축: (CropTop - RectTop) / scale
                                        // CropTop = CropY + CropHeight
                                        // RectTop = rect.Y + rect.Height (Bottom-Up이므로)
                                        double cropTop = pageVm.CropY + pageVm.CropHeightPoint;
                                        double rectTop = rect.Y + rect.Height;
                                        double annY = (cropTop - rectTop) / scaleY;

                                        if (annX < 0) annX = 0;
                                        if (annY < 0) annY = 0;

                                        if (subtype == "/FreeText")
                                        {
                                            string content = pdfAnn.Contents;
                                            double fontSize = 14;
                                            Color fontColor = Colors.Black;
                                            string da = pdfAnn.Elements.GetString("/DA");
                                            
                                            if (!string.IsNullOrEmpty(da)) { try { /* DA 파싱 */ } catch { } }

                                            var newAnn = new PdfAnnotation
                                            {
                                                Type = AnnotationType.FreeText,
                                                X = annX, Y = annY, Width = annW, Height = annH,
                                                TextContent = content, FontSize = fontSize,
                                                Foreground = new SolidColorBrush(fontColor), Background = Brushes.Transparent
                                            };
                                            pageVm.Annotations.Add(newAnn);
                                        }
                                        else if (subtype == "/Highlight")
                                        {
                                            var newAnn = new PdfAnnotation { Type = AnnotationType.Highlight, X = annX, Y = annY, Width = annW, Height = annH, AnnotationColor = Colors.Yellow, Background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)) };
                                            pageVm.Annotations.Add(newAnn);
                                        }
                                        else if (subtype == "/Underline")
                                        {
                                            var newAnn = new PdfAnnotation { Type = AnnotationType.Underline, X = annX, Y = annY, Width = annW, AnnotationColor = Colors.Black, Background = Brushes.Black, Height = 2 };
                                            newAnn.Y = annY + annH - 2; 
                                            pageVm.Annotations.Add(newAnn);
                                        }
                                    }
                                }
                            }
                            model.Pages.Add(pageVm);
                        }
                    }
                }
                return model;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                return null;
            }
        }

        public async Task RenderPagesAsync(PdfDocumentModel document)
        {
            if (document == null || document.DocReader == null) return;
            await Task.Run(() =>
            {
                for (int i = 0; i < document.Pages.Count; i++)
                {
                    var page = document.Pages[i];
                    try
                    {
                        using (var pageReader = document.DocReader.GetPageReader(i))
                        {
                            var width = pageReader.GetPageWidth();
                            var height = pageReader.GetPageHeight();
                            var rawBytes = pageReader.GetImage(); 
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                var bitmap = BitmapSource.Create(width, height, 96, 96, PixelFormats.Bgra32, null, rawBytes, width * 4); 
                                bitmap.Freeze(); 
                                page.ImageSource = bitmap;
                            });
                        }
                    }
                    catch { }
                }
            });
        }

        public void SavePdf(PdfDocumentModel document, string savePath)
        {
            if (document == null) return;
            try
            {
                var originalBytes = File.ReadAllBytes(document.FilePath);
                using (var ms = new MemoryStream(originalBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    if (doc.Version < 14) doc.Version = 14;
                    var catalog = doc.Internals.Catalog;
                    var acroForm = catalog.Elements.GetDictionary("/AcroForm");
                    if (acroForm == null) { acroForm = new PdfDictionary(doc); catalog.Elements.Add("/AcroForm", acroForm); }
                    if (acroForm != null) { if (acroForm.Elements.ContainsKey("/NeedAppearances") == false) acroForm.Elements.Add("/NeedAppearances", new PdfBoolean(true)); else acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true); }

                    for (int i = 0; i < doc.PageCount && i < document.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i];
                        var pageVM = document.Pages[i];

                        // [수정] 저장 시 좌표 변환 (Scale 곱하기, 여백 더하기)
                        // Scale = Point / Pixel
                        double scaleX = pageVM.CropWidthPoint / pageVM.Width;
                        double scaleY = pageVM.CropHeightPoint / pageVM.Height;

                        if (pdfPage.Annotations != null)
                        {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            foreach(var item in pdfPage.Annotations) { var existingAnn = item as PdfSharp.Pdf.Annotations.PdfAnnotation; if (existingAnn == null) continue; string st = existingAnn.Elements.GetString("/Subtype"); if (st == "/FreeText" || st == "/Highlight" || st == "/Underline") toRemove.Add(existingAnn); }
                            foreach (var a in toRemove) pdfPage.Annotations.Remove(a);
                        }

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;

                            // 1. Pixel -> Point 변환 (곱하기)
                            // 2. 여백(CropX) 더하기 (절대 좌표로 복구)
                            double aw = ann.Width * scaleX;
                            double ah = ann.Height * scaleY;
                            double ax = (ann.X * scaleX) + pageVM.CropX;
                            double ay = (ann.Y * scaleY); // 일단 화면상 상대 Y (Top 기준)

                            // 3. Y축 반전 (Top-Down -> Bottom-Up)
                            // PDF Y = CropTop - 화면Y(Point) - 높이(Point)
                            double cropTop = pageVM.CropY + pageVM.CropHeightPoint;
                            double pdfY_BottomUp = cropTop - ay - ah;

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                pdfAnnot.Flags = PdfAnnotationFlags.Print;

                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");

                                var xrect = new XRect(0, 0, aw, ah);
                                var form = new XForm(doc, xrect);
                                using (var formGfx = XGraphics.FromForm(form)) { var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode); var font = new XFont(ann.FontFamily, ann.FontSize, XFontStyleEx.Regular, fontOptions); var brush = new XSolidBrush(XColor.FromArgb(color.A, color.R, color.G, color.B)); formGfx.DrawString(ann.TextContent, font, brush, xrect, XStringFormats.TopLeft); }
                                var pdfForm = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)?.GetValue(form) as PdfFormXObject;
                                if (pdfForm != null) { var apDict = new PdfDictionary(doc); apDict.Elements["/N"] = pdfForm.Reference; pdfAnnot.Elements["/AP"] = apDict; }
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else 
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                if (ann.Type == AnnotationType.Underline) pdfY_BottomUp = cropTop - ay - 2; // 밑줄은 조금 다르게
                                
                                var rect = new XRect(ax, pdfY_BottomUp, aw, ah);
                                pdfAnnot.Rectangle = new PdfRectangle(rect);
                                pdfAnnot.Elements.SetName("/Subtype", subtype);
                                pdfAnnot.Flags = PdfAnnotationFlags.Print;
                                var quadPoints = new PdfArray(doc, new PdfReal(rect.X), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X), new PdfReal(rect.Y), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);
                                double r = ann.AnnotationColor.R / 255.0; double g = ann.AnnotationColor.G / 255.0; double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }
                        // OCR 저장 부분도 동일하게 좌표 변환 필요 (생략 가능하나 일관성 위해 수정 권장)
                        // 여기서는 일단 기존 로직 유지 (주석 저장에 집중)
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }
    }
}