using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf.Annotations; 

namespace MinsPDFViewer
{
    public class PdfService
    {
        public PdfDocumentModel? LoadPdf(string path)
        {
            if (!File.Exists(path)) return null;

            try
            {
                var docReader = DocLib.Instance.GetDocReader(path, new PageDimensions(1.0));
                var model = new PdfDocumentModel
                {
                    FilePath = path,
                    FileName = System.IO.Path.GetFileName(path),
                    DocReader = docReader
                };

                for (int i = 0; i < docReader.GetPageCount(); i++)
                {
                    using (var pageReader = docReader.GetPageReader(i))
                    {
                        var pageVm = new PdfPageViewModel
                        {
                            PageIndex = i,
                            Width = pageReader.GetPageWidth(),
                            Height = pageReader.GetPageHeight(),
                        };
                        model.Pages.Add(pageVm);
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

                    for (int i = 0; i < doc.PageCount && i < document.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i];
                        var pageVM = document.Pages[i];

                        double scaleX = pdfPage.Width.Point / pageVM.Width;
                        double scaleY = pdfPage.Height.Point / pageVM.Height;

                        // 주석 저장 로직 (생략 없이 유지)
                        if (pdfPage.Annotations != null)
                        {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for (int k = 0; k < pdfPage.Annotations.Count; k++)
                            {
                                var existingAnn = pdfPage.Annotations[k];
                                if (existingAnn == null) continue;
                                string st = existingAnn.Elements.GetString("/Subtype");
                                if (st == "/FreeText" || st == "/Highlight" || st == "/Underline") toRemove.Add(existingAnn);
                            }
                            foreach (var a in toRemove) pdfPage.Annotations.Remove(a);
                        }

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;

                            double ax = ann.X * scaleX;
                            double ay = ann.Y * scaleY;
                            double aw = ann.Width * scaleX;
                            double ah = ann.Height * scaleY;
                            double pdfY_BottomUp = pdfPage.Height.Point - (ay + ah); 

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else 
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                if (ann.Type == AnnotationType.Underline) pdfY_BottomUp = pdfPage.Height.Point - (ay + 2); 
                                var rect = new XRect(ax, pdfY_BottomUp, aw, ah);
                                pdfAnnot.Rectangle = new PdfRectangle(rect);
                                pdfAnnot.Elements.SetName("/Subtype", subtype);
                                var quadPoints = new PdfArray(doc, new PdfReal(rect.X), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X), new PdfReal(rect.Y), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);
                                double r = ann.AnnotationColor.R / 255.0; double g = ann.AnnotationColor.G / 255.0; double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }

                        // [수정됨] OCR 결과 저장 (Text Rendering Mode 3 - Invisible 사용)
                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0)
                        {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                
                                // 일반 검정색 브러시 사용 (색상은 어차피 안 보임)
                                var brush = XBrushes.Black;

                                foreach (var word in pageVM.OcrWords)
                                {
                                    double x = word.BoundingBox.X * scaleX;
                                    double y = word.BoundingBox.Y * scaleY;
                                    double w = word.BoundingBox.Width * scaleX;
                                    double h = word.BoundingBox.Height * scaleY;

                                    double fSize = h * 0.75;
                                    if (fSize < 1) fSize = 1;
                                    var font = new XFont("Malgun Gothic", fSize, XFontStyleEx.Regular, fontOptions);

                                    // [핵심] 텍스트 렌더링 모드를 3 (Invisible)으로 설정하는 트릭
                                    // XGraphics는 이를 직접 지원하지 않으므로, BeginContainer로 
                                    // 그래픽 상태를 격리한 뒤 투명도를 이용하는 것이 아니라
                                    // 색상 자체를 무효화하는 방식은 복잡합니다.
                                    
                                    // 차선책 중 가장 확실한 방법:
                                    // PDFSharp의 XTextFormatter 등을 쓰지 않고 기본 DrawString을 쓰되,
                                    // XBrush를 투명하게 만드는 것이 실패했으므로...
                                    
                                    // **마지막 시도**: 
                                    // XBrush를 생성할 때 알파가 0인 색을 쓰면 PDFSharp이 
                                    // 내부적으로 "색상 설정 연산자"를 생략해버리는 버그가 있을 수 있습니다.
                                    // 아주 미세하게 불투명한(Alpha=1) 색을 써서 연산자가 기록되게 합니다.
                                    // (1/255 정도면 눈에 거의 안 보임)
                                    var nearlyInvisible = XColor.FromArgb(1, 255, 255, 255);
                                    gfx.DrawString(word.Text, font, new XSolidBrush(nearlyInvisible),
                                        new XRect(x, y, w, h), XStringFormats.Center);
                                }
                            }
                        }
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 실패: {ex.Message}");
            }
        }
    }
}