using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Reflection; // [추가] 리플렉션 사용
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
using PdfSharp.Pdf.Advanced; // [추가] PdfFormXObject 사용

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
                                double scaleX = pdfPage.Width.Point / width;
                                double scaleY = pdfPage.Height.Point / height;

                                if (pdfPage.Annotations != null)
                                {
                                    foreach (var item in pdfPage.Annotations)
                                    {
                                        var pdfAnn = item as PdfSharp.Pdf.Annotations.PdfAnnotation;
                                        if (pdfAnn == null) continue;

                                        string subtype = pdfAnn.Elements.GetString("/Subtype");
                                        XRect rect = pdfAnn.Rectangle.ToXRect(); 

                                        double annW = rect.Width / scaleX;
                                        double annH = rect.Height / scaleY;
                                        double annY = (pdfPage.Height.Point - (rect.Y + rect.Height)) / scaleY;
                                        double annX = rect.X / scaleX;

                                        if (annX < 0) annX = 0;
                                        if (annY < 0) annY = 0;

                                        if (subtype == "/FreeText")
                                        {
                                            string content = pdfAnn.Contents;
                                            
                                            double fontSize = 14;
                                            Color fontColor = Colors.Black;
                                            string da = pdfAnn.Elements.GetString("/DA");
                                            
                                            if (!string.IsNullOrEmpty(da))
                                            {
                                                try {
                                                    var parts = da.Split(' ');
                                                    for (int k = 0; k < parts.Length; k++)
                                                    {
                                                        if (parts[k] == "Tf" && k > 0)
                                                            double.TryParse(parts[k - 1], out fontSize);
                                                        
                                                        if (parts[k] == "rg" && k >= 3)
                                                        {
                                                            double r = double.Parse(parts[k - 3]);
                                                            double g = double.Parse(parts[k - 2]);
                                                            double b = double.Parse(parts[k - 1]);
                                                            fontColor = Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255));
                                                        }
                                                        else if (parts[k] == "g" && k >= 1)
                                                        {
                                                            double gray = double.Parse(parts[k - 1]);
                                                            fontColor = Color.FromRgb((byte)(gray * 255), (byte)(gray * 255), (byte)(gray * 255));
                                                        }
                                                    }
                                                } catch { }
                                            }

                                            var newAnn = new PdfAnnotation
                                            {
                                                Type = AnnotationType.FreeText,
                                                X = annX, Y = annY, Width = annW, Height = annH,
                                                TextContent = content,
                                                FontSize = fontSize,
                                                Foreground = new SolidColorBrush(fontColor),
                                                Background = Brushes.Transparent
                                            };
                                            pageVm.Annotations.Add(newAnn);
                                        }
                                        else if (subtype == "/Highlight")
                                        {
                                            var newAnn = new PdfAnnotation
                                            {
                                                Type = AnnotationType.Highlight,
                                                X = annX, Y = annY, Width = annW, Height = annH,
                                                AnnotationColor = Colors.Yellow,
                                                Background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0))
                                            };
                                            pageVm.Annotations.Add(newAnn);
                                        }
                                        else if (subtype == "/Underline")
                                        {
                                            var newAnn = new PdfAnnotation
                                            {
                                                Type = AnnotationType.Underline,
                                                X = annX, Y = annY, Width = annW,
                                                AnnotationColor = Colors.Black,
                                                Background = Brushes.Black,
                                                Height = 2
                                            };
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
                    if (acroForm == null)
                    {
                        acroForm = new PdfDictionary(doc);
                        catalog.Elements.Add("/AcroForm", acroForm);
                    }
                    
                    if (acroForm != null)
                    {
                        if (acroForm.Elements.ContainsKey("/NeedAppearances") == false)
                            acroForm.Elements.Add("/NeedAppearances", new PdfBoolean(true));
                        else
                            acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);
                    }

                    for (int i = 0; i < doc.PageCount && i < document.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i];
                        var pageVM = document.Pages[i];

                        double scaleX = pdfPage.Width.Point / pageVM.Width;
                        double scaleY = pdfPage.Height.Point / pageVM.Height;

                        if (pdfPage.Annotations != null)
                        {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            foreach(var item in pdfPage.Annotations)
                            {
                                var existingAnn = item as PdfSharp.Pdf.Annotations.PdfAnnotation;
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
                                pdfAnnot.Flags = PdfAnnotationFlags.Print;

                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");

                                // [핵심 수정] 외부 뷰어 호환성을 위한 AP Stream 생성 (XForm)
                                var xrect = new XRect(0, 0, aw, ah);
                                var form = new XForm(doc, xrect);
                                using (var formGfx = XGraphics.FromForm(form))
                                {
                                    var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                    var font = new XFont(ann.FontFamily, ann.FontSize, XFontStyleEx.Regular, fontOptions);
                                    var brush = new XSolidBrush(XColor.FromArgb(color.A, color.R, color.G, color.B));
                                    formGfx.DrawString(ann.TextContent, font, brush, xrect, XStringFormats.TopLeft);
                                }

                                // [핵심 수정] 리플렉션을 사용하여 숨겨진 PdfForm 속성에 접근 (CS1061 해결)
                                var pdfForm = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)?.GetValue(form) as PdfFormXObject;
                                if (pdfForm != null)
                                {
                                    var apDict = new PdfDictionary(doc);
                                    apDict.Elements["/N"] = pdfForm.Reference;
                                    pdfAnnot.Elements["/AP"] = apDict;
                                }

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
                                pdfAnnot.Flags = PdfAnnotationFlags.Print;

                                var quadPoints = new PdfArray(doc, new PdfReal(rect.X), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y + rect.Height), new PdfReal(rect.X), new PdfReal(rect.Y), new PdfReal(rect.X + rect.Width), new PdfReal(rect.Y));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);
                                
                                double r = ann.AnnotationColor.R / 255.0; double g = ann.AnnotationColor.G / 255.0; double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }

                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0)
                        {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);
                                foreach (var word in pageVM.OcrWords)
                                {
                                    double x = word.BoundingBox.X * scaleX;
                                    double y = word.BoundingBox.Y * scaleY;
                                    double w = word.BoundingBox.Width * scaleX;
                                    double h = word.BoundingBox.Height * scaleY;
                                    double fSize = h * 0.75; if (fSize < 1) fSize = 1;
                                    var font = new XFont("Malgun Gothic", fSize, XFontStyleEx.Regular, fontOptions);
                                    
                                    var nearlyInvisible = XColor.FromArgb(1, 255, 255, 255);
                                    gfx.DrawString(word.Text, font, new XSolidBrush(nearlyInvisible), new XRect(x, y, w, h), XStringFormats.Center);
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