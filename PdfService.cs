using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MinsPDFViewer
{
    public class PdfService
    {
        private readonly IDocLib _docLib;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        public PdfDocumentModel? LoadPdf(string path)
        {
            try
            {
                var newDoc = new PdfDocumentModel { FilePath = path, FileName = Path.GetFileName(path) };
                var fileBytes = File.ReadAllBytes(path);

                var extractedRawData = new Dictionary<int, List<RawAnnotationInfo>>();
                var pdfPageSizes = new Dictionary<int, XSize>();
                var pageRotations = new Dictionary<int, int>();
                var pageCropOffsets = new Dictionary<int, Point>();

                using (var msInput = new MemoryStream(fileBytes))
                using (var doc = PdfReader.Open(msInput, PdfDocumentOpenMode.Modify))
                {
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        pdfPageSizes[i] = new XSize(page.Width.Point, page.Height.Point);
                        pageRotations[i] = page.Rotate;
                        
                        var crop = page.CropBox.ToXRect();
                        pageCropOffsets[i] = new Point(crop.X, crop.Y);

                        extractedRawData[i] = new List<RawAnnotationInfo>();

                        if (page.Annotations != null)
                        {
                            var annotsToRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for (int k = 0; k < page.Annotations.Count; k++)
                            {
                                var ann = page.Annotations[k];
                                var subtype = ann.Elements.GetString("/Subtype");

                                if (subtype == "/FreeText") {
                                    var rect = ann.Rectangle.ToXRect();
                                    extractedRawData[i].Add(new RawAnnotationInfo { Type = AnnotationType.FreeText, Rect = rect.ToRect(), Content = ann.Contents, FontSize = 14, Color = Colors.Red });
                                    annotsToRemove.Add(ann); 
                                }
                                else if (subtype == "/Highlight" || subtype == "/Underline") {
                                    var rect = ann.Rectangle.ToXRect();
                                    Color cColor = Colors.Yellow; 
                                    var cArray = ann.Elements.GetArray("/C");
                                    if (cArray != null && cArray.Elements.Count >= 3) {
                                        double r = (cArray.Elements[0] as PdfReal)?.Value ?? 1.0;
                                        double g = (cArray.Elements[1] as PdfReal)?.Value ?? 1.0;
                                        double b = (cArray.Elements[2] as PdfReal)?.Value ?? 0.0;
                                        cColor = Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255));
                                    }
                                    AnnotationType type = (subtype == "/Highlight") ? AnnotationType.Highlight : AnnotationType.Underline;
                                    extractedRawData[i].Add(new RawAnnotationInfo { Type = type, Rect = rect.ToRect(), Color = cColor });
                                    annotsToRemove.Add(ann);
                                }
                            }
                            foreach (var item in annotsToRemove) page.Annotations.Remove(item);
                        }
                    }
                    
                    var cleanStream = new MemoryStream(); 
                    doc.Save(cleanStream);
                    newDoc.DocReader = _docLib.GetDocReader(cleanStream.ToArray(), new PageDimensions(2.0));
                }

                if (newDoc.DocReader != null)
                {
                    int pc = newDoc.DocReader.GetPageCount();
                    for (int i = 0; i < pc; i++)
                    {
                        using (var r = newDoc.DocReader.GetPageReader(i))
                        {
                            double viewW = r.GetPageWidth(); 
                            double viewH = r.GetPageHeight();
                            
                            var pvm = new PdfPageViewModel 
                            { 
                                PageIndex = i, 
                                Width = viewW, 
                                Height = viewH,
                                PdfPageWidthPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Width : viewW / 2.0,
                                PdfPageHeightPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Height : viewH / 2.0,
                                Rotation = pageRotations.ContainsKey(i) ? pageRotations[i] : 0,
                                CropX = pageCropOffsets.ContainsKey(i) ? pageCropOffsets[i].X : 0,
                                CropY = pageCropOffsets.ContainsKey(i) ? pageCropOffsets[i].Y : 0
                            };

                            double scaleX = viewW / pvm.PdfPageWidthPoint; 
                            double scaleY = viewH / pvm.PdfPageHeightPoint;

                            if (extractedRawData.ContainsKey(i)) {
                                foreach (var raw in extractedRawData[i]) {
                                    double pdfTopY = raw.Rect.Y + raw.Rect.Height; 
                                    double viewY = (pvm.PdfPageHeightPoint - pdfTopY) * scaleY;
                                    var ann = new PdfAnnotation {
                                        Type = raw.Type, X = raw.Rect.X * scaleX, Y = viewY, Width = raw.Rect.Width * scaleX, Height = raw.Rect.Height * scaleY,
                                        TextContent = raw.Content, FontSize = raw.FontSize, FontFamily = "Malgun Gothic", Foreground = new SolidColorBrush(raw.Color), AnnotationColor = raw.Color
                                    };
                                    if (raw.Type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, raw.Color.R, raw.Color.G, raw.Color.B));
                                    else if (raw.Type == AnnotationType.Underline) { ann.Background = new SolidColorBrush(raw.Color); ann.Y = viewY + (raw.Rect.Height * scaleY) - 2; ann.Height = 2; }
                                    pvm.Annotations.Add(ann);
                                }
                            }
                            newDoc.Pages.Add(pvm);
                        }
                    }
                }

                return newDoc;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                return null;
            }
        }

        public async Task RenderPagesAsync(PdfDocumentModel doc)
        {
            await Task.Run(() =>
            {
                Parallel.For(0, doc.Pages.Count, new ParallelOptions { MaxDegreeOfParallelism = 4 }, i =>
                {
                    if (doc.DocReader == null) return;
                    using (var r = doc.DocReader.GetPageReader(i))
                    {
                        var bytes = r.GetImage();
                        var w = r.GetPageWidth();
                        var h = r.GetPageHeight();
                        
                        Application.Current.Dispatcher.Invoke(() => 
                        { 
                            if (i < doc.Pages.Count) 
                                doc.Pages[i].ImageSource = RawBytesToBitmapImage(bytes, w, h); 
                        });
                    }
                });
            });
        }

        public void SavePdf(PdfDocumentModel docModel, string savePath)
        {
            try
            {
                var originalBytes = File.ReadAllBytes(docModel.FilePath);
                using (var ms = new MemoryStream(originalBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    if (doc.Version < 14) doc.Version = 14;
                    for (int i = 0; i < doc.PageCount && i < docModel.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i]; var pageVM = docModel.Pages[i];
                        double scaleX = pdfPage.Width.Point / pageVM.Width; 
                        double scaleY = pdfPage.Height.Point / pageVM.Height;
                        
                        if (pdfPage.Annotations != null) {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for(int k=0;k<pdfPage.Annotations.Count;k++) {
                                string st = pdfPage.Annotations[k].Elements.GetString("/Subtype");
                                if (st == "/FreeText" || st == "/Highlight" || st == "/Underline") toRemove.Add(pdfPage.Annotations[k]);
                            }
                            foreach(var a in toRemove) pdfPage.Annotations.Remove(a);
                        }

                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0) {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage)) {
                                foreach (var word in pageVM.OcrWords) {
                                    double x = word.BoundingBox.X * scaleX; 
                                    double y = word.BoundingBox.Y * scaleY; 
                                    double w = word.BoundingBox.Width * scaleX; 
                                    double h = word.BoundingBox.Height * scaleY;
                                    double fSize = h * 0.75; if(fSize < 1) fSize = 1;
                                    gfx.DrawString(word.Text, new XFont("Malgun Gothic", fSize), XBrushes.Transparent, new XRect(x, y, w, h), XStringFormats.Center);
                                }
                            }
                        }

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;
                            double ax = ann.X * scaleX; double ay = ann.Y * scaleY; double aw = ann.Width * scaleX; double ah = ann.Height * scaleY;
                            double pdfY_BottomUp = pdfPage.Height.Point - (ay + ah);

                            if (ann.Type == AnnotationType.FreeText) {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                if (ann.Type == AnnotationType.Underline) pdfY_BottomUp = pdfPage.Height.Point - (ay + 2);
                                var rect = new XRect(ax, pdfY_BottomUp, aw, ah);
                                pdfAnnot.Rectangle = new PdfRectangle(rect);
                                pdfAnnot.Elements.SetName("/Subtype", subtype);
                                double qLeft = rect.X; double qRight = rect.X + rect.Width; double qBottom = rect.Y; double qTop = rect.Y + rect.Height;
                                var quadPoints = new PdfArray(doc, new PdfReal(qLeft), new PdfReal(qTop), new PdfReal(qRight), new PdfReal(qTop), new PdfReal(qLeft), new PdfReal(qBottom), new PdfReal(qRight), new PdfReal(qBottom));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);
                                double r = ann.AnnotationColor.R / 255.0; double g = ann.AnnotationColor.G / 255.0; double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }

        private BitmapImage RawBytesToBitmapImage(byte[] b, int w, int h)
        {
            var bm = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
            bm.WritePixels(new Int32Rect(0, 0, w, h), b, w * 4, 0);
            if (bm.CanFreeze) bm.Freeze();
            
            using (var ms = new MemoryStream())
            {
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bm));
                enc.Save(ms);
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.StreamSource = ms;
                img.EndInit();
                if (img.CanFreeze) img.Freeze();
                return img;
            }
        }
    }
    
    public static class RectExtensions {
        public static Rect ToRect(this XRect xr) => new Rect(xr.X, xr.Y, xr.Width, xr.Height);
    }
}