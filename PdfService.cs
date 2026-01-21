using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfiumViewer;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.Advanced; // 추가: PdfFormXObject 사용을 위해
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;
using PdfiumDoc = PdfiumViewer.PdfDocument;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private const double RENDER_SCALE = 2.0;

        // Pdfium P/Invoke declarations
        private static class NativeMethods
        {
            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern IntPtr FPDF_LoadDocument([System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)] string path, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)] string? password);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDF_CloseDocument(IntPtr document);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern IntPtr FPDF_LoadPage(IntPtr document, int page_index);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDF_ClosePage(IntPtr page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern double FPDF_GetPageWidth(IntPtr page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern double FPDF_GetPageHeight(IntPtr page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFPage_GetAnnotCount(IntPtr page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPage_GetAnnot(IntPtr page, int index);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDFPage_CloseAnnot(IntPtr annot);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFAnnot_GetSubtype(IntPtr annot);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_GetRect(IntPtr annot, ref FS_RECTF rect);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern ulong FPDFAnnot_GetStringValue(IntPtr annot, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)] string key, IntPtr buffer, ulong buflen);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_GetColor(IntPtr annot, int color_type, ref uint R, ref uint G, ref uint B, ref uint A);

            [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
            public struct FS_RECTF
            {
                public float left;
                public float bottom;
                public float right;
                public float top;
            }

            public const int FPDF_ANNOT_UNKNOWN = 0;
            public const int FPDF_ANNOT_TEXT = 1;
            public const int FPDF_ANNOT_LINK = 2;
            public const int FPDF_ANNOT_FREETEXT = 3;
            public const int FPDF_ANNOT_HIGHLIGHT = 9;

            public const int FPDFANNOT_COLORTYPE_Color = 0;
        }

        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath)) return null;
            return await Task.Run(() =>
            {
                try
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    var memoryStream = new MemoryStream(fileBytes);
                    IPdfDocument? pdfDoc = null;
                    lock (PdfiumLock) { pdfDoc = PdfiumDoc.Load(memoryStream); }
                    var model = new PdfDocumentModel
                    {
                        FilePath = filePath,
                        FileName = Path.GetFileName(filePath),
                        PdfDocument = pdfDoc,
                        FileStream = memoryStream
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

        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.PdfDocument == null || model.IsDisposed) return;
            await Task.Run(() =>
            {
                if (model.IsDisposed) return;
                int pageCount = 0;
                lock (PdfiumLock) { if (model.PdfDocument != null) pageCount = model.PdfDocument.PageCount; }
                if (pageCount == 0) return;

                var tempPageList = new List<PdfPageViewModel>();
                for (int i = 0; i < pageCount; i++)
                {
                    double pageW = 0, pageH = 0;
                    lock (PdfiumLock)
                    {
                        if (model.PdfDocument != null)
                        {
                            var size = model.PdfDocument.PageSizes[i];
                            pageW = size.Width;
                            pageH = size.Height;
                        }
                    }
                    tempPageList.Add(new PdfPageViewModel
                    {
                        PageIndex = i,
                        OriginalFilePath = model.FilePath,
                        OriginalPageIndex = i,
                        Width = pageW,
                        Height = pageH,
                        PdfPageWidthPoint = pageW,
                        PdfPageHeightPoint = pageH,
                        CropWidthPoint = pageW,
                        CropHeightPoint = pageH
                    });
                }
                Application.Current.Dispatcher.Invoke(() => { foreach (var p in tempPageList) model.Pages.Add(p); });
            });
        }

        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage) return;
            if (pageVM.ImageSource == null)
            {
                BitmapSource? bmpSource = null;
                lock (PdfiumLock)
                {
                    if (model.PdfDocument != null && !model.IsDisposed)
                    {
                        try
                        {
                            int renderW = (int)(pageVM.Width * RENDER_SCALE);
                            int renderH = (int)(pageVM.Height * RENDER_SCALE);
                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, (int)(96 * RENDER_SCALE), (int)(96 * RENDER_SCALE), PdfRenderFlags.Annotations))
                            {
                                bmpSource = ToBitmapSource((DrawingBitmap)bitmap);
                                bmpSource.Freeze();
                            }
                        }
                        catch { }
                    }
                }
                if (bmpSource != null) Application.Current.Dispatcher.Invoke(() => pageVM.ImageSource = bmpSource);
            }
            LoadAnnotationsLazy(model, pageVM);
        }

        public static BitmapSource ToBitmapSource(DrawingBitmap bitmap)
        {
            var rect = new DrawingRectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            try { return BitmapSource.Create(bitmap.Width, bitmap.Height, bitmap.HorizontalResolution, bitmap.VerticalResolution, PixelFormats.Bgra32, null, bitmapData.Scan0, bitmapData.Stride * bitmap.Height, bitmapData.Stride); } 
            finally { bitmap.UnlockBits(bitmapData); }
        }

        private void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (pageVM.AnnotationsLoaded) return;
            pageVM.AnnotationsLoaded = true;
            Task.Run(() =>
            {
                try
                {
                    string path = pageVM.OriginalFilePath ?? model.FilePath;
                    if (!File.Exists(path)) return;

                    var extracted = new List<PdfAnnotation>();

                    IntPtr doc = NativeMethods.FPDF_LoadDocument(path, null);
                    if (doc == IntPtr.Zero) return;

                    try
                    {
                        IntPtr page = NativeMethods.FPDF_LoadPage(doc, pageVM.OriginalPageIndex);
                        if (page == IntPtr.Zero) return;

                        try
                        {
                            double pageW = NativeMethods.FPDF_GetPageWidth(page);
                            double pageH = NativeMethods.FPDF_GetPageHeight(page);

                            int annotCount = NativeMethods.FPDFPage_GetAnnotCount(page);
                            for (int i = 0; i < annotCount; i++)
                            {
                                IntPtr annot = NativeMethods.FPDFPage_GetAnnot(page, i);
                                if (annot == IntPtr.Zero) continue;

                                try
                                {
                                    int subtype = NativeMethods.FPDFAnnot_GetSubtype(annot);

                                    if (subtype == NativeMethods.FPDF_ANNOT_FREETEXT || subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT)
                                    {
                                        var rect = new NativeMethods.FS_RECTF();
                                        if (!NativeMethods.FPDFAnnot_GetRect(annot, ref rect)) continue;

                                        double uiX = rect.left * (pageVM.Width / pageW);
                                        double uiY = (pageH - rect.top) * (pageVM.Height / pageH);
                                        double uiW = Math.Abs(rect.right - rect.left) * (pageVM.Width / pageW);
                                        double uiH = Math.Abs(rect.top - rect.bottom) * (pageVM.Height / pageH);

                                        string content = GetAnnotationStringValue(annot, "Contents");
                                        if (content.StartsWith("Title: ")) // Adobe Note 무시
                                            continue;

                                        double fSize = 12;
                                        Brush brush = Brushes.Black;

                                        string da = GetAnnotationStringValue(annot, "DA");
                                        if (!string.IsNullOrEmpty(da))
                                        {
                                            var fsMatch = Regex.Match(da, @"([\d.]+)\s+Tf");
                                            if (fsMatch.Success) double.TryParse(fsMatch.Groups[1].Value, out fSize);

                                            var colMatch = Regex.Match(da, @"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg");
                                            if (colMatch.Success)
                                            {
                                                double.TryParse(colMatch.Groups[1].Value, out double r);
                                                double.TryParse(colMatch.Groups[2].Value, out double g);
                                                double.TryParse(colMatch.Groups[3].Value, out double b);
                                                brush = new SolidColorBrush(Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255)));
                                            }
                                        }

                                        Brush background = Brushes.Transparent;
                                        if (subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT)
                                        {
                                            background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0));
                                            uint r = 0, g = 0, b = 0, a = 0;
                                            if (NativeMethods.FPDFAnnot_GetColor(annot, NativeMethods.FPDFANNOT_COLORTYPE_Color, ref r, ref g, ref b, ref a))
                                            {
                                                background = new SolidColorBrush(Color.FromArgb(80, (byte)r, (byte)g, (byte)b));
                                            }
                                        }

                                        extracted.Add(new PdfAnnotation
                                        {
                                            Type = subtype == NativeMethods.FPDF_ANNOT_FREETEXT ? AnnotationType.FreeText : AnnotationType.Highlight,
                                            X = uiX,
                                            Y = uiY,
                                            Width = uiW,
                                            Height = uiH,
                                            TextContent = content,
                                            FontSize = fSize,
                                            Foreground = brush,
                                            Background = background
                                        });
                                    }
                                }
                                finally
                                {
                                    NativeMethods.FPDFPage_CloseAnnot(annot);
                                }
                            }
                        }
                        finally
                        {
                            NativeMethods.FPDF_ClosePage(page);
                        }
                    }
                    finally
                    {
                        NativeMethods.FPDF_CloseDocument(doc);
                    }

                    if (extracted.Count > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                                foreach (var item in extracted) pageVM.Annotations.Add(item);
                        });
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Annotation Load Error: {ex.Message}");
                }
            });
        }

        private string GetAnnotationStringValue(IntPtr annot, string key)
        {
            ulong len = NativeMethods.FPDFAnnot_GetStringValue(annot, key, IntPtr.Zero, 0);
            if (len == 0) return "";
            IntPtr buffer = System.Runtime.InteropServices.Marshal.AllocHGlobal((int)len);
            try
            {
                NativeMethods.FPDFAnnot_GetStringValue(annot, key, buffer, len);
                return System.Runtime.InteropServices.Marshal.PtrToStringUni(buffer) ?? "";
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FreeHGlobal(buffer);
            }
        }

        // 저장 시 스레드 간 데이터 전달을 위한 DTO
        private class PageSaveData
        {
            public int OriginalPageIndex { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public double PdfPageWidthPoint { get; set; }
            public double PdfPageHeightPoint { get; set; }
            public List<AnnotationSaveData> Annotations { get; set; } = new List<AnnotationSaveData>();
        }

        private class AnnotationSaveData
        {
            public AnnotationType Type { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public string TextContent { get; set; } = "";
            public double FontSize { get; set; }
            public string FontFamily { get; set; } = "Malgun Gothic";
            public bool IsBold { get; set; }
            public Color ForegroundColor { get; set; }
            public Color BackgroundColor { get; set; }
        }

        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null) return;

            string originalFilePath = model.FilePath;

            // 1. [UI 스레드] 저장에 필요한 데이터를 미리 추출 (스냅샷 생성)
            // UI 객체(Brush 등)는 다른 스레드에서 접근 불가능하므로 여기서 값(Color)으로 변환해둡니다.
            List<PageSaveData> pagesSnapshot = new List<PageSaveData>();

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var p in model.Pages)
                {
                    var pageData = new PageSaveData
                    {
                        OriginalPageIndex = p.OriginalPageIndex,
                        Width = p.Width,
                        Height = p.Height,
                        PdfPageWidthPoint = p.PdfPageWidthPoint,
                        PdfPageHeightPoint = p.PdfPageHeightPoint
                    };

                    foreach (var ann in p.Annotations)
                    {
                        if (ann.Type == AnnotationType.SearchHighlight ||
                            ann.Type == AnnotationType.SignaturePlaceholder ||
                            ann.Type == AnnotationType.SignatureField) continue;

                        var annData = new AnnotationSaveData
                        {
                            Type = ann.Type,
                            X = ann.X,
                            Y = ann.Y,
                            Width = ann.Width,
                            Height = ann.Height,
                            TextContent = ann.TextContent ?? "",
                            FontSize = ann.FontSize,
                            FontFamily = ann.FontFamily ?? "Malgun Gothic",
                            IsBold = ann.IsBold,
                            // Brush에서 Color 추출
                            ForegroundColor = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black,
                            BackgroundColor = (ann.Background as SolidColorBrush)?.Color ?? Colors.Transparent
                        };
                        pageData.Annotations.Add(annData);
                    }
                    pagesSnapshot.Add(pageData);
                }
            });
            
            await Task.Run(() =>
            {
                // 원본 파일을 임시 파일로 복사하여 작업
                string tempWorkPath = Path.GetTempFileName();
                File.Copy(originalFilePath, tempWorkPath, true);

                try
                {
                    // PdfSharp을 사용하여 PDF 편집 및 저장
                    using (var doc = PdfReader.Open(tempWorkPath, PdfDocumentOpenMode.Modify))
                    {
                        for (int i = 0; i < pagesSnapshot.Count; i++)
                        {
                            var pageData = pagesSnapshot[i];
                            int pdfPageIndex = pageData.OriginalPageIndex;

                            if (pdfPageIndex < 0 || pdfPageIndex >= doc.PageCount) continue;

                            var pdfPage = doc.Pages[pdfPageIndex];

                            // 1. 기존 주석 중 FreeText, Highlight, Underline 등 편집 가능한 타입 제거
                            if (pdfPage.Annotations != null)
                            {
                                var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                                
                                int count = pdfPage.Annotations.Count;
                                for (int idx = 0; idx < count; idx++)
                                {
                                    var item = pdfPage.Annotations[idx];
                                    
                                    string subtype = "";
                                    if (item.Elements.ContainsKey("/Subtype"))
                                        subtype = item.Elements.GetString("/Subtype");

                                    if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                    {
                                        toRemove.Add(item);
                                    }
                                }
                                
                                foreach (var item in toRemove)
                                {
                                    pdfPage.Annotations.Remove(item);
                                }
                            }

                            // 2. 스냅샷 데이터(DTO)를 사용하여 주석 추가
                            foreach (var ann in pageData.Annotations)
                            {
                                double pdfPageH = pdfPage.Height.Point;
                                double pdfPageW = pdfPage.Width.Point;
                                
                                double scaleX = pdfPageW / pageData.Width;
                                double scaleY = pdfPageH / pageData.Height;
                                
                                double rectX = ann.X * scaleX;
                                double rectY = pdfPageH - ((ann.Y + ann.Height) * scaleY);
                                double rectW = ann.Width * scaleX;
                                double rectH = ann.Height * scaleY;
                                
                                var rect = new PdfSharp.Pdf.PdfRectangle(new XRect(rectX, rectY, rectW, rectH));

                                if (ann.Type == AnnotationType.FreeText)
                                {
                                    var pdfAnnot = new GenericPdfAnnotation(doc);
                                    pdfAnnot.Elements["/Subtype"] = new PdfName("/FreeText");
                                    pdfAnnot.Elements["/Rect"] = rect;
                                    pdfAnnot.Elements["/Contents"] = new PdfString(ann.TextContent, PdfStringEncoding.Unicode);
                                    
                                    var form = new XForm(doc, new XRect(0, 0, rectW, rectH));
                                    using (var gfx = XGraphics.FromForm(form))
                                    {
                                        double fontSize = ann.FontSize * scaleY;
                                        if (fontSize < 1) fontSize = 10;
                                        
                                        var fontStyle = ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular;
                                        var fontFamily = string.IsNullOrEmpty(ann.FontFamily) ? "Malgun Gothic" : ann.FontFamily;
                                        var font = new XFont(fontFamily, fontSize, fontStyle);
                                        
                                        var c = ann.ForegroundColor;
                                        var brush = new XSolidBrush(XColor.FromArgb(c.A, c.R, c.G, c.B));
                                        
                                        var lines = ann.TextContent.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
                                        double lineHeight = font.GetHeight();
                                        double currentY = 0;
                                        
                                        foreach (var line in lines)
                                        {
                                            gfx.DrawString(line, font, brush, new XPoint(0, currentY + lineHeight * 0.8)); 
                                            currentY += lineHeight;
                                        }
                                    }
                                    
                                    var apDict = new PdfDictionary(doc);
                                    var pdfFormProp = typeof(XForm).GetProperty("PdfForm", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
                                    if (pdfFormProp != null)
                                    {
                                        var obj = pdfFormProp.GetValue(form);
                                        if (obj is PdfFormXObject pdfForm)
                                        {
                                            apDict.Elements["/N"] = pdfForm.Reference;
                                            pdfAnnot.Elements["/AP"] = apDict;
                                        }
                                    }
                                    
                                    pdfPage.Annotations.Add(pdfAnnot);
                                }
                                else if (ann.Type == AnnotationType.Highlight)
                                {
                                    var pdfAnnot = new GenericPdfAnnotation(doc);
                                    pdfAnnot.Elements["/Subtype"] = new PdfName("/Highlight");
                                    pdfAnnot.Elements["/Rect"] = rect;
                                    
                                    var c = ann.BackgroundColor; // 배경색 사용
                                    pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(c.R/255.0), new PdfReal(c.G/255.0), new PdfReal(c.B/255.0));
                                    pdfAnnot.Elements["/T"] = new PdfString(Environment.UserName);
                                    
                                    double qL = rectX;
                                    double qR = rectX + rectW;
                                    double qT = rectY + rectH;
                                    double qB = rectY;
                                    
                                    var quadPoints = new PdfArray(doc);
                                    quadPoints.Elements.Add(new PdfReal(qL)); quadPoints.Elements.Add(new PdfReal(qT));
                                    quadPoints.Elements.Add(new PdfReal(qR)); quadPoints.Elements.Add(new PdfReal(qT));
                                    quadPoints.Elements.Add(new PdfReal(qL)); quadPoints.Elements.Add(new PdfReal(qB));
                                    quadPoints.Elements.Add(new PdfReal(qR)); quadPoints.Elements.Add(new PdfReal(qB));
                                    
                                    pdfAnnot.Elements["/QuadPoints"] = quadPoints;
                                    pdfPage.Annotations.Add(pdfAnnot);
                                }
                                else if (ann.Type == AnnotationType.Underline)
                                {
                                    var pdfAnnot = new GenericPdfAnnotation(doc);
                                    pdfAnnot.Elements["/Subtype"] = new PdfName("/Underline");
                                    pdfAnnot.Elements["/Rect"] = rect;
                                    
                                    var c = Colors.Black;
                                    pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(c.R/255.0), new PdfReal(c.G/255.0), new PdfReal(c.B/255.0));
                                    
                                    double qL = rectX;
                                    double qR = rectX + rectW;
                                    double qT = rectY + rectH;
                                    double qB = rectY;
                                    
                                    var quadPoints = new PdfArray(doc);
                                    quadPoints.Elements.Add(new PdfReal(qL)); quadPoints.Elements.Add(new PdfReal(qT));
                                    quadPoints.Elements.Add(new PdfReal(qR)); quadPoints.Elements.Add(new PdfReal(qT));
                                    quadPoints.Elements.Add(new PdfReal(qL)); quadPoints.Elements.Add(new PdfReal(qB));
                                    quadPoints.Elements.Add(new PdfReal(qR)); quadPoints.Elements.Add(new PdfReal(qB));
                                    
                                    pdfAnnot.Elements["/QuadPoints"] = quadPoints;
                                    pdfPage.Annotations.Add(pdfAnnot);
                                }
                            }
                        }
                        
                        doc.Save(outputPath);
                    }
                }
                finally
                {
                    try { if (File.Exists(tempWorkPath)) File.Delete(tempWorkPath); } catch {}
                }
            });

            var newModel = await LoadPdfAsync(outputPath);
            if (newModel != null)
            {
                lock (PdfiumLock)
                {
                    try
                    {
                        model.PdfDocument?.Dispose();
                        model.FileStream?.Dispose();
                    }
                    catch { }
                }
                
                model.PdfDocument = newModel.PdfDocument;
                model.FileStream = newModel.FileStream;
                model.FilePath = outputPath;
                model.IsDisposed = false;
                
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var p in model.Pages)
                    {
                        p.AnnotationsLoaded = false;
                        p.Annotations.Clear();
                        p.ImageSource = null;
                    }
                });
            }
        }

        public void LoadBookmarks(PdfDocumentModel model)
        {
            lock (PdfiumLock)
            {
                if (model.PdfDocument?.Bookmarks == null) return;
                var list = new System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>();
                foreach (var b in model.PdfDocument.Bookmarks) list.Add(MapBm(b, null));
                Application.Current.Dispatcher.Invoke(() =>
                {
                    model.Bookmarks.Clear();
                    foreach (var i in list) model.Bookmarks.Add(i);
                });
            }
        }

        private PdfBookmarkViewModel MapBm(PdfBookmark b, PdfBookmarkViewModel? p)
        {
            var v = new PdfBookmarkViewModel { Title = b.Title, PageIndex = b.PageIndex, Parent = p };
            foreach (var c in b.Children) v.Children.Add(MapBm(c, v));
            return v;
        }
    }
}
