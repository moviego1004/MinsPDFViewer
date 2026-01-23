using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfiumViewer;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout; 
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.Advanced;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;

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
            public const int FPDF_ANNOT_UNDERLINE = 10;

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
                    lock (PdfiumLock) { pdfDoc = PdfiumViewer.PdfDocument.Load(memoryStream); }
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

        private static void Log(string message)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                string logMsg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [PdfService] {message}{Environment.NewLine}";
                File.AppendAllText(logPath, logMsg);
            }
            catch { }
        }

        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage) return;
            Log($"RenderPageImage started for Page {pageVM.PageIndex}");
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
                            Log($"RenderPageImage rendering... Size: {renderW}x{renderH}");
                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, (int)(96 * RENDER_SCALE), (int)(96 * RENDER_SCALE), PdfRenderFlags.Annotations))
                            {
                                bmpSource = ToBitmapSource((DrawingBitmap)bitmap);
                                bmpSource.Freeze();
                            }
                            Log($"RenderPageImage rendering done.");
                        }
                        catch (Exception ex)
                        {
                            Log($"[CRITICAL] RenderPageImage Error: {ex}");
                        }
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
                lock (PdfiumLock)
                {
                    try
                    {
                        Log($"LoadAnnotationsLazy started for Page {pageVM.PageIndex}");
                        string path = pageVM.OriginalFilePath ?? model.FilePath;
                        if (!File.Exists(path)) return;

                        var extracted = new List<PdfAnnotation>();

                        IntPtr doc = NativeMethods.FPDF_LoadDocument(path, null);
                        if (doc == IntPtr.Zero)
                        {
                            Log($"LoadAnnotationsLazy: Failed to load document native {path}");
                            return;
                        }

                        try
                        {
                            IntPtr page = NativeMethods.FPDF_LoadPage(doc, pageVM.OriginalPageIndex);
                            if (page == IntPtr.Zero)
                            {
                                Log($"LoadAnnotationsLazy: Failed to load page native {pageVM.OriginalPageIndex}");
                                return;
                            }

                            try
                            {
                                double pageW = NativeMethods.FPDF_GetPageWidth(page);
                                double pageH = NativeMethods.FPDF_GetPageHeight(page);

                                int annotCount = NativeMethods.FPDFPage_GetAnnotCount(page);
                                Log($"LoadAnnotationsLazy: Page {pageVM.PageIndex} has {annotCount} annotations");

                                for (int i = 0; i < annotCount; i++)
                                {
                                    IntPtr annot = NativeMethods.FPDFPage_GetAnnot(page, i);
                                    if (annot == IntPtr.Zero) continue;

                                    try
                                    {
                                        int subtype = NativeMethods.FPDFAnnot_GetSubtype(annot);

                                        if (subtype == NativeMethods.FPDF_ANNOT_FREETEXT || 
                                            subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT || 
                                            subtype == NativeMethods.FPDF_ANNOT_UNDERLINE)
                                        {
                                            var rect = new NativeMethods.FS_RECTF();
                                            if (!NativeMethods.FPDFAnnot_GetRect(annot, ref rect)) continue;

                                            double uiX = rect.left * (pageVM.Width / pageW);
                                            double uiY = (pageH - rect.top) * (pageVM.Height / pageH);
                                            double uiW = Math.Abs(rect.right - rect.left) * (pageVM.Width / pageW);
                                            double uiH = Math.Abs(rect.top - rect.bottom) * (pageVM.Height / pageH);

                                            string content = GetAnnotationStringValue(annot, "Contents");
                                            if (content.StartsWith("Title: ")) // Adobe Note
                                                continue;

                                            double fSize = 12;
                                            Brush brush = Brushes.Black;
                                            Brush background = Brushes.Transparent;
                                            AnnotationType annType = AnnotationType.FreeText;

                                            if (subtype == NativeMethods.FPDF_ANNOT_FREETEXT)
                                            {
                                                annType = AnnotationType.FreeText;
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
                                                        var scb = new SolidColorBrush(Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255)));
                                                        scb.Freeze();
                                                        brush = scb;
                                                    }
                                                }
                                            }
                                            else if (subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT)
                                            {
                                                annType = AnnotationType.Highlight;
                                                var bgBrush = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0));
                                                bgBrush.Freeze();
                                                background = bgBrush;
                                                uint r = 0, g = 0, b = 0, a = 0;
                                                if (NativeMethods.FPDFAnnot_GetColor(annot, NativeMethods.FPDFANNOT_COLORTYPE_Color, ref r, ref g, ref b, ref a))
                                                {
                                                    var colorBrush = new SolidColorBrush(Color.FromArgb(80, (byte)r, (byte)g, (byte)b));
                                                    colorBrush.Freeze();
                                                    background = colorBrush;
                                                }
                                            }
                                            else if (subtype == NativeMethods.FPDF_ANNOT_UNDERLINE)
                                            {
                                                annType = AnnotationType.Underline;
                                            }

                                            extracted.Add(new PdfAnnotation
                                            {
                                                Type = annType,
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
                        Log($"LoadAnnotationsLazy finished for Page {pageVM.PageIndex}");
                    }
                    catch (Exception ex)
                    {
                        Log($"[CRITICAL] Annotation Load Error: {ex}");
                        System.Diagnostics.Debug.WriteLine($"Annotation Load Error: {ex.Message}");
                    }
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

        private class PageSaveData
        {
            public int OriginalPageIndex { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public double PdfPageWidthPoint { get; set; }
            public double PdfPageHeightPoint { get; set; }
            public List<AnnotationSaveData> Annotations { get; set; } = new List<AnnotationSaveData>();
            public List<OcrWordInfo> OcrWords { get; set; } = new List<OcrWordInfo>();
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

        private PdfFormXObject? GetPdfForm(XForm form)
        {
            var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (prop != null) return prop.GetValue(form) as PdfFormXObject;
            return null;
        }

        private void LogToFile(string msg)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [PdfService] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null) return;

            string originalFilePath = model.FilePath;
            LogToFile($"Starting SavePdf for: {originalFilePath}");

            // 1. [UI 스레드] 데이터 스냅샷
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
                        PdfPageHeightPoint = p.PdfPageHeightPoint,
                        OcrWords = p.OcrWords != null ? new List<OcrWordInfo>(p.OcrWords) : new List<OcrWordInfo>()
                    };

                    foreach (var ann in p.Annotations)
                    {
                        if (ann.Type == AnnotationType.SearchHighlight ||
                            ann.Type == AnnotationType.SignaturePlaceholder) continue;

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
                            ForegroundColor = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black,
                            BackgroundColor = (ann.Background as SolidColorBrush)?.Color ?? Colors.Transparent
                        };
                        pageData.Annotations.Add(annData);
                    }
                    pagesSnapshot.Add(pageData);
                }
            });
            LogToFile($"Snapshot created. Pages: {pagesSnapshot.Count}");

            await Task.Run(() =>
            {
                string tempWorkPath = Path.GetTempFileName();
                File.Copy(originalFilePath, tempWorkPath, true);
                LogToFile($"Copied to temp: {tempWorkPath}");

                bool standardSaveSuccess = false;

                // 1. Standard Save (Import)
                try
                {
                    using (var sourceDoc = PdfReader.Open(tempWorkPath, PdfDocumentOpenMode.Import))
                    using (var outputDoc = new PdfSharp.Pdf.PdfDocument()) 
                    {
                        LogToFile($"Document opened for Import. PageCount: {sourceDoc.PageCount}");

                        for (int i = 0; i < sourceDoc.PageCount; i++)
                        {
                            var pdfPage = outputDoc.AddPage(sourceDoc.Pages[i]);
                            // OCR 텍스트 레이어 추가 (Import 모드에서도 가능하면 그림)
                            if (i < pagesSnapshot.Count) DrawOcrText(outputDoc, pdfPage, pagesSnapshot[i].OcrWords, pagesSnapshot[i].Width, pagesSnapshot[i].Height);
                            
                            DrawAnnotationsOnPage(outputDoc, pdfPage, i, pagesSnapshot);
                        }
                        outputDoc.Save(outputPath);
                        standardSaveSuccess = true;
                        LogToFile("Standard save successful.");
                    }
                }
                catch (Exception ex)
                {
                    LogToFile($"Standard save failed: {ex.Message}. Trying fallback.");
                }

                // 2. Fallback Save (Rasterization + OCR Text)
                if (!standardSaveSuccess)
                {
                    try
                    {
                        // PdfiumViewer로 열어서 이미지로 변환
                        using (var pdfiumDoc = PdfiumViewer.PdfDocument.Load(tempWorkPath))
                        using (var outputDoc = new PdfSharp.Pdf.PdfDocument())
                        {
                            LogToFile("Starting fallback save (Rasterization).");
                            for (int i = 0; i < pdfiumDoc.PageCount; i++)
                            {
                                var size = pdfiumDoc.PageSizes[i];
                                // 고해상도 렌더링 (2배)
                                using (var bitmap = pdfiumDoc.Render(i, (int)size.Width * 2, (int)size.Height * 2, 192, 192, PdfRenderFlags.Annotations))
                                using (var ms = new MemoryStream())
                                {
                                    bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                    ms.Position = 0;

                                    var page = outputDoc.AddPage();
                                    page.Width = size.Width;
                                    page.Height = size.Height;

                                    using (var gfx = XGraphics.FromPdfPage(page))
                                    using (var xImage = XImage.FromStream(ms))
                                    {
                                        // 1. 이미지 그리기
                                        gfx.DrawImage(xImage, 0, 0, page.Width, page.Height);
                                        
                                        // 2. OCR 텍스트 그리기 (Searchable)
                                        if (i < pagesSnapshot.Count)
                                        {
                                            var pData = pagesSnapshot[i];
                                            DrawOcrText(outputDoc, page, pData.OcrWords, pData.Width, pData.Height);
                                        }
                                        
                                        // 3. 주석 그리기
                                        DrawAnnotationsOnPage(outputDoc, page, i, pagesSnapshot);
                                    }
                                }
                            }
                            outputDoc.Save(outputPath);
                            LogToFile("Fallback save successful.");
                        }
                    }
                    catch (Exception ex2)
                    {
                        LogToFile($"Fallback save failed: {ex2}");
                        throw;
                    }
                }

                try { if (File.Exists(tempWorkPath)) File.Delete(tempWorkPath); } catch { }
            });
        }

        private void DrawOcrText(PdfSharp.Pdf.PdfDocument doc, PdfPage page, List<OcrWordInfo> words, double viewWidth, double viewHeight)
        {
            if (words == null || words.Count == 0) return;

            try 
            {
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    var transparentBrush = new XSolidBrush(XColor.FromArgb(0, 0, 0, 0)); // 완전 투명
                    
                    double pdfPageW = page.Width.Point;
                    double pdfPageH = page.Height.Point;
                    double scaleX = pdfPageW / viewWidth;
                    double scaleY = pdfPageH / viewHeight;

                    foreach (var word in words)
                    {
                        double fSize = word.BoundingBox.Height * scaleY;
                        if (fSize <= 0) fSize = 10;
                        var font = new XFont("Malgun Gothic", fSize, XFontStyleEx.Regular);
                        
                        double x = word.BoundingBox.X * scaleX;
                        double y = word.BoundingBox.Y * scaleY;
                        
                        // 투명 텍스트 그리기 (검색 가능)
                        gfx.DrawString(word.Text, font, transparentBrush, x, y + (fSize * 0.8));
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile($"Error drawing OCR text: {ex.Message}");
            }
        }

        private void DrawAnnotationsOnPage(PdfSharp.Pdf.PdfDocument doc, PdfPage pdfPage, int pageIndex, List<PageSaveData> snapshots)
        {
            if (pageIndex >= snapshots.Count) return;
            var pageData = snapshots[pageIndex];

            foreach (var ann in pageData.Annotations)
            {
                try
                {
                    double pdfPageH = pdfPage.Height.Point;
                    double pdfPageW = pdfPage.Width.Point;
                    double scaleX = pdfPageW / pageData.Width;
                    double scaleY = pdfPageH / pageData.Height;

                    var rect = new PdfSharp.Pdf.PdfRectangle(new XRect(
                        ann.X * scaleX,
                        pdfPageH - ((ann.Y + ann.Height) * scaleY),
                        ann.Width * scaleX,
                        ann.Height * scaleY));

                    if (ann.Type == AnnotationType.FreeText)
                    {
                        var pdfAnnot = new GenericPdfAnnotation(doc);
                        pdfAnnot.Elements["/Subtype"] = new PdfName("/FreeText");
                        pdfAnnot.Elements["/Rect"] = rect;
                        pdfAnnot.Elements["/Contents"] = new PdfString(ann.TextContent, PdfStringEncoding.Unicode);

                        var form = new XForm(doc, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY));
                        using (var gfx = XGraphics.FromForm(form))
                        {
                            var font = new XFont("Malgun Gothic", ann.FontSize * scaleY, ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);
                            var brush = new XSolidBrush(XColor.FromArgb(ann.ForegroundColor.A, ann.ForegroundColor.R, ann.ForegroundColor.G, ann.ForegroundColor.B));
                            gfx.DrawString(ann.TextContent, font, brush, new XRect(0, 0, form.PixelWidth, form.PixelHeight), XStringFormats.TopLeft);
                        }

                        var apDict = new PdfDictionary(doc);
                        var pdfForm = GetPdfForm(form);
                        if (pdfForm != null) apDict.Elements["/N"] = pdfForm.Reference;
                        pdfAnnot.Elements["/AP"] = apDict;

                        string colorStr = $"{ann.ForegroundColor.R / 255.0:0.##} {ann.ForegroundColor.G / 255.0:0.##} {ann.ForegroundColor.B / 255.0:0.##} rg";
                        pdfAnnot.Elements["/DA"] = new PdfString($"/MalgunGothic {ann.FontSize * scaleY:0.##} Tf {colorStr}");

                        pdfPage.Annotations.Add(pdfAnnot);
                    }
                    else if (ann.Type == AnnotationType.Highlight || ann.Type == AnnotationType.Underline)
                    {
                        var pdfAnnot = new GenericPdfAnnotation(doc);
                        pdfAnnot.Elements["/Subtype"] = new PdfName(ann.Type == AnnotationType.Highlight ? "/Highlight" : "/Underline");
                        pdfAnnot.Elements["/Rect"] = rect;

                        var form = new XForm(doc, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY));
                        using (var gfx = XGraphics.FromForm(form))
                        {
                            if (ann.Type == AnnotationType.Highlight)
                            {
                                var brush = new XSolidBrush(XColor.FromArgb(ann.BackgroundColor.A, ann.BackgroundColor.R, ann.BackgroundColor.G, ann.BackgroundColor.B));
                                gfx.DrawRectangle(brush, 0, 0, form.PixelWidth, form.PixelHeight);
                            }
                            else
                            {
                                var pen = new XPen(XColors.Black, 1 * scaleY);
                                gfx.DrawLine(pen, 0, form.PixelHeight - 1, form.PixelWidth, form.PixelHeight - 1);
                            }
                        }
                        var apDict = new PdfDictionary(doc);
                        var pdfForm = GetPdfForm(form);
                        if (pdfForm != null) apDict.Elements["/N"] = pdfForm.Reference;
                        pdfAnnot.Elements["/AP"] = apDict;
                        pdfPage.Annotations.Add(pdfAnnot);
                    }
                }
                catch (Exception ex)
                {
                    LogToFile($"Annotation drawing error: {ex.Message}");
                }
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