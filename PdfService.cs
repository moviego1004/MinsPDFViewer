using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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
    /// <summary>
    /// Render request for the LIFO queue. Most recently added requests are processed first,
    /// ensuring that the currently visible page is rendered before pages that scrolled past.
    /// </summary>
    internal class RenderRequest
    {
        public PdfDocumentModel Model { get; set; } = null!;
        public PdfPageViewModel PageVM { get; set; } = null!;
        public CancellationToken CancellationToken { get; set; }
        public long Timestamp { get; set; }
        public double Zoom { get; set; } = 1.0;
    }

    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private static readonly object _logLock = new object();
        private const int PAGE_CACHE_SIZE = 50;

        // LIFO render queue: most recent requests are processed first
        private static readonly ConcurrentStack<RenderRequest> _renderStack = new ConcurrentStack<RenderRequest>();
        private static readonly SemaphoreSlim _renderSignal = new SemaphoreSlim(0);
        private static readonly Thread _renderWorker;
        private static volatile bool _shutdownRequested = false;

        // Separate annotation queue - lower priority than page rendering
        private static readonly ConcurrentQueue<RenderRequest> _annotationQueue = new ConcurrentQueue<RenderRequest>();
        private static readonly SemaphoreSlim _annotationSignal = new SemaphoreSlim(0);
        private static readonly Thread _annotationWorker;

        static PdfService()
        {
            _renderWorker = new Thread(RenderWorkerLoop)
            {
                IsBackground = true,
                Name = "PdfRenderWorker",
                Priority = ThreadPriority.AboveNormal
            };
            _renderWorker.Start();

            _annotationWorker = new Thread(AnnotationWorkerLoop)
            {
                IsBackground = true,
                Name = "PdfAnnotationWorker",
                Priority = ThreadPriority.BelowNormal
            };
            _annotationWorker.Start();
        }

        /// <summary>
        /// Dedicated annotation worker processes annotation loading at lower priority,
        /// ensuring it doesn't block page rendering.
        /// </summary>
        private static void AnnotationWorkerLoop()
        {
            while (!_shutdownRequested)
            {
                _annotationSignal.Wait();
                if (_shutdownRequested) break;

                while (_annotationQueue.TryDequeue(out var req))
                {
                    if (req.CancellationToken.IsCancellationRequested || req.Model.IsDisposed)
                        continue;

                    // Wait if render worker is active (renders have priority)
                    while (!_renderStack.IsEmpty && !req.CancellationToken.IsCancellationRequested)
                        Thread.Sleep(50);

                    if (req.CancellationToken.IsCancellationRequested) continue;

                    LoadAnnotationsLazy(req.Model, req.PageVM, req.CancellationToken);
                }
            }
        }

        /// <summary>
        /// Two-pass render worker: First pass renders ALL visible pages at low-res (fast),
        /// second pass upgrades them to high-res. This ensures all visible pages show content
        /// immediately when jumping to a distant page.
        /// </summary>
        private static void RenderWorkerLoop()
        {
            while (!_shutdownRequested)
            {
                _renderSignal.Wait();
                if (_shutdownRequested) break;

                // Drain all pending requests from the stack
                var batch = new List<RenderRequest>();
                while (_renderStack.TryPop(out var req))
                {
                    batch.Add(req);
                }

                if (batch.Count == 0) continue;

                // Drain excess semaphore signals to avoid unnecessary wake-ups
                while (_renderSignal.CurrentCount > 0 && _renderSignal.Wait(0)) { }

                // Filter out cancelled requests
                var activeBatch = new List<RenderRequest>();
                foreach (var req in batch)
                {
                    if (!req.CancellationToken.IsCancellationRequested && !req.Model.IsDisposed)
                        activeBatch.Add(req);
                }

                if (activeBatch.Count == 0) continue;

                // [Optimization] Limit processing to the most recent requests (e.g., 20 pages).
                // When zooming or fast scrolling, hundreds of requests might accumulate.
                // We only care about what's likely visible now (the top of the stack).
                if (activeBatch.Count > 20)
                {
                    // activeBatch is ordered by LIFO (index 0 is newest).
                    // Keep first 20, discard rest.
                    activeBatch = activeBatch.Take(20).ToList();
                }

                // PASS 1: Fast low-res render for ALL visible pages first
                // This ensures all pages show content within milliseconds
                foreach (var req in activeBatch)
                {
                    if (req.CancellationToken.IsCancellationRequested || req.Model.IsDisposed)
                        continue;
                    ExecuteFastRender(req);
                }

                // Check if new requests arrived (user scrolled again) - if so, skip high-res
                if (!_renderStack.IsEmpty) continue;

                // PASS 2: High-res render for all pages
                foreach (var req in activeBatch)
                {
                    if (req.CancellationToken.IsCancellationRequested || req.Model.IsDisposed)
                        continue;

                    // If new requests arrived during high-res pass, abort remaining
                    if (!_renderStack.IsEmpty) break;

                    ExecuteHighResRender(req);
                }
            }
        }

        /// <summary>
        /// Phase 1: Fast low-res render only. Shows content immediately at reduced quality.
        /// </summary>
        private static void ExecuteFastRender(RenderRequest req)
        {
            var model = req.Model;
            var pageVM = req.PageVM;
            var ct = req.CancellationToken;

            if (ct.IsCancellationRequested || model.IsDisposed) return;

            BitmapSource? fastBmp = null;
            lock (PdfiumLock)
            {
                if (ct.IsCancellationRequested || model.IsDisposed || model.PdfDocument == null) return;
                try
                {
                    // [Dynamic Scale] Use 1:1 scale for fast render (capped at 1.5x for performance)
                    // This fixes the "blurry" issue by matching screen pixels initially.
                    double scale = Math.Min(req.Zoom, 1.5);
                    int fastW = Math.Max(1, (int)(pageVM.Width * scale));
                    int fastH = Math.Max(1, (int)(pageVM.Height * scale));
                    using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, fastW, fastH, (int)(96 * scale), (int)(96 * scale), PdfRenderFlags.Annotations))
                    {
                        fastBmp = ToBitmapSource((DrawingBitmap)bitmap);
                        fastBmp.Freeze();
                    }
                }
                catch (Exception ex)
                {
                    Log($"[ERROR] Fast render failed for page {pageVM.PageIndex}: {ex.Message}");
                }
            }

            if (ct.IsCancellationRequested) return;

            // Show low-res immediately (non-blocking dispatch)
            if (fastBmp != null)
            {
                var bmp = fastBmp;
                Application.Current.Dispatcher.BeginInvoke(() =>
                {
                    if (!ct.IsCancellationRequested && pageVM.ImageSource == null)
                        pageVM.ImageSource = bmp;
                });
            }
        }

        /// <summary>
        /// Phase 2: High-res render. Upgrades the page to full quality and enqueues annotation loading.
        /// </summary>
        private static void ExecuteHighResRender(RenderRequest req)
        {
            var model = req.Model;
            var pageVM = req.PageVM;
            var ct = req.CancellationToken;

            if (ct.IsCancellationRequested || model.IsDisposed) return;

            BitmapSource? highBmp = null;
            lock (PdfiumLock)
            {
                if (ct.IsCancellationRequested || model.IsDisposed || model.PdfDocument == null) return;
                try
                {
                    // [Dynamic Scale] Render at 1.5x zoom for crisp text (Supersampling)
                    double scale = req.Zoom * 1.5;
                    int hiW = (int)(pageVM.Width * scale);
                    int hiH = (int)(pageVM.Height * scale);
                    using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, hiW, hiH, (int)(96 * scale), (int)(96 * scale), PdfRenderFlags.Annotations))
                    {
                        highBmp = ToBitmapSource((DrawingBitmap)bitmap);
                        highBmp.Freeze();
                    }
                }
                catch (Exception ex)
                {
                    Log($"[ERROR] High-res render failed for page {pageVM.PageIndex}: {ex.Message}");
                }
            }

            if (ct.IsCancellationRequested) return;

            if (highBmp != null)
            {
                string cacheKey = GetCacheKey(model.FilePath, pageVM.OriginalPageIndex);
                AddToCache(cacheKey, highBmp);
                var bmp = highBmp;
                Application.Current.Dispatcher.BeginInvoke(() =>
                {
                    if (!ct.IsCancellationRequested)
                    {
                        pageVM.ImageSource = bmp;
                        pageVM.IsHighResRendered = true;
                    }
                });
            }

            // Enqueue annotation loading on separate low-priority worker
            if (!ct.IsCancellationRequested)
            {
                _annotationQueue.Enqueue(req);
                _annotationSignal.Release();
            }
        }

        // Page image cache (key: filePath_pageIndex, value: frozen BitmapSource)
        private static readonly LinkedList<string> _cacheOrder = new LinkedList<string>();
        private static readonly Dictionary<string, BitmapSource> _pageCache = new Dictionary<string, BitmapSource>();
        private static readonly object _cacheLock = new object();

        private static string GetCacheKey(string filePath, int pageIndex) => $"{filePath}_{pageIndex}";

        private static void AddToCache(string key, BitmapSource bitmap)
        {
            lock (_cacheLock)
            {
                if (_pageCache.ContainsKey(key))
                {
                    _cacheOrder.Remove(key);
                }
                _pageCache[key] = bitmap;
                _cacheOrder.AddFirst(key);

                while (_cacheOrder.Count > PAGE_CACHE_SIZE)
                {
                    var last = _cacheOrder.Last!.Value;
                    _cacheOrder.RemoveLast();
                    _pageCache.Remove(last);
                }
            }
        }

        private static BitmapSource? GetFromCache(string key)
        {
            lock (_cacheLock)
            {
                if (_pageCache.TryGetValue(key, out var bmp))
                {
                    _cacheOrder.Remove(key);
                    _cacheOrder.AddFirst(key);
                    return bmp;
                }
            }
            return null;
        }

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
                public float top;
                public float right;
                public float bottom;
            }

            public const int FPDF_ANNOT_UNKNOWN = 0;
            public const int FPDF_ANNOT_TEXT = 1;
            public const int FPDF_ANNOT_LINK = 2;
            public const int FPDF_ANNOT_FREETEXT = 3;
            public const int FPDF_ANNOT_HIGHLIGHT = 9;
            public const int FPDF_ANNOT_UNDERLINE = 10;
            public const int FPDF_ANNOT_WIDGET = 12;

            public const int FPDFANNOT_COLORTYPE_Color = 0;

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern IntPtr FPDFText_LoadPage(IntPtr page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDFText_ClosePage(IntPtr text_page);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern IntPtr FPDFText_FindStart(IntPtr text_page, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string findwhat, ulong flags, int start_index);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern bool FPDFText_FindNext(IntPtr handle);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDFText_FindClose(IntPtr handle);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFText_GetSchResultIndex(IntPtr handle);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFText_GetSchCount(IntPtr handle);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFText_CountRects(IntPtr text_page, int start_index, int count);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern bool FPDFText_GetRect(IntPtr text_page, int rect_index, ref double left, ref double top, ref double right, ref double bottom);

            public const int FPDF_MATCHCASE = 0x00000001;
            public const int FPDF_MATCHWHOLEWORD = 0x00000002;
        }

        public List<Rect> FindTextRects(string filePath, int pageIndex, string keyword)
        {
            var results = new List<Rect>();
            if (!File.Exists(filePath)) return results;

            lock (PdfiumLock)
            {
                IntPtr doc = NativeMethods.FPDF_LoadDocument(filePath, null);
                if (doc == IntPtr.Zero) return results;

                try
                {
                    IntPtr page = NativeMethods.FPDF_LoadPage(doc, pageIndex);
                    if (page == IntPtr.Zero) return results;

                    try
                    {
                        IntPtr textPage = NativeMethods.FPDFText_LoadPage(page);
                        if (textPage == IntPtr.Zero) return results;

                        try
                        {
                            IntPtr search = NativeMethods.FPDFText_FindStart(textPage, keyword, NativeMethods.FPDF_MATCHCASE, 0);
                            if (search != IntPtr.Zero)
                            {
                                try
                                {
                                    while (NativeMethods.FPDFText_FindNext(search))
                                    {
                                        int charIndex = NativeMethods.FPDFText_GetSchResultIndex(search);
                                        int charCount = NativeMethods.FPDFText_GetSchCount(search);
                                        
                                        int rectCount = NativeMethods.FPDFText_CountRects(textPage, charIndex, charCount);
                                        for (int i = 0; i < rectCount; i++)
                                        {
                                            double left = 0, top = 0, right = 0, bottom = 0;
                                            if (NativeMethods.FPDFText_GetRect(textPage, i, ref left, ref top, ref right, ref bottom))
                                            {
                                                // PDF coordinates (0,0 is bottom-left) to UI coordinates logic is handled by caller or here?
                                                // Let's return raw PDF coordinates here, caller (SearchService) will convert.
                                                // However, FPDFText_GetRect returns page coordinates.
                                                // Pdfium coordinates: origin bottom-left.
                                                // We will return generic Rect(left, top, width, height) but keep in mind Y is inverted relative to UI.
                                                
                                                results.Add(new Rect(left, top, right - left, top - bottom)); 
                                                // Note: top > bottom in PDF coords usually.
                                                // But let's verify: FPDFText_GetRect returns (left, top, right, bottom).
                                                // Wait, standard PDF: Y increases upwards.
                                                // FPDFText_GetRect usually follows PDF coord system.
                                                // Width = right - left
                                                // Height = top - bottom (since top is higher Y value)
                                            }
                                        }
                                    }
                                }
                                finally
                                {
                                    NativeMethods.FPDFText_FindClose(search);
                                }
                            }
                        }
                        finally
                        {
                            NativeMethods.FPDFText_ClosePage(textPage);
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
            }
            return results;
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

        /// <summary>
        /// Enqueues a two-phase render via LIFO stack.
        /// The most recently requested page is rendered first, so jumping to page 400
        /// renders that page immediately instead of waiting for intermediate pages.
        /// </summary>
        public void RenderPageAsync(PdfDocumentModel model, PdfPageViewModel pageVM, CancellationToken ct)
        {
            if (model.IsDisposed || pageVM.IsBlankPage) return;

            // Check cache first
            string cacheKey = GetCacheKey(model.FilePath, pageVM.OriginalPageIndex);
            var cached = GetFromCache(cacheKey);
            if (cached != null)
            {
                Application.Current.Dispatcher.BeginInvoke(() =>
                {
                    pageVM.ImageSource = cached;
                    pageVM.IsHighResRendered = true;
                });
                // Enqueue annotation loading on low-priority worker
                _annotationQueue.Enqueue(new RenderRequest
                {
                    Model = model,
                    PageVM = pageVM,
                    CancellationToken = ct,
                    Timestamp = Environment.TickCount64,
                    Zoom = model.Zoom
                });
                _annotationSignal.Release();
                return;
            }

            // Push to LIFO stack - most recent pages are rendered first
            _renderStack.Push(new RenderRequest
            {
                Model = model,
                PageVM = pageVM,
                CancellationToken = ct,
                Timestamp = Environment.TickCount64,
                Zoom = model.Zoom
            });
            _renderSignal.Release();
        }

        // Keep legacy method for compatibility with other callers (save, OCR, etc.)
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
                            double scale = model.Zoom * 1.5;
                            int renderW = (int)(pageVM.Width * scale);
                            int renderH = (int)(pageVM.Height * scale);
                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, (int)(96 * scale), (int)(96 * scale), PdfRenderFlags.Annotations))
                            {
                                bmpSource = ToBitmapSource((DrawingBitmap)bitmap);
                                bmpSource.Freeze();
                            }
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

        private static void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM, CancellationToken ct = default)
        {
            if (pageVM.AnnotationsLoaded) return;
            pageVM.AnnotationsLoaded = true;

            if (ct.IsCancellationRequested || model.IsDisposed)
            {
                pageVM.AnnotationsLoaded = false;
                return;
            }

            // Wait until no render requests are pending (renders have priority)
            while (!_renderStack.IsEmpty && !ct.IsCancellationRequested)
                Thread.Sleep(30);

            if (ct.IsCancellationRequested || model.IsDisposed)
            {
                pageVM.AnnotationsLoaded = false;
                return;
            }

            lock (PdfiumLock)
            {
                if (ct.IsCancellationRequested || model.IsDisposed)
                {
                    pageVM.AnnotationsLoaded = false;
                    return;
                }
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

                                    if (subtype == NativeMethods.FPDF_ANNOT_FREETEXT ||
                                        subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT ||
                                        subtype == NativeMethods.FPDF_ANNOT_UNDERLINE ||
                                        subtype == NativeMethods.FPDF_ANNOT_WIDGET)
                                    {
                                        var rect = new NativeMethods.FS_RECTF();
                                        if (!NativeMethods.FPDFAnnot_GetRect(annot, ref rect)) continue;

                                        double uiX = rect.left * (pageVM.Width / pageW);
                                        double uiY = (pageH - rect.top) * (pageVM.Height / pageH);
                                        double uiW = Math.Abs(rect.right - rect.left) * (pageVM.Width / pageW);
                                        double uiH = Math.Abs(rect.top - rect.bottom) * (pageVM.Height / pageH);

                                        string content = GetAnnotationStringValue(annot, "Contents");
                                        // Adobe Note check skipped for Widget
                                        if (subtype != NativeMethods.FPDF_ANNOT_WIDGET && content.StartsWith("Title: ")) 
                                            continue;

                                        double fSize = 12;
                                        Brush brush = Brushes.Black;
                                        Brush background = Brushes.Transparent;
                                        AnnotationType annType = AnnotationType.FreeText;
                                        string fieldName = "";

                                        if (subtype == NativeMethods.FPDF_ANNOT_FREETEXT)
                                        {
                                            // ... (Existing FreeText logic) ...
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
                                        else if (subtype == NativeMethods.FPDF_ANNOT_WIDGET)
                                        {
                                            annType = AnnotationType.SignatureField;
                                            // Try to get Field Name (T)
                                            // Note: GetStringValue might not work for 'T' if it's inherited, but worth a try for simple cases
                                            fieldName = GetAnnotationStringValue(annot, "T");
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
                                            Background = background,
                                            FieldName = fieldName
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
                        Application.Current.Dispatcher.BeginInvoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                                foreach (var item in extracted) pageVM.Annotations.Add(item);
                        });
                    }
                }
                catch (Exception ex)
                {
                    Log($"[CRITICAL] Annotation Load Error: {ex}");
                }
            }
        }

        private static string GetAnnotationStringValue(IntPtr annot, string key)
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

        private class BookmarkSaveData
        {
            public string Title { get; set; } = "";
            public int OriginalPageIndex { get; set; }
            public List<BookmarkSaveData> Children { get; set; } = new List<BookmarkSaveData>();
        }

        private BookmarkSaveData MapBookmarkSnapshot(PdfBookmarkViewModel vm)
        {
            var data = new BookmarkSaveData
            {
                Title = vm.Title,
                OriginalPageIndex = vm.PageIndex
            };
            foreach (var child in vm.Children)
            {
                data.Children.Add(MapBookmarkSnapshot(child));
            }
            return data;
        }

        private PdfFormXObject? GetPdfForm(XForm form)
        {
            var prop = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (prop != null) return prop.GetValue(form) as PdfFormXObject;
            return null;
        }

        private void LogToFile(string msg)
        {
            lock (_logLock)
            {
                try
                {
                    string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                    File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [PdfService] {msg}{Environment.NewLine}");
                }
                catch { }
            }
        }

        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null) return;

            string originalFilePath = model.FilePath;
            LogToFile($"Starting SavePdf for: {originalFilePath}");

            // 1. [UI 스레드] 데이터 스냅샷
            List<PageSaveData> pagesSnapshot = new List<PageSaveData>();
            List<BookmarkSaveData> bookmarksSnapshot = new List<BookmarkSaveData>();

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

                // Bookmark Snapshot
                foreach (var bm in model.Bookmarks)
                {
                    bookmarksSnapshot.Add(MapBookmarkSnapshot(bm));
                }
            });
            LogToFile($"Snapshot created. Pages: {pagesSnapshot.Count}, Bookmarks: {bookmarksSnapshot.Count}");

            await Task.Run(() =>
            {
                string tempWorkPath = Path.GetTempFileName();
                File.Copy(originalFilePath, tempWorkPath, true);
                LogToFile($"Copied to temp: {tempWorkPath}");

                bool standardSaveSuccess = false;

                // Page Mapping for Bookmarks: OriginalPageIndex -> New PdfPage
                Dictionary<int, PdfPage> pageMapping = new Dictionary<int, PdfPage>();

                // 1. Standard Save (Import)
                try
                {
                    using (var sourceDoc = PdfReader.Open(tempWorkPath, PdfDocumentOpenMode.Import))
                    using (var outputDoc = new PdfSharp.Pdf.PdfDocument()) 
                    {
                        LogToFile($"Document opened for Import. PageCount: {sourceDoc.PageCount}, SnapshotCount: {pagesSnapshot.Count}");

                        // [Fix] Iterate through UI pages snapshot
                        for (int i = 0; i < pagesSnapshot.Count; i++)
                        {
                            var pageData = pagesSnapshot[i];
                            int originalIdx = pageData.OriginalPageIndex;

                            if (originalIdx >= 0 && originalIdx < sourceDoc.PageCount)
                            {
                                var pdfPage = outputDoc.AddPage(sourceDoc.Pages[originalIdx]);
                                pageMapping[originalIdx] = pdfPage; // Register mapping

                                // Clear existing annotations to avoid duplication
                                pdfPage.Annotations.Clear();

                                // OCR Text Layer
                                DrawOcrText(outputDoc, pdfPage, pageData.OcrWords, pageData.Width, pageData.Height);
                                
                                // Annotations
                                DrawAnnotationsOnPage(outputDoc, pdfPage, i, pagesSnapshot);
                            }
                        }

                        // Restore Bookmarks
                        foreach (var bm in bookmarksSnapshot)
                        {
                            AddBookmarkToPdf(outputDoc.Outlines, bm, pageMapping);
                        }

                        // Set NeedAppearances
                        var catalog = outputDoc.Internals.Catalog;
                        var acroForm = catalog.Elements.GetDictionary("/AcroForm");
                        if (acroForm == null)
                        {
                            acroForm = new PdfDictionary(outputDoc);
                            catalog.Elements["/AcroForm"] = acroForm;
                        }
                        acroForm.Elements["/NeedAppearances"] = new PdfBoolean(true);

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
                        // Reset Mapping for fallback
                        pageMapping.Clear();

                        using (var pdfiumDoc = PdfiumViewer.PdfDocument.Load(tempWorkPath))
                        using (var outputDoc = new PdfSharp.Pdf.PdfDocument())
                        {
                            LogToFile("Starting fallback save (Rasterization).");
                            
                            // [Fix] Iterate through snapshot pages only
                            for (int i = 0; i < pagesSnapshot.Count; i++)
                            {
                                var pageData = pagesSnapshot[i];
                                int originalIdx = pageData.OriginalPageIndex;

                                if (originalIdx >= 0 && originalIdx < pdfiumDoc.PageCount)
                                {
                                    var size = pdfiumDoc.PageSizes[originalIdx];
                                    using (var bitmap = pdfiumDoc.Render(originalIdx, (int)size.Width * 2, (int)size.Height * 2, 192, 192, PdfRenderFlags.Annotations))
                                    using (var ms = new MemoryStream())
                                    {
                                        bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        ms.Position = 0;

                                        var page = outputDoc.AddPage();
                                        page.Width = size.Width;
                                        page.Height = size.Height;
                                        pageMapping[originalIdx] = page; // Register mapping

                                        using (var gfx = XGraphics.FromPdfPage(page))
                                        using (var xImage = XImage.FromStream(ms))
                                        {
                                            gfx.DrawImage(xImage, 0, 0, page.Width, page.Height);
                                            DrawOcrText(outputDoc, page, pageData.OcrWords, pageData.Width, pageData.Height);
                                            DrawAnnotationsOnPage(outputDoc, page, i, pagesSnapshot);
                                        }
                                    }
                                }
                            }

                            // Restore Bookmarks (Fallback)
                            foreach (var bm in bookmarksSnapshot)
                            {
                                AddBookmarkToPdf(outputDoc.Outlines, bm, pageMapping);
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

        private void AddBookmarkToPdf(PdfOutlineCollection collection, BookmarkSaveData bmData, Dictionary<int, PdfPage> mapping)
        {
            // Find destination page
            PdfPage? destPage = null;
            if (mapping.ContainsKey(bmData.OriginalPageIndex))
            {
                destPage = mapping[bmData.OriginalPageIndex];
            }
            else
            {
                // If the target page was deleted, try to find the nearest previous page?
                // Or just don't add destination (just a title).
                // Or map to first page if lost.
                // Let's search for nearest valid page >= OriginalIndex?
                // Actually, if page is deleted, bookmark to it might be invalid.
                // We'll skip destination if invalid, but keep title.
            }

            var outline = destPage != null ? collection.Add(bmData.Title, destPage, true) : collection.Add(bmData.Title, null, true);
            
            foreach (var child in bmData.Children)
            {
                AddBookmarkToPdf(outline.Outlines, child, mapping);
            }
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
                            // Try Noto Sans KR first, fallback to Malgun Gothic usually handled by system or font resolver
                            var fontName = "Noto Sans KR"; 
                            var font = new XFont(fontName, ann.FontSize * scaleY, ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);
                            
                            // If creation failed or needs fallback (simple check, though XFont usually doesn't throw immediately)
                            if (font.FontFamily.Name != fontName) font = new XFont("Malgun Gothic", ann.FontSize * scaleY, ann.IsBold ? XFontStyleEx.Bold : XFontStyleEx.Regular);

                            var brush = new XSolidBrush(XColor.FromArgb(ann.ForegroundColor.A, ann.ForegroundColor.R, ann.ForegroundColor.G, ann.ForegroundColor.B));
                            gfx.DrawString(ann.TextContent, font, brush, new XRect(0, 0, ann.Width * scaleX, ann.Height * scaleY), XStringFormats.TopLeft);
                        }

                        // Add Resource Dictionary (/DR) so Pdfium can find the font
                        var pdfForm = GetPdfForm(form);
                        string fontKey = "/F1"; // Default fallback
                        
                        var dr = new PdfDictionary(doc);
                        var fontDict = new PdfDictionary(doc);
                        bool fontFound = false;

                        if (pdfForm != null)
                        {
                            var resources = pdfForm.Elements.GetDictionary("/Resources");
                            if (resources != null)
                            {
                                var formFontDict = resources.Elements.GetDictionary("/Font");
                                if (formFontDict != null && formFontDict.Elements.Count > 0)
                                {
                                    foreach (var key in formFontDict.Elements.Keys)
                                    {
                                        fontDict.Elements[key] = formFontDict.Elements[key];
                                        fontKey = key;
                                        fontFound = true;
                                    }
                                }
                            }
                        }

                        // If PdfSharp didn't put the font in XForm resources, try to find it from the document
                        if (!fontFound)
                        {
                            PdfDictionary? fallbackFont = null;
                            foreach (var obj in doc.Internals.GetAllObjects())
                            {
                                if (obj is PdfDictionary d && d.Elements.GetName("/Type") == "/Font")
                                {
                                    var baseFont = d.Elements.GetName("/BaseFont");
                                    var subtype = d.Elements.GetName("/Subtype");
                                    
                                    // High priority: Korean fonts we just added
                                    if (baseFont.Contains("Noto") || baseFont.Contains("Malgun") || subtype == "/Type0")
                                    {
                                        fontDict.Elements["/F1"] = d.Reference;
                                        fontKey = "/F1";
                                        fontFound = true;
                                        break; 
                                    }
                                    
                                    // Lower priority fallback: any font
                                    if (fallbackFont == null) fallbackFont = d;
                                }
                            }
                            
                            if (!fontFound && fallbackFont != null)
                            {
                                fontDict.Elements["/F1"] = fallbackFont.Reference;
                                fontKey = "/F1";
                                fontFound = true;
                            }
                        }

                        // If STILL no font found (Clean PDF case), create a standard fallback font
                        if (!fontFound)
                        {
                            var f = new PdfDictionary(doc);
                            f.Elements["/Type"] = new PdfName("/Font");
                            f.Elements["/Subtype"] = new PdfName("/Type1");
                            f.Elements["/BaseFont"] = new PdfName("/Helvetica");
                            f.Elements["/Encoding"] = new PdfName("/WinAnsiEncoding");
                            doc.Internals.AddObject(f);
                            fontDict.Elements["/F1"] = f.Reference;
                            fontKey = "/F1";
                            fontFound = true;
                        }
                        
                        dr.Elements["/Font"] = fontDict;
                        pdfAnnot.Elements["/DR"] = dr;

                        if (pdfForm != null)
                        {
                            var apDict = new PdfDictionary(doc);
                            apDict.Elements["/N"] = pdfForm.Reference;
                            pdfAnnot.Elements["/AP"] = apDict;
                        }

                        // Use InvariantCulture to ensure dot decimal separator and prevent "Size 0" issues
                        double finalFontSize = Math.Max(1.0, ann.FontSize * scaleY);
                        string colorStr = string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} rg", 
                            ann.ForegroundColor.R / 255.0, ann.ForegroundColor.G / 255.0, ann.ForegroundColor.B / 255.0);
                        
                        pdfAnnot.Elements["/DA"] = new PdfString(string.Format(System.Globalization.CultureInfo.InvariantCulture, 
                            "{0} {1:0.###} Tf {2}", fontKey, finalFontSize, colorStr));

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
                                gfx.DrawRectangle(brush, 0, 0, ann.Width * scaleX, ann.Height * scaleY);
                            }
                            else
                            {
                                var pen = new XPen(XColors.Black, 1 * scaleY);
                                gfx.DrawLine(pen, 0, (ann.Height * scaleY) - 1, ann.Width * scaleX, (ann.Height * scaleY) - 1);
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