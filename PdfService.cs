using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfiumViewer;
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

        // Page image cache (key: file identity + pageIndex, value: frozen BitmapSource)
        private static readonly LinkedList<string> _cacheOrder = new LinkedList<string>();
        private static readonly Dictionary<string, BitmapSource> _pageCache = new Dictionary<string, BitmapSource>();
        private static readonly object _cacheLock = new object();

        private static string NormalizeCachePath(string filePath)
        {
            try
            {
                return Path.GetFullPath(filePath);
            }
            catch
            {
                return filePath;
            }
        }

        private static string GetCacheKey(string filePath, int pageIndex)
        {
            string normalizedPath = NormalizeCachePath(filePath);
            long lastWriteTicks = 0;
            long length = 0;

            try
            {
                var info = new FileInfo(normalizedPath);
                if (info.Exists)
                {
                    lastWriteTicks = info.LastWriteTimeUtc.Ticks;
                    length = info.Length;
                }
            }
            catch { }

            return $"{normalizedPath}|{lastWriteTicks}|{length}|{pageIndex}";
        }

        public static void ClearPageCacheForFile(string filePath)
        {
            string normalizedPath = NormalizeCachePath(filePath);
            string currentPrefix = normalizedPath + "|";
            string legacyPrefix = filePath + "_";
            int removed = 0;

            lock (_cacheLock)
            {
                var keysToRemove = _pageCache.Keys
                    .Where(k =>
                        k.StartsWith(currentPrefix, StringComparison.OrdinalIgnoreCase) ||
                        k.StartsWith(legacyPrefix, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                foreach (var key in keysToRemove)
                {
                    if (_pageCache.Remove(key))
                    {
                        _cacheOrder.Remove(key);
                        removed++;
                    }
                }
            }

            if (removed > 0)
                Log($"Cleared {removed} cached page image(s) for {normalizedPath}");
        }

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
            public const int FPDF_ANNOT_STAMP = 13;

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

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDFText_GetBoundedText(IntPtr text_page, double left, double top, double right, double bottom, IntPtr buffer, int buflen);

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
                            IntPtr search = NativeMethods.FPDFText_FindStart(textPage, keyword, 0, 0);
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

        public string ExtractTextInRect(string filePath, int pageIndex, Rect uiRect, double viewWidth, double viewHeight, double pdfWidth, double pdfHeight)
        {
            if (!File.Exists(filePath) || uiRect.Width <= 0 || uiRect.Height <= 0 || viewWidth <= 0 || viewHeight <= 0)
                return string.Empty;

            double scaleX = pdfWidth / viewWidth;
            double scaleY = pdfHeight / viewHeight;
            double left = uiRect.Left * scaleX;
            double right = uiRect.Right * scaleX;
            double top = pdfHeight - (uiRect.Top * scaleY);
            double bottom = pdfHeight - (uiRect.Bottom * scaleY);

            lock (PdfiumLock)
            {
                IntPtr doc = NativeMethods.FPDF_LoadDocument(filePath, null);
                if (doc == IntPtr.Zero) return string.Empty;

                try
                {
                    IntPtr page = NativeMethods.FPDF_LoadPage(doc, pageIndex);
                    if (page == IntPtr.Zero) return string.Empty;

                    try
                    {
                        IntPtr textPage = NativeMethods.FPDFText_LoadPage(page);
                        if (textPage == IntPtr.Zero) return string.Empty;

                        try
                        {
                            int len = NativeMethods.FPDFText_GetBoundedText(textPage, left, top, right, bottom, IntPtr.Zero, 0);
                            if (len <= 0) return string.Empty;

                            IntPtr buffer = System.Runtime.InteropServices.Marshal.AllocHGlobal((len + 1) * 2);
                            try
                            {
                                int written = NativeMethods.FPDFText_GetBoundedText(textPage, left, top, right, bottom, buffer, len + 1);
                                if (written <= 0) return string.Empty;
                                return System.Runtime.InteropServices.Marshal.PtrToStringUni(buffer, written)?.Trim() ?? string.Empty;
                            }
                            finally
                            {
                                System.Runtime.InteropServices.Marshal.FreeHGlobal(buffer);
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
        }

        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath)) return null;
            return await Task.Run(() =>
            {
                try
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    Log($"LoadPdfAsync: {filePath}, Bytes={fileBytes.Length}");
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
                Log($"InitializeDocumentAsync: {model.FilePath}, PageCount={pageCount}");

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
                RunOnUiDispatcherOrDirect(() => { foreach (var p in tempPageList) model.Pages.Add(p); });
            });
        }

        private static void RunOnUiDispatcherOrDirect(Action action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            var operation = dispatcher.InvokeAsync(action);
            if (!operation.Task.Wait(TimeSpan.FromSeconds(5)))
                action();
        }

        private static async Task RunOnUiDispatcherOrDirectAsync(Action action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            var operation = dispatcher.InvokeAsync(action);
            if (await Task.WhenAny(operation.Task, Task.Delay(TimeSpan.FromSeconds(5))) == operation.Task)
                await operation.Task;
            else
                action();
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
                RunOnUiDispatcherOrDirect(() =>
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
                if (bmpSource != null) RunOnUiDispatcherOrDirect(() => pageVM.ImageSource = bmpSource);
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
                                        subtype == NativeMethods.FPDF_ANNOT_WIDGET ||
                                        subtype == NativeMethods.FPDF_ANNOT_STAMP)
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
                                        string fontFamily = "Malgun Gothic";
                                        bool isBold = false;

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
                                        else if (subtype == NativeMethods.FPDF_ANNOT_STAMP)
                                        {
                                            if (PdfiumEditService.TryDecodeManagedFreeText(content, out PdfiumEditService.ManagedFreeTextMetadata managedFreeText))
                                            {
                                                annType = AnnotationType.FreeText;
                                                content = managedFreeText.Text;
                                                fSize = managedFreeText.FontSize;
                                                fontFamily = managedFreeText.FontFamily;
                                                isBold = managedFreeText.IsBold;
                                                var managedBrush = new SolidColorBrush(managedFreeText.ForegroundColor);
                                                managedBrush.Freeze();
                                                brush = managedBrush;
                                            }
                                            else if (PdfiumEditService.TryExtractManagedImageStampBytes(annot, content, out byte[] imageBytes))
                                            {
                                                annType = AnnotationType.ImageStamp;
                                                content = string.Empty;
                                                extracted.Add(new PdfAnnotation
                                                {
                                                    Type = annType,
                                                    X = uiX,
                                                    Y = uiY,
                                                    Width = uiW,
                                                    Height = uiH,
                                                    TextContent = content,
                                                    FontSize = fSize,
                                                    FontFamily = fontFamily,
                                                    IsBold = isBold,
                                                    Foreground = brush,
                                                    Background = background,
                                                    FieldName = fieldName,
                                                    ImageBytes = imageBytes,
                                                    ImageSource = CreateBitmapSource(imageBytes, uiW, uiH)
                                                });
                                                continue;
                                            }
                                            else if (PdfiumEditService.TryDecodeManagedFreeText(content, out string decodedText))
                                            {
                                                annType = AnnotationType.FreeText;
                                                content = decodedText;
                                            }
                                            else
                                            {
                                                continue;
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
                                            FontFamily = fontFamily,
                                            IsBold = isBold,
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
                        RunOnUiDispatcherOrDirect(() =>
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

            bool shouldRewriteBookmarks = false;
            await RunOnUiDispatcherOrDirectAsync(() =>
            {
                shouldRewriteBookmarks = model.HasBookmarkChanges;
            });

            List<MinsPDFViewer.BookmarkSaveData> pdfiumBookmarkSnapshot = shouldRewriteBookmarks
                ? await PdfSaveSnapshotFactory.CreateBookmarkSnapshotAsync(model)
                : new List<MinsPDFViewer.BookmarkSaveData>();

            try
            {
                var pdfiumEditService = new PdfiumEditService();
                string pdfiumUnsupportedReason = PdfiumEditService.GetUnsupportedReason(model);
                if (!string.IsNullOrEmpty(pdfiumUnsupportedReason))
                    LogToFile($"[SaveEngine:{PdfSaveEngine.Pdfium}] Not eligible: {pdfiumUnsupportedReason}");

                if (await pdfiumEditService.TrySaveAnnotationsAsync(model, outputPath))
                {
                    if (shouldRewriteBookmarks)
                    {
                        await new PdfBookmarkRewriteService(LogToFile).RewriteBookmarksAsync(outputPath, pdfiumBookmarkSnapshot);
                        await RunOnUiDispatcherOrDirectAsync(() => model.HasBookmarkChanges = false);
                        LogToFile($"[SaveEngine:{PdfSaveEngine.PdfiumWithBookmarkRewrite}] Save completed.");
                        LogToFile("PDFium edit save successful.");
                        return;
                    }
                    LogToFile($"[SaveEngine:{PdfSaveEngine.Pdfium}] Save completed.");
                    LogToFile("PDFium edit save successful.");
                    return;
                }

                throw new InvalidOperationException(
                    string.IsNullOrEmpty(pdfiumUnsupportedReason)
                        ? "PDFium save did not complete."
                        : $"PDFium save is not available: {pdfiumUnsupportedReason}");
            }
            catch (Exception ex)
            {
                LogToFile($"[SaveEngine:{PdfSaveEngine.Pdfium}] Save failed: {ex.Message}");
                throw;
            }
        }

        private static BitmapSource? CreateBitmapSource(byte[] imageBytes, double displayWidth, double displayHeight)
        {
            if (imageBytes.Length == 0)
                return null;

            try
            {
                var bitmap = new BitmapImage();
                using var stream = new MemoryStream(imageBytes);
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                int decodeWidth = GetDecodePixelWidth(displayWidth, displayHeight);
                if (decodeWidth > 0)
                    bitmap.DecodePixelWidth = decodeWidth;
                bitmap.StreamSource = stream;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        private static int GetDecodePixelWidth(double displayWidth, double displayHeight)
        {
            double maxDisplaySide = Math.Max(displayWidth, displayHeight);
            if (maxDisplaySide <= 0)
                return 0;

            return (int)Math.Max(64, Math.Min(2048, Math.Ceiling(maxDisplaySide * 2)));
        }

        public void LoadBookmarks(PdfDocumentModel model)
        {
            lock (PdfiumLock)
            {
                if (model.PdfDocument?.Bookmarks == null) return;
                var list = new System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>();
                foreach (var b in model.PdfDocument.Bookmarks) list.Add(MapBm(b, null));
                RunOnUiDispatcherOrDirect(() =>
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
