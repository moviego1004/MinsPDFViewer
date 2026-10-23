using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Runtime.InteropServices;
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
        public static bool HasPendingRenders => !_renderStack.IsEmpty;
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
                private static double GetDisplayDpiScale()
        {
            try
            {
                var window = Application.Current?.MainWindow;
                if (window != null)
                {
                    var dpi = VisualTreeHelper.GetDpi(window);
                    if (dpi.DpiScaleX > 0.1)
                        return dpi.DpiScaleX;
                }
            }
            catch { }
            return 1.0;
        }

        /// <summary>미리보기 패스의 해상도 비율 (고해상도 대비). 픽셀 수 1/4 → 렌더 비용도 대략 1/4.</summary>
        internal const double FastRenderScale = 0.5;

        private static void ExecuteFastRender(RenderRequest req)
        {
            var model = req.Model;
            var pageVM = req.PageVM;
            var ct = req.CancellationToken;

            if (ct.IsCancellationRequested || model.IsDisposed) return;

            // 이미 그림이 있으면(줌 변경·저장 후 재렌더 등) 흐린 미리보기는 필요 없다 → 고해상도 패스만.
            // 예전에는 두 패스가 같은 해상도라 같은 페이지를 두 번 그렸다.
            if (pageVM.ImageSource != null) return;

            BitmapSource? fastBmp = null;
            lock (PdfiumLock)
            {
                if (ct.IsCancellationRequested || model.IsDisposed || model.PdfDocument == null) return;
                try
                {
                    double scale = GetDisplayDpiScale() * FastRenderScale;
                    int fastW = Math.Max(1, (int)Math.Round(pageVM.Width * scale));
                    int fastH = Math.Max(1, (int)Math.Round(pageVM.Height * scale));
                    float dpi = (float)(96.0 * scale);
                    // 미리보기는 LCD 텍스트·후처리 없이 (곧 고해상도로 바뀐다)
                    using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, fastW, fastH, dpi, dpi, PdfRenderFlags.Annotations))
                    {
                        fastBmp = ToBitmapSource((DrawingBitmap)bitmap, applyStemDarkening: false, applySharpen: false);
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
                    double dpiScale = GetDisplayDpiScale();
                    int hiW = Math.Max(1, (int)Math.Round(pageVM.Width * dpiScale));
                    int hiH = Math.Max(1, (int)Math.Round(pageVM.Height * dpiScale));
                    float dpi = (float)(96.0 * dpiScale);
                    var flags = PdfRenderFlags.Annotations | PdfRenderFlags.LcdText | PdfRenderFlags.CorrectFromDpi;
                    using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, hiW, hiH, dpi, dpi, flags))
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
                string cacheKey = GetCacheKey(model.FilePath, pageVM.OriginalPageIndex, req.Zoom);
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

        // Page image cache (key: file identity + pageIndex + zoom, value: frozen BitmapSource)
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

        private static string GetCacheKey(string filePath, int pageIndex, double zoom)
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

            int zoomBucket = (int)Math.Round(zoom * 100);
            return $"{normalizedPath}|{lastWriteTicks}|{length}|{pageIndex}|{zoomBucket}";
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
            public static extern IntPtr FPDF_LoadMemDocument(IntPtr data_buf, int size, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPStr)] string? password);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern void FPDF_CloseDocument(IntPtr document);

            [System.Runtime.InteropServices.DllImport("pdfium.dll", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
            public static extern int FPDF_GetLastError();

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

        public List<Rect> FindTextRects(string filePath, int pageIndex, string keyword, string? password = null)
        {
            var results = new List<Rect>();
            if (!File.Exists(filePath)) return results;

            lock (PdfiumLock)
            {
                IntPtr doc = NativeMethods.FPDF_LoadDocument(filePath, password);
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

        public string ExtractTextInRect(string filePath, int pageIndex, Rect uiRect, double viewWidth, double viewHeight, double pdfWidth, double pdfHeight, string? password = null)
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
                IntPtr doc = NativeMethods.FPDF_LoadDocument(filePath, password);
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

                private static IPdfDocument LoadPdfiumDocument(Stream stream, string? password)
        {
            if (string.IsNullOrEmpty(password))
                return PdfiumViewer.PdfDocument.Load(stream);

            var method = typeof(PdfiumViewer.PdfDocument).GetMethod("Load", new[] { typeof(Stream), typeof(string) });
            if (method == null)
                throw new InvalidOperationException("PdfDocument.Load(Stream, string) method not found.");
            return (IPdfDocument)method.Invoke(null, new object[] { stream, password })!;
        }

private static string? PromptPasswordDialog(string? prompt)
        {
            string? result = null;
            RunOnUiDispatcherOrDirect(() =>
            {
                var owner = Application.Current?.MainWindow;
                var dlg = new PasswordDialog(prompt, owner);
                if (dlg.ShowDialog() == true)
                {
                    result = dlg.PasswordText;
                }
            });
            return result;
        }

        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath, string? password = null)
        {
            if (!File.Exists(filePath)) return null;
            return await Task.Run(() =>
            {
                var profileWatch = Stopwatch.StartNew();
                void LogOpenProfile(string step)
                {
                    Log($"[OPEN_PROFILE] {step} elapsed={profileWatch.ElapsedMilliseconds}ms path='{filePath}'");
                    profileWatch.Restart();
                }

                var fileInfo = new FileInfo(filePath);
                Log($"LoadPdfAsync: {filePath}, Bytes={fileInfo.Length}");
                LogOpenProfile("stat file");

                byte[] pdfBytes = File.ReadAllBytes(filePath);
                LogOpenProfile("read file bytes");

                string? currentPassword = password;
                bool passwordPrompted = false;
                IPdfDocument? pdfDoc = null;
                MemoryStream? pdfStream = null;

                while (true)
                {
                    try
                    {
                        pdfStream?.Dispose();
                        pdfStream = new MemoryStream(pdfBytes);
                        lock (PdfiumLock)
                        {
                            pdfDoc = LoadPdfiumDocument(pdfStream, currentPassword);
                        }
                        LogOpenProfile("PdfDocument.Load stream");
                        break;
                    }
                    catch (Exception ex)
                    {
                        bool isPasswordError = ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!isPasswordError)
                        {
                            lock (PdfiumLock)
                            {
                                if (NativeMethods.FPDF_GetLastError() == 4) // FPDF_ERR_PASSWORD
                                    isPasswordError = true;
                            }
                        }

                        if (isPasswordError)
                        {
                            string prompt = passwordPrompted
                                ? "암호가 올바르지 않습니다. 다시 입력해 주세요."
                                : "문서가 암호로 보호되어 있습니다.\n비밀번호를 입력해 주세요.";
                            string? entered = PromptPasswordDialog(prompt);
                            if (entered == null)
                            {
                                pdfStream?.Dispose();
                                return null;
                            }
                            currentPassword = entered;
                            passwordPrompted = true;
                        }
                        else
                        {
                            pdfStream?.Dispose();
                            MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                            return null;
                        }
                    }
                }

                var model = new PdfDocumentModel
                {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath),
                    PdfDocument = pdfDoc,
                    FileStream = pdfStream,
                    SourceBytes = pdfBytes,
                    Password = currentPassword
                };
                LogOpenProfile("create model");
                return model;
            });
        }

        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.PdfDocument == null || model.IsDisposed) return;
            await Task.Run(() =>
            {
                if (model.IsDisposed) return;
                int pageCount = 0;
                var tempPageList = new List<PdfPageViewModel>();

                lock (PdfiumLock)
                {
                    if (model.PdfDocument != null)
                    {
                        pageCount = model.PdfDocument.PageCount;
                        if (pageCount > 0)
                        {
                            var sizes = model.PdfDocument.PageSizes;
                            for (int i = 0; i < pageCount; i++)
                            {
                                var size = sizes[i];
                                double pageW = size.Width;
                                double pageH = size.Height;
                                tempPageList.Add(new PdfPageViewModel
                                {
                                    PageIndex = i,
                                    OriginalFilePath = model.FilePath,
                                    OriginalPageIndex = i,
                                    Width = pageW * model.Zoom,
                                    Height = pageH * model.Zoom,
                                    PdfPageWidthPoint = pageW,
                                    PdfPageHeightPoint = pageH,
                                    CropWidthPoint = pageW,
                                    CropHeightPoint = pageH
                                });
                            }
                        }
                    }
                }

                if (tempPageList.Count == 0) return;
                Log($"InitializeDocumentAsync: {model.FilePath}, PageCount={pageCount}");

                RunOnUiDispatcherOrDirect(() => { foreach (var p in tempPageList) model.Pages.Add(p); });
            });
        }

        private static void RunOnUiDispatcherOrDirect(Action action)
        {
            if (UseDirectDispatcherForTests)
            {
                action();
                return;
            }

            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            dispatcher.Invoke(action);
        }

        private static async Task RunOnUiDispatcherOrDirectAsync(Action action)
        {
            if (UseDirectDispatcherForTests)
            {
                action();
                return;
            }

            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }

            await dispatcher.InvokeAsync(action);
        }

        internal static bool UseDirectDispatcherForTests { get; set; }

        private static void Log(string message)
        {
            try
            {
                DebugLog.Append("PdfService", message);
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
            string cacheKey = GetCacheKey(model.FilePath, pageVM.OriginalPageIndex, model.Zoom);
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

        public async Task RenderThumbnailAsync(PdfDocumentModel model, PdfPageViewModel pageVM, int maxPixelWidth = 140, CancellationToken ct = default)
        {
            if (model.IsDisposed || pageVM.IsBlankPage || pageVM.ThumbnailSource != null)
                return;

            BitmapSource? thumbnail = null;
            await Task.Run(() =>
            {
                if (ct.IsCancellationRequested || model.IsDisposed || model.PdfDocument == null)
                    return;

                lock (PdfiumLock)
                {
                    if (ct.IsCancellationRequested || model.IsDisposed || model.PdfDocument == null)
                        return;

                    try
                    {
                        double widthPoint = pageVM.PdfPageWidthPoint > 0 ? pageVM.PdfPageWidthPoint : pageVM.Width;
                        double heightPoint = pageVM.PdfPageHeightPoint > 0 ? pageVM.PdfPageHeightPoint : pageVM.Height;
                        double scale = maxPixelWidth / Math.Max(1.0, widthPoint);
                        int renderW = Math.Max(1, (int)Math.Round(widthPoint * scale));
                        int renderH = Math.Max(1, (int)Math.Round(heightPoint * scale));

                        using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, 96, 96, PdfRenderFlags.Annotations))
                        {
                            thumbnail = ToBitmapSource((DrawingBitmap)bitmap);
                            thumbnail.Freeze();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log($"[WARN] Thumbnail render failed for page {pageVM.PageIndex}: {ex.Message}");
                    }
                }
            }, ct);

            if (thumbnail == null || ct.IsCancellationRequested)
                return;

            RunOnUiDispatcherOrDirect(() =>
            {
                if (!ct.IsCancellationRequested && pageVM.ThumbnailSource == null)
                    pageVM.ThumbnailSource = thumbnail;
            });
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
                            double dpiScale = GetDisplayDpiScale();
                            int renderW = Math.Max(1, (int)Math.Round(pageVM.Width * dpiScale));
                            int renderH = Math.Max(1, (int)Math.Round(pageVM.Height * dpiScale));
                            float dpi = (float)(96.0 * dpiScale);
                            var flags = PdfRenderFlags.Annotations | PdfRenderFlags.LcdText | PdfRenderFlags.CorrectFromDpi;
                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, dpi, dpi, flags))
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

        private static readonly byte[] GammaLut = CreateGammaLut(1.20);

        private static byte[] CreateGammaLut(double gamma)
        {
            var lut = new byte[256];
            for (int i = 0; i < 256; i++)
            {
                double normalized = i / 255.0;
                double adjusted = Math.Pow(normalized, gamma);
                lut[i] = (byte)Math.Clamp(Math.Round(adjusted * 255.0), 0, 255);
            }
            return lut;
        }

        /// <summary>
        /// 밝기 언샤프 마스크 강도 (0 = 끔). 스캔 PDF(1비트 600dpi 이미지)를 화면 크기로 줄이면 획 가장자리가 회색으로 퍼져
        /// Foxit보다 흐려 보인다. 감마 LUT 뒤에 밝기(Y)에만 3×3 가우시안 언샤프를 더해 획을 세운다.
        /// 색 차이는 건드리지 않아 LCD 서브픽셀 색번짐을 키우지 않는다.
        /// </summary>
        public static double SharpenAmount { get; set; } = 0.6;

        public static BitmapSource ToBitmapSource(DrawingBitmap bitmap, bool applyStemDarkening = true, bool applySharpen = true)
        {
            var rect = new DrawingRectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadWrite, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                if (applyStemDarkening)
                {
                    unsafe
                    {
                        byte* scan0 = (byte*)bitmapData.Scan0;
                        int stride = bitmapData.Stride;
                        int height = bitmap.Height;
                        int width = bitmap.Width;

                        if (stride == width * 4)
                        {
                            int totalBytes = stride * height;
                            for (int i = 0; i < totalBytes; i += 4)
                            {
                                scan0[i] = GammaLut[scan0[i]];         // B
                                scan0[i + 1] = GammaLut[scan0[i + 1]]; // G
                                scan0[i + 2] = GammaLut[scan0[i + 2]]; // R
                            }
                        }
                        else
                        {
                            for (int y = 0; y < height; y++)
                            {
                                byte* row = scan0 + (y * stride);
                                for (int x = 0; x < width; x++)
                                {
                                    int idx = x * 4;
                                    row[idx] = GammaLut[row[idx]];         // B
                                    row[idx + 1] = GammaLut[row[idx + 1]]; // G
                                    row[idx + 2] = GammaLut[row[idx + 2]]; // R
                                }
                            }
                        }
                    }
                }

                // 감마 다음에 샤픈 (감마로 진해진 획의 가장자리를 세운다)
                if (applySharpen && SharpenAmount > 0.001)
                {
                    unsafe
                    {
                        SharpenLuma((byte*)bitmapData.Scan0, bitmap.Width, bitmap.Height, bitmapData.Stride, SharpenAmount);
                    }
                }

                return BitmapSource.Create(
                    bitmap.Width,
                    bitmap.Height,
                    bitmap.HorizontalResolution,
                    bitmap.VerticalResolution,
                    PixelFormats.Bgra32,
                    null,
                    bitmapData.Scan0,
                    bitmapData.Stride * bitmap.Height,
                    bitmapData.Stride);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }

        /// <summary>
        /// BGRA 픽셀의 밝기 Y에 언샤프 마스크: Y' = Y + amount·(Y − blur3x3(Y)), 차이 ΔY를 B·G·R에 똑같이 더한다.
        /// blur = [1 2 1]ᵀ[1 2 1]/16 (σ≈0.85px). 균일한 면(흰 바탕, 검은 획 안쪽)은 그대로이고 가장자리만 대비가 커진다.
        /// </summary>
        internal static unsafe void SharpenLuma(byte* scan0, int width, int height, int stride, double amount)
        {
            if (width < 3 || height < 3 || amount <= 0)
                return;

            int n = width * height;
            int[] luma = System.Buffers.ArrayPool<int>.Shared.Rent(n);
            int[] hsum = System.Buffers.ArrayPool<int>.Shared.Rent(n);
            try
            {
                for (int y = 0; y < height; y++)
                {
                    byte* row = scan0 + (long)y * stride;
                    int o = y * width;
                    for (int x = 0; x < width; x++)
                    {
                        byte* px = row + x * 4;
                        luma[o + x] = (px[2] * 77 + px[1] * 150 + px[0] * 29) >> 8; // R G B
                    }
                }

                // 가로 [1 2 1] (가장자리는 복제)
                for (int y = 0; y < height; y++)
                {
                    int o = y * width;
                    hsum[o] = luma[o] * 3 + luma[o + 1];
                    for (int x = 1; x < width - 1; x++)
                        hsum[o + x] = luma[o + x - 1] + 2 * luma[o + x] + luma[o + x + 1];
                    hsum[o + width - 1] = luma[o + width - 2] + luma[o + width - 1] * 3;
                }

                int a256 = (int)Math.Round(amount * 256);
                for (int y = 0; y < height; y++)
                {
                    int up = Math.Max(0, y - 1) * width;
                    int mid = y * width;
                    int down = Math.Min(height - 1, y + 1) * width;
                    byte* row = scan0 + (long)y * stride;
                    for (int x = 0; x < width; x++)
                    {
                        int blur16 = hsum[up + x] + 2 * hsum[mid + x] + hsum[down + x]; // = 16 × blur
                        int diff16 = (luma[mid + x] << 4) - blur16;
                        if (diff16 == 0)
                            continue;
                        int delta = (a256 * diff16) / (256 * 16);
                        if (delta == 0)
                            continue;
                        byte* px = row + x * 4;
                        px[0] = ClampByte(px[0] + delta);
                        px[1] = ClampByte(px[1] + delta);
                        px[2] = ClampByte(px[2] + delta);
                    }
                }
            }
            finally
            {
                System.Buffers.ArrayPool<int>.Shared.Return(luma);
                System.Buffers.ArrayPool<int>.Shared.Return(hsum);
            }
        }

        private static byte ClampByte(int v) => (byte)(v < 0 ? 0 : (v > 255 ? 255 : v));

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
                    var extracted = new List<PdfAnnotation>();
                    byte[]? sourceBytes = model.SourceBytes;
                    GCHandle sourceHandle = default;
                    IntPtr doc = IntPtr.Zero;

                    if (sourceBytes != null && sourceBytes.Length > 0)
                    {
                        sourceHandle = GCHandle.Alloc(sourceBytes, GCHandleType.Pinned);
                        doc = NativeMethods.FPDF_LoadMemDocument(sourceHandle.AddrOfPinnedObject(), sourceBytes.Length, model.Password);
                    }
                    else
                    {
                        string path = pageVM.OriginalFilePath ?? model.FilePath;
                        if (!File.Exists(path)) return;
                        doc = NativeMethods.FPDF_LoadDocument(path, model.Password);
                    }

                    if (doc == IntPtr.Zero)
                    {
                        if (sourceHandle.IsAllocated)
                            sourceHandle.Free();
                        return;
                    }

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
                        if (sourceHandle.IsAllocated)
                            sourceHandle.Free();
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
                    DebugLog.Append("PdfService", msg);
                }
                catch { }
            }
        }

        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null) return;

            string originalFilePath = model.FilePath;
            LogToFile($"Starting SavePdf for: {originalFilePath}");
            var profileWatch = Stopwatch.StartNew();
            void LogProfile(string step)
            {
                LogToFile($"[SAVE_PROFILE] {step} elapsed={profileWatch.ElapsedMilliseconds}ms");
                profileWatch.Restart();
            }

            bool shouldRewriteBookmarks = false;
            await RunOnUiDispatcherOrDirectAsync(() =>
            {
                shouldRewriteBookmarks = model.HasBookmarkChanges;
            });
            LogProfile("read bookmark state");

            List<MinsPDFViewer.BookmarkSaveData> pdfiumBookmarkSnapshot = shouldRewriteBookmarks
                ? await PdfSaveSnapshotFactory.CreateBookmarkSnapshotAsync(model)
                : new List<MinsPDFViewer.BookmarkSaveData>();
            LogProfile("bookmark snapshot");

            try
            {
                var pdfiumEditService = new PdfiumEditService();
                string pdfiumUnsupportedReason = PdfiumEditService.GetUnsupportedReason(model);
                LogProfile("eligibility check");
                if (!string.IsNullOrEmpty(pdfiumUnsupportedReason))
                    LogToFile($"[SaveEngine:{PdfSaveEngine.Pdfium}] Not eligible: {pdfiumUnsupportedReason}");

                if (await pdfiumEditService.TrySaveAnnotationsAsync(model, outputPath))
                {
                    LogProfile("pdfium edit save");
                    // 저장본의 페이지 순서 = model.Pages 순서. 좌표는 포인트(원본 화면 기준), 앱 내 회전은 따로 넘긴다.
                    var ocrPages = model.Pages
                        .Select((p, index) => (Page: p, Index: index))
                        .Where(x => !x.Page.IsBlankPage && x.Page.OcrWords != null && x.Page.OcrWords.Count > 0)
                        .Select(x => new OcrPageData
                        {
                            PageIndex = x.Index,
                            UserRotation = x.Page.Rotation,
                            WidthPoint = x.Page.SourceViewWidthPoint,
                            HeightPoint = x.Page.SourceViewHeightPoint,
                            Words = new List<OcrWordInfo>(x.Page.OcrWords!)
                        }).ToList();

                    if (ocrPages.Count > 0)
                    {
                        bool injected = await new PdfGlyphlessOcrService(LogToFile).InjectOcrLayerAsync(outputPath, ocrPages);
                        LogProfile($"glyphless ocr layer inject injected={injected}");
                    }

                    if (shouldRewriteBookmarks)
                    {
                        await new PdfBookmarkRewriteService(LogToFile).RewriteBookmarksAsync(outputPath, pdfiumBookmarkSnapshot);
                        LogProfile("bookmark rewrite");
                        await RunOnUiDispatcherOrDirectAsync(() => model.HasBookmarkChanges = false);
                        LogProfile("clear bookmark dirty flag");
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

        public Task LoadBookmarksAsync(PdfDocumentModel model)
        {
            return Task.Run(() =>
            {
                if (model.IsDisposed)
                    return;

                var profileWatch = Stopwatch.StartNew();
                try
                {
                    LoadBookmarks(model);
                    Log($"[OPEN_PROFILE] LoadBookmarksAsync elapsed={profileWatch.ElapsedMilliseconds}ms path='{model.FilePath}'");
                }
                catch (Exception ex)
                {
                    Log($"[WARN] LoadBookmarksAsync failed: {ex.Message}");
                }
            });
        }

        private PdfBookmarkViewModel MapBm(PdfBookmark b, PdfBookmarkViewModel? p)
        {
            var v = new PdfBookmarkViewModel { Title = b.Title, PageIndex = b.PageIndex, Parent = p };
            foreach (var c in b.Children) v.Children.Add(MapBm(c, v));
            return v;
        }
    }
}
