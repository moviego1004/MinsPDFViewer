using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;

namespace MinsPDFViewer
{
    /// <summary>
    /// 문서 열기·초기화·주석 읽기·텍스트(검색·추출)·북마크·저장 조율. 화면 렌더링은 <see cref="PageRenderScheduler"/>.
    /// PDFium은 <see cref="PdfiumDocument"/> 하나(문서당 핸들 하나)로만 접근한다.
    /// </summary>
    public class PdfService
    {
        /// <summary>모든 PDFium 호출을 직렬화하는 전역 잠금.</summary>
        public static object PdfiumLock => Pdfium.Lock;

        /// <summary>화면 렌더 작업이 남아 있는지 (썸네일 등 낮은 우선순위 작업이 양보할 때 확인).</summary>
        public static bool HasPendingRenders => PageRenderScheduler.Instance.HasPendingWork;

        private static readonly object _logLock = new object();

        /// <summary>저장 후 다시 열 때 이전 그림이 보이지 않도록 렌더 캐시를 비운다.</summary>
        public static void ClearPageCacheForFile(string filePath) => PageRenderScheduler.Instance.ClearCache();

        internal static bool UseDirectDispatcherForTests { get; set; }

        internal static int ToPdfiumRotate(int rotationDegrees) => (((rotationDegrees % 360) + 360) % 360) / 90;

        // ── 열기 ─────────────────────────────────────────────────────────

        private static string? PromptPasswordDialog(string? prompt)
        {
            string? result = null;
            RunOnUiDispatcherOrDirect(() =>
            {
                var owner = Application.Current?.MainWindow;
                var dlg = new PasswordDialog(prompt, owner);
                if (dlg.ShowDialog() == true)
                    result = dlg.PasswordText;
            });
            return result;
        }

        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath, string? password = null)
        {
            if (!File.Exists(filePath)) return null;
            return await Task.Run(() =>
            {
                var watch = Stopwatch.StartNew();
                string? currentPassword = password;
                bool prompted = false;
                PdfiumDocument? document;

                while (true)
                {
                    try
                    {
                        document = PdfiumDocument.Open(filePath, currentPassword);
                        break;
                    }
                    catch (PdfiumException ex) when (ex.IsPasswordError)
                    {
                        string prompt = prompted
                            ? "암호가 올바르지 않습니다. 다시 입력해 주세요."
                            : "문서가 암호로 보호되어 있습니다.\n비밀번호를 입력해 주세요.";
                        string? entered = PromptPasswordDialog(prompt);
                        if (entered == null)
                            return null;
                        currentPassword = entered;
                        prompted = true;
                    }
                    catch (Exception ex)
                    {
                        Log($"[ERROR] Load failed: {ex.Message}");
                        RunOnUiDispatcherOrDirect(() => MessageBox.Show($"PDF 로드 실패: {ex.Message}"));
                        return null;
                    }
                }

                Log($"[OPEN_PROFILE] Open '{filePath}' pages={document.PageCount} elapsed={watch.ElapsedMilliseconds}ms");
                return new PdfDocumentModel
                {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath),
                    Document = document,
                    Password = currentPassword
                };
            });
        }

        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.Document == null || model.IsDisposed) return;
            await Task.Run(() =>
            {
                var pages = new List<PdfPageViewModel>();
                lock (Pdfium.Lock)
                {
                    var doc = model.Document;
                    if (doc == null || doc.IsDisposed)
                        return;
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var (w, h) = doc.GetPageSize(i);
                        pages.Add(new PdfPageViewModel
                        {
                            PageIndex = i,
                            OriginalFilePath = model.FilePath,
                            OriginalPageIndex = i,
                            Width = w * model.Zoom * PdfDocumentModel.DipPerPoint,
                            Height = h * model.Zoom * PdfDocumentModel.DipPerPoint,
                            PdfPageWidthPoint = w,
                            PdfPageHeightPoint = h,
                            CropWidthPoint = w,
                            CropHeightPoint = h
                        });
                    }
                }

                if (pages.Count == 0) return;
                Log($"InitializeDocumentAsync: {model.FilePath}, PageCount={pages.Count}");
                RunOnUiDispatcherOrDirect(() => { foreach (var p in pages) model.Pages.Add(p); });
            });
        }

        // ── 렌더링 도우미 ────────────────────────────────────────────────

        /// <summary>
        /// 페이지를 fullWidth×fullHeight로 그렸을 때 region 부분을 WPF 비트맵으로 (Pdfium.Lock은 내부에서 잡음).
        /// postProcess: 화면용 감마·샤픈. cancel이 true를 돌려주면 null.
        /// </summary>
        internal static BitmapSource? RenderBitmap(PdfiumDocument document, int pageIndex, int fullWidth, int fullHeight, int rotate,
            int flags, PixelRegion region, bool postProcess, double dpiScale = 1.0, Func<bool>? cancel = null)
        {
            byte[]? buffer;
            lock (Pdfium.Lock)
            {
                if (document.IsDisposed)
                    return null;
                buffer = document.RenderToBuffer(pageIndex, fullWidth, fullHeight, rotate, flags, region, cancel);
            }
            if (buffer == null)
                return null;

            if (postProcess)
                PostProcessForScreen(buffer, region.Width, region.Height);

            double dpi = 96.0 * Math.Max(0.1, dpiScale);
            var bitmap = BitmapSource.Create(region.Width, region.Height, dpi, dpi, PixelFormats.Bgr32, null, buffer, region.Width * 4);
            bitmap.Freeze();
            return bitmap;
        }

        /// <summary>화면 표시용 후처리: 감마 LUT(획 진하게) → 밝기 언샤프(가장자리 세우기).</summary>
        internal static unsafe void PostProcessForScreen(byte[] bgra, int width, int height)
        {
            fixed (byte* p = bgra)
            {
                int total = width * height * 4;
                for (int i = 0; i < total; i += 4)
                {
                    p[i] = GammaLut[p[i]];
                    p[i + 1] = GammaLut[p[i + 1]];
                    p[i + 2] = GammaLut[p[i + 2]];
                }
                if (SharpenAmount > 0.001)
                    SharpenLuma(p, width, height, width * 4, SharpenAmount);
            }
        }

        internal static int ScreenRenderFlags => PdfiumNative.FPDF_ANNOT | PdfiumNative.FPDF_LCD_TEXT;

        /// <summary>미리 그리기 요청 (보이는 영역 계산과 별개로 이 페이지도 그려 둔다).</summary>
        public void RenderPageAsync(PdfDocumentModel model, PdfPageViewModel pageVM, CancellationToken ct)
        {
            if (model.IsDisposed || pageVM.IsBlankPage) return;
            PageRenderScheduler.Instance.RequestPage(model, pageVM);
        }

        public async Task RenderThumbnailAsync(PdfDocumentModel model, PdfPageViewModel pageVM, int maxPixelWidth = 140, CancellationToken ct = default)
        {
            if (model.IsDisposed || pageVM.IsBlankPage || pageVM.ThumbnailSource != null || model.Document == null)
                return;

            var document = model.Document;
            BitmapSource? thumbnail = await Task.Run(() =>
            {
                if (ct.IsCancellationRequested || model.IsDisposed)
                    return null;
                int rotate = ToPdfiumRotate(pageVM.Rotation);
                double w = pageVM.PdfPageWidthPoint > 0 ? pageVM.PdfPageWidthPoint : pageVM.Width;
                double h = pageVM.PdfPageHeightPoint > 0 ? pageVM.PdfPageHeightPoint : pageVM.Height;
                double scale = maxPixelWidth / Math.Max(1.0, w);
                int pw = Math.Max(1, (int)Math.Round(w * scale));
                int ph = Math.Max(1, (int)Math.Round(h * scale));
                try
                {
                    return RenderBitmap(document, pageVM.OriginalPageIndex, pw, ph, rotate, PdfiumNative.FPDF_ANNOT,
                        PixelRegion.Full(pw, ph), postProcess: false, cancel: () => ct.IsCancellationRequested || model.IsDisposed);
                }
                catch (Exception ex)
                {
                    Log($"[WARN] Thumbnail render failed for page {pageVM.PageIndex}: {ex.Message}");
                    return null;
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

        /// <summary>동기 렌더 + 주석 읽기 (테스트·저장 전처리용). 화면은 스케줄러를 쓴다.</summary>
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage || model.Document == null) return;
            if (pageVM.ImageSource == null)
            {
                try
                {
                    int w = Math.Max(1, (int)Math.Round(pageVM.Width));
                    int h = Math.Max(1, (int)Math.Round(pageVM.Height));
                    var bitmap = RenderBitmap(model.Document, pageVM.OriginalPageIndex, w, h, ToPdfiumRotate(pageVM.Rotation),
                        ScreenRenderFlags, PixelRegion.Full(w, h), postProcess: true);
                    if (bitmap != null)
                        RunOnUiDispatcherOrDirect(() => pageVM.ImageSource = bitmap);
                }
                catch (Exception ex)
                {
                    Log($"[CRITICAL] RenderPageImage Error: {ex}");
                }
            }
            LoadAnnotationsLazy(model, pageVM);
        }

        // ── 텍스트 ───────────────────────────────────────────────────────

        /// <summary>
        /// 파일에서 검색 (PDF 좌표: left, top, width, height — top은 아래 원점 기준 위쪽 y). 테스트·외부 파일용.
        /// 열린 문서는 <see cref="FindTextRectsInView"/>.
        /// </summary>
        public List<Rect> FindTextRects(string filePath, int pageIndex, string keyword, string? password = null)
        {
            var results = new List<Rect>();
            if (!File.Exists(filePath) || string.IsNullOrEmpty(keyword)) return results;
            using var document = PdfiumDocument.Open(filePath, password);
            lock (Pdfium.Lock)
            {
                foreach (var (l, t, r, b) in document.FindText(pageIndex, keyword))
                    results.Add(new Rect(l, t, r - l, t - b));
            }
            return results;
        }

        /// <summary>열린 문서에서 검색, 결과는 페이지 화면 좌표(DIP). 원본 /Rotate·CropBox·앱 회전 반영.</summary>
        public List<Rect> FindTextRectsInView(PdfDocumentModel model, PdfPageViewModel pageVM, string keyword)
        {
            var results = new List<Rect>();
            var document = model.Document;
            if (document == null || pageVM.IsBlankPage || string.IsNullOrEmpty(keyword)) return results;
            int rotate = ToPdfiumRotate(pageVM.Rotation);
            lock (Pdfium.Lock)
            {
                if (document.IsDisposed) return results;
                foreach (var (l, t, r, b) in document.FindText(pageVM.OriginalPageIndex, keyword))
                    results.Add(document.PageRectToView(pageVM.OriginalPageIndex, pageVM.Width, pageVM.Height, rotate, l, t, r, b));
            }
            return results;
        }

        /// <summary>화면에서 고른 영역(페이지 DIP)의 텍스트.</summary>
        public string ExtractTextInRect(PdfDocumentModel model, PdfPageViewModel pageVM, Rect uiRect)
        {
            var document = model.Document;
            if (document == null || pageVM.IsBlankPage || uiRect.Width <= 0 || uiRect.Height <= 0) return string.Empty;
            int rotate = ToPdfiumRotate(pageVM.Rotation);
            lock (Pdfium.Lock)
            {
                if (document.IsDisposed) return string.Empty;
                int index = pageVM.OriginalPageIndex;
                var a = document.ViewToPage(index, pageVM.Width, pageVM.Height, rotate, uiRect.Left, uiRect.Top);
                var b = document.ViewToPage(index, pageVM.Width, pageVM.Height, rotate, uiRect.Right, uiRect.Bottom);
                return document.GetBoundedText(index,
                    Math.Min(a.X, b.X), Math.Max(a.Y, b.Y), Math.Max(a.X, b.X), Math.Min(a.Y, b.Y));
            }
        }

        // ── 주석 ─────────────────────────────────────────────────────────

        /// <summary>파일의 주석을 화면 주석으로 (한 번만). 렌더 스케줄러가 보이는 페이지에 대해 부른다.</summary>
        internal static void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM, CancellationToken ct = default)
        {
            if (pageVM.AnnotationsLoaded || pageVM.IsBlankPage) return;
            var document = model.Document;
            if (document == null || ct.IsCancellationRequested || model.IsDisposed) return;
            pageVM.AnnotationsLoaded = true;

            try
            {
                var extracted = new List<PdfAnnotation>();
                lock (Pdfium.Lock)
                {
                    if (document.IsDisposed)
                    {
                        pageVM.AnnotationsLoaded = false;
                        return;
                    }

                    bool hasSignatures = document.SignatureCount > 0;
                    int index = pageVM.OriginalPageIndex;
                    int rotate = ToPdfiumRotate(pageVM.Rotation);
                    double viewW = pageVM.Width, viewH = pageVM.Height;
                    // PDF 사용자 공간 → 화면(페이지 DIP): CropBox 오프셋·원본 /Rotate·앱 회전을 PDFium과 같게
                    Rect ToView(double l, double t, double r, double b) => document.PageRectToView(index, viewW, viewH, rotate, l, t, r, b);
                    double dipPerPoint = pageVM.PdfPageWidthPoint > 0 ? pageVM.Width / pageVM.PdfPageWidthPoint : 1.0;

                    document.ReadAnnotations(pageVM.OriginalPageIndex, info =>
                    {
                        if (info.Subtype == PdfiumNative.FPDF_ANNOT_STAMP && PdfiumEditService.IsManagedImageStampContents(info.Contents) &&
                            PdfiumEditService.TryExtractManagedImageStampBytes(info.Handle, info.Contents, out byte[] imageBytes))
                            info.ImageBytes = imageBytes;
                    }).ForEach(info =>
                    {
                        var annotation = ToPdfAnnotation(info, ToView, dipPerPoint, hasSignatures);
                        if (annotation != null)
                            extracted.Add(annotation);
                    });
                }

                // 파일 상태 기록 (인쇄 시 저장 안 한 편집 판단용). 좌표는 페이지 크기 비율이라 줌과 무관.
                pageVM.LoadedAnnotationSignature = PrintService.ComputeAnnotationSignature(extracted, pageVM.Width, pageVM.Height);

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

        /// <param name="dipPerPoint">지금 화면에서 1포인트의 DIP (글자 크기 변환).</param>
        private static PdfAnnotation? ToPdfAnnotation(PdfiumDocument.AnnotationInfo info, Func<double, double, double, double, Rect> toView, double dipPerPoint, bool hasSignatures)
        {
            int subtype = info.Subtype;
            bool isWidget = subtype == PdfiumNative.FPDF_ANNOT_WIDGET;
            if (subtype != PdfiumNative.FPDF_ANNOT_FREETEXT && subtype != PdfiumNative.FPDF_ANNOT_HIGHLIGHT &&
                subtype != PdfiumNative.FPDF_ANNOT_UNDERLINE && subtype != PdfiumNative.FPDF_ANNOT_STAMP && !isWidget)
                return null;
            // 서명이 있는 문서의 위젯만 서명 칸으로 본다 (일반 양식 칸은 제외)
            if (isWidget && (!hasSignatures || string.IsNullOrEmpty(info.FieldName)))
                return null;

            var view = toView(info.Left, info.Top, info.Right, info.Bottom);
            double uiX = view.X;
            double uiY = view.Y;
            double uiW = view.Width;
            double uiH = view.Height;

            string content = info.Contents;
            if (!isWidget && content.StartsWith("Title: ", StringComparison.Ordinal))
                return null;

            double fontSize = 12;
            Brush foreground = Brushes.Black;
            Brush background = Brushes.Transparent;
            AnnotationType type = AnnotationType.FreeText;
            string fontFamily = "Malgun Gothic";
            bool isBold = false;

            if (subtype == PdfiumNative.FPDF_ANNOT_FREETEXT)
            {
                string da = info.DefaultAppearance;
                if (!string.IsNullOrEmpty(da))
                {
                    var fs = Regex.Match(da, @"([\d.]+)\s+Tf");
                    if (fs.Success && double.TryParse(fs.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt))
                        fontSize = pt * dipPerPoint;
                    var rgb = Regex.Match(da, @"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg");
                    if (rgb.Success)
                    {
                        double.TryParse(rgb.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double r);
                        double.TryParse(rgb.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double g);
                        double.TryParse(rgb.Groups[3].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double b);
                        foreground = Frozen(new SolidColorBrush(Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255))));
                    }
                }
            }
            else if (subtype == PdfiumNative.FPDF_ANNOT_STAMP)
            {
                if (PdfiumEditService.TryDecodeManagedFreeText(content, out PdfiumEditService.ManagedFreeTextMetadata managed))
                {
                    content = managed.Text;
                    fontSize = managed.FontSizePt.HasValue ? managed.FontSizePt.Value * dipPerPoint : managed.FontSize;
                    fontFamily = managed.FontFamily;
                    isBold = managed.IsBold;
                    foreground = Frozen(new SolidColorBrush(managed.ForegroundColor));
                }
                else if (info.ImageBytes != null)
                {
                    return new PdfAnnotation
                    {
                        Type = AnnotationType.ImageStamp,
                        X = uiX, Y = uiY, Width = uiW, Height = uiH,
                        ImageBytes = info.ImageBytes,
                        ImageSource = CreateBitmapSource(info.ImageBytes, uiW, uiH)
                    };
                }
                else if (PdfiumEditService.TryDecodeManagedFreeText(content, out string decoded))
                {
                    content = decoded;
                }
                else
                {
                    return null;
                }
            }
            else if (subtype == PdfiumNative.FPDF_ANNOT_HIGHLIGHT)
            {
                type = AnnotationType.Highlight;
                background = info.Color is { } c
                    ? Frozen(new SolidColorBrush(Color.FromArgb(80, c.R, c.G, c.B)))
                    : Frozen(new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)));
            }
            else if (subtype == PdfiumNative.FPDF_ANNOT_UNDERLINE)
            {
                type = AnnotationType.Underline;
            }
            else if (isWidget)
            {
                type = AnnotationType.SignatureField;
            }

            return new PdfAnnotation
            {
                Type = type,
                X = uiX, Y = uiY, Width = uiW, Height = uiH,
                TextContent = content,
                FontSize = fontSize,
                FontFamily = fontFamily,
                IsBold = isBold,
                Foreground = foreground,
                Background = background,
                FieldName = isWidget ? info.FieldName : ""
            };
        }

        private static Brush Frozen(SolidColorBrush brush)
        {
            brush.Freeze();
            return brush;
        }

        // ── 북마크 ───────────────────────────────────────────────────────

        public void LoadBookmarks(PdfDocumentModel model)
        {
            var document = model.Document;
            if (document == null) return;
            List<PdfiumDocument.BookmarkNode> roots;
            lock (Pdfium.Lock)
            {
                if (document.IsDisposed) return;
                roots = document.GetBookmarks();
            }
            var list = roots.Select(n => MapBookmark(n, null)).ToList();
            RunOnUiDispatcherOrDirect(() =>
            {
                model.Bookmarks.Clear();
                foreach (var item in list) model.Bookmarks.Add(item);
            });
        }

        public Task LoadBookmarksAsync(PdfDocumentModel model)
        {
            return Task.Run(() =>
            {
                if (model.IsDisposed)
                    return;
                var watch = Stopwatch.StartNew();
                try
                {
                    LoadBookmarks(model);
                    Log($"[OPEN_PROFILE] LoadBookmarksAsync elapsed={watch.ElapsedMilliseconds}ms path='{model.FilePath}'");
                }
                catch (Exception ex)
                {
                    Log($"[WARN] LoadBookmarksAsync failed: {ex.Message}");
                }
            });
        }

        private static PdfBookmarkViewModel MapBookmark(PdfiumDocument.BookmarkNode node, PdfBookmarkViewModel? parent)
        {
            var vm = new PdfBookmarkViewModel { Title = node.Title, PageIndex = Math.Max(0, node.PageIndex), Parent = parent };
            foreach (var child in node.Children)
                vm.Children.Add(MapBookmark(child, vm));
            return vm;
        }

        // ── 공통 ─────────────────────────────────────────────────────────

        internal static void RunOnUiDispatcherOrDirect(Action action)
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

        private static void Log(string message)
        {
            try
            {
                DebugLog.Append("PdfService", message);
            }
            catch { }
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

        public Task SavePdf(PdfDocumentModel model, string outputPath) => SavePdf(model, outputPath, forPrinting: false);

        /// <param name="forPrinting">
        /// 인쇄용 임시 저장: 화면에 보이는 것(페이지 순서·회전·삭제·빈 페이지·주석)만 반영하고
        /// 보이지 않는 OCR 텍스트 층·북마크 재작성은 건너뛴다 (북마크 변경 표시도 그대로 둔다).
        /// </param>
        public async Task SavePdf(PdfDocumentModel model, string outputPath, bool forPrinting)
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
                shouldRewriteBookmarks = model.HasBookmarkChanges && !forPrinting;
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

                if (await pdfiumEditService.TrySaveAnnotationsAsync(model, outputPath, removeSecurity: forPrinting))
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

                    if (ocrPages.Count > 0 && !forPrinting)
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

    }
}
