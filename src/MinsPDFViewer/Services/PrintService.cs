using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using PdfiumViewer;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;

namespace MinsPDFViewer
{
    /// <summary>
    /// 인쇄할 PDF (화면의 현재 상태). 원본을 그대로 쓰거나, 저장 안 한 편집이 있으면 임시로 저장한 PDF를 메모리에 연다.
    /// 뷰어 문서와 별도의 PDFium 문서라 인쇄 중에도 화면은 그대로다.
    /// </summary>
    public sealed class PrintSource : IDisposable
    {
        internal PrintSource(IPdfDocument document, MemoryStream stream, string documentName, bool fromPendingEdits, bool printRestricted)
        {
            Document = document;
            Stream = stream;
            DocumentName = documentName;
            FromPendingEdits = fromPendingEdits;
            PrintRestricted = printRestricted;
            lock (PdfService.PdfiumLock)
            {
                PageCount = document.PageCount;
                PageSizes = document.PageSizes.Select(s => (Width: (double)s.Width, Height: (double)s.Height)).ToList();
            }
        }

        internal IPdfDocument Document { get; }
        private MemoryStream Stream { get; }
        public string DocumentName { get; }
        public int PageCount { get; }

        /// <summary>페이지 크기(포인트, /Rotate 적용 후 = 화면에 보이는 방향).</summary>
        public IReadOnlyList<(double Width, double Height)> PageSizes { get; }

        /// <summary>저장 안 한 편집(주석·페이지 순서·회전 등)을 반영한 임시 PDF인지.</summary>
        public bool FromPendingEdits { get; }

        /// <summary>문서 권한에 '인쇄 금지'가 있는지. MinsPDFViewer는 무시하고 인쇄한다.</summary>
        public bool PrintRestricted { get; }

        private bool _disposed;

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            lock (PdfService.PdfiumLock)
            {
                try
                {
                    Document.Dispose();
                }
                catch
                {
                }
            }
            Stream.Dispose();
        }
    }

    public sealed class PrintOptions
    {
        public string PrinterName { get; set; } = "";
        /// <summary>null이면 프린터 기본 용지</summary>
        public PaperSize? PaperSize { get; set; }
        public short Copies { get; set; } = 1;
        public bool Collate { get; set; } = true;
        /// <summary>인쇄할 페이지 (0부터, 인쇄 순서대로)</summary>
        public List<int> PageIndices { get; set; } = new();
        public PrintScaleMode ScaleMode { get; set; } = PrintScaleMode.Fit;
        public PrintOrientationMode Orientation { get; set; } = PrintOrientationMode.Auto;
        public bool Grayscale { get; set; }
        public bool IncludeAnnotations { get; set; } = true;
        /// <summary>프린터 속성 창에서 바꾼 드라이버 설정 (DEVMODE). null이면 기본값.</summary>
        public PrinterSettings? PrinterSettings { get; set; }
    }

    public class PrintService
    {
        private readonly PdfService _pdfService;

        public PrintService(PdfService pdfService)
        {
            _pdfService = pdfService;
        }

        private static void Log(string message)
        {
            try
            {
                DebugLog.Append("PrintService", message);
            }
            catch
            {
            }
        }

        // ── 인쇄 대상 준비 ───────────────────────────────────────────────

        /// <summary>
        /// 화면의 현재 상태로 인쇄할 PDF를 준비한다. UI 스레드에서 호출 (페이지·주석 목록을 읽는다).
        /// </summary>
        public async Task<PrintSource> PrepareAsync(PdfDocumentModel model)
        {
            var watch = Stopwatch.StartNew();
            int sourcePageCount;
            lock (PdfService.PdfiumLock)
            {
                sourcePageCount = model.PdfDocument?.PageCount ?? 0;
            }

            bool pendingEdits = HasPendingChanges(model, sourcePageCount, out string reason);
            byte[] originalBytes = model.SourceBytes ?? await File.ReadAllBytesAsync(model.FilePath);
            byte[] bytes;
            if (!pendingEdits)
            {
                bytes = originalBytes;
            }
            else
            {
                Log($"Pending edits ({reason}) → temp save for printing");
                string tempPath = Path.Combine(Path.GetTempPath(), "MinsPrint_" + Guid.NewGuid().ToString("N") + ".pdf");
                try
                {
                    await _pdfService.SavePdf(model, tempPath, forPrinting: true);
                    bytes = await File.ReadAllBytesAsync(tempPath);
                }
                finally
                {
                    try
                    {
                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                    }
                    catch
                    {
                    }
                }
            }

            string? password = model.Password;
            var result = await Task.Run(() =>
            {
                // 권한은 원본 기준 (인쇄용 임시본은 보안을 벗겨 저장하므로)
                uint? permissions = PdfService.GetDocPermissions(originalBytes, password);
                // 암호화되지 않은 문서는 0xFFFFFFFF. bit 3(0x4)이 꺼져 있으면 인쇄 금지 문서.
                bool restricted = permissions.HasValue && (permissions.Value & 0x4) == 0;

                var stream = new MemoryStream(bytes, writable: false);
                IPdfDocument document;
                lock (PdfService.PdfiumLock)
                {
                    document = PdfService.LoadPdfiumDocument(stream, password);
                }
                return new PrintSource(document, stream, model.FileName, pendingEdits, restricted);
            });

            Log($"PrepareAsync pages={result.PageCount} pendingEdits={pendingEdits} restricted={result.PrintRestricted} elapsed={watch.ElapsedMilliseconds}ms");
            return result;
        }

        /// <summary>
        /// 원본 파일과 화면이 다른지: 페이지 구성(순서·삭제·빈 페이지·회전)이나 저장형 주석이 파일에서 읽은 것과 다르면 true.
        /// OCR 단어·북마크는 인쇄 결과에 영향이 없어 보지 않는다.
        /// </summary>
        internal static bool HasPendingChanges(PdfDocumentModel model, int sourcePageCount, out string reason)
        {
            if (model.Pages.Count != sourcePageCount)
            {
                reason = $"page count {model.Pages.Count} != {sourcePageCount}";
                return true;
            }

            for (int i = 0; i < model.Pages.Count; i++)
            {
                var page = model.Pages[i];
                if (page.IsBlankPage)
                {
                    reason = $"blank page at {i + 1}";
                    return true;
                }
                if (page.OriginalPageIndex != i)
                {
                    reason = $"page order changed at {i + 1}";
                    return true;
                }
                if (((page.Rotation % 360) + 360) % 360 != 0)
                {
                    reason = $"page {i + 1} rotated";
                    return true;
                }

                string current = ComputeAnnotationSignature(page.Annotations, page.Width, page.Height);
                string loaded = page.LoadedAnnotationSignature ?? "";
                if (!string.Equals(current, loaded, StringComparison.Ordinal))
                {
                    reason = $"annotations changed on page {i + 1}";
                    return true;
                }
            }

            reason = "";
            return false;
        }

        private static bool IsPersistent(AnnotationType type) =>
            type is AnnotationType.FreeText or AnnotationType.Highlight or AnnotationType.Underline or AnnotationType.ImageStamp;

        /// <summary>
        /// 저장형 주석(텍스트·형광펜·밑줄·이미지)의 내용 서명. 위치·크기는 페이지 크기에 대한 비율(0.1% 단위)이라 줌과 무관하다.
        /// 순서는 무시한다.
        /// </summary>
        public static string ComputeAnnotationSignature(IEnumerable<PdfAnnotation> annotations, double pageWidth, double pageHeight)
        {
            double w = pageWidth > 0 ? pageWidth : 1;
            double h = pageHeight > 0 ? pageHeight : 1;
            var items = new List<string>();
            foreach (var a in annotations)
            {
                if (!IsPersistent(a.Type))
                    continue;

                string fg = (a.Foreground as SolidColorBrush)?.Color.ToString() ?? "";
                string bg = (a.Background as SolidColorBrush)?.Color.ToString() ?? "";
                items.Add(string.Join("|",
                    a.Type.ToString(),
                    Ratio(a.X / w), Ratio(a.Y / h), Ratio(a.Width / w), Ratio(a.Height / h),
                    a.Type is AnnotationType.FreeText ? (a.TextContent ?? "") : "",
                    a.Type is AnnotationType.FreeText ? a.FontSize.ToString("0.0", CultureInfo.InvariantCulture) : "",
                    a.Type is AnnotationType.FreeText ? (a.FontFamily ?? "") : "",
                    a.Type is AnnotationType.FreeText ? (a.IsBold ? "B" : "") : "",
                    a.Type is AnnotationType.FreeText ? fg : "",
                    a.Type is AnnotationType.Highlight ? bg : "",
                    a.Type is AnnotationType.ImageStamp ? (a.ImageBytes?.Length ?? 0).ToString(CultureInfo.InvariantCulture) : ""));
            }

            items.Sort(StringComparer.Ordinal);
            return string.Join("\n", items);
        }

        private static string Ratio(double v) => Math.Round(v, 3).ToString("0.000", CultureInfo.InvariantCulture);

        // ── 미리보기 ─────────────────────────────────────────────────────

        internal static PdfRenderFlags GetRenderFlags(PrintOptions options)
        {
            var flags = PdfRenderFlags.ForPrinting;
            if (options.IncludeAnnotations)
                flags |= PdfRenderFlags.Annotations;
            if (options.Grayscale)
                flags |= PdfRenderFlags.Grayscale;
            return flags;
        }

        /// <summary>미리보기용 페이지 그림 (화면 후처리 없음 = 인쇄 농도 그대로).</summary>
        public static BitmapSource? RenderPreviewPage(PrintSource source, int pageIndex, int pixelWidth, int pixelHeight, bool includeAnnotations, bool grayscale)
        {
            if (pageIndex < 0 || pageIndex >= source.PageCount || pixelWidth <= 0 || pixelHeight <= 0)
                return null;

            var flags = PdfRenderFlags.ForPrinting;
            if (includeAnnotations)
                flags |= PdfRenderFlags.Annotations;
            if (grayscale)
                flags |= PdfRenderFlags.Grayscale;

            lock (PdfService.PdfiumLock)
            {
                using var image = source.Document.Render(pageIndex, pixelWidth, pixelHeight, 96, 96, flags);
                var bitmap = PdfService.ToBitmapSource((DrawingBitmap)image, applyStemDarkening: false, applySharpen: false);
                bitmap.Freeze();
                return bitmap;
            }
        }

        // ── 인쇄 ─────────────────────────────────────────────────────────

        /// <summary>백그라운드에서 인쇄한다. progress = 보낸 페이지 수.</summary>
        public Task PrintAsync(PrintSource source, PrintOptions options, IProgress<int>? progress, CancellationToken ct = default)
        {
            return Task.Run(() =>
            {
                var watch = Stopwatch.StartNew();
                using var document = new MinsPrintDocument(source, options, progress, ct);
                document.Print();
                Log($"Printed {document.PrintedPages} page(s) to '{options.PrinterName}' copies={options.Copies} elapsed={watch.ElapsedMilliseconds}ms restricted={source.PrintRestricted}");
            }, ct);
        }

        /// <summary>
        /// PDFium이 프린터 DC에 직접 그린다 (벡터 그대로, 비트맵 스풀 아님). 페이지마다 PdfiumLock을 잡아 화면 렌더링과 충돌하지 않게.
        /// PDF 권한(인쇄 금지)은 보지 않는다.
        /// </summary>
        private sealed class MinsPrintDocument : PrintDocument
        {
            private readonly PrintSource _source;
            private readonly PrintOptions _options;
            private readonly IProgress<int>? _progress;
            private readonly CancellationToken _ct;
            private int _cursor;

            public int PrintedPages { get; private set; }

            public MinsPrintDocument(PrintSource source, PrintOptions options, IProgress<int>? progress, CancellationToken ct)
            {
                _source = source;
                _options = options;
                _progress = progress;
                _ct = ct;

                DocumentName = string.IsNullOrWhiteSpace(source.DocumentName) ? "MinsPDFViewer" : source.DocumentName;
                PrintController = new StandardPrintController(); // 상태 창 없이 (WinForms 의존 없음)

                PrinterSettings settings = options.PrinterSettings != null
                    ? (PrinterSettings)options.PrinterSettings.Clone()
                    : new PrinterSettings();
                if (!string.IsNullOrWhiteSpace(options.PrinterName))
                    settings.PrinterName = options.PrinterName;
                settings.Copies = Math.Max((short)1, options.Copies);
                settings.Collate = options.Collate;
                settings.PrintRange = PrintRange.AllPages; // 범위는 PageIndices로 직접 고른다
                PrinterSettings = settings;

                if (options.PaperSize != null)
                    DefaultPageSettings.PaperSize = options.PaperSize;
                if (options.Grayscale)
                    DefaultPageSettings.Color = false;
            }

            protected override void OnBeginPrint(PrintEventArgs e)
            {
                _cursor = 0;
                PrintedPages = 0;
                if (_options.PageIndices.Count == 0)
                    e.Cancel = true;
                base.OnBeginPrint(e);
            }

            protected override void OnQueryPageSettings(QueryPageSettingsEventArgs e)
            {
                base.OnQueryPageSettings(e);
                if (_cursor >= _options.PageIndices.Count)
                    return;

                if (_options.PaperSize != null)
                    e.PageSettings.PaperSize = _options.PaperSize;

                var size = _source.PageSizes[_options.PageIndices[_cursor]];
                bool wantLandscape = PrintLayout.IsLandscape(size.Width, size.Height, _options.Orientation, DefaultPageSettings.Landscape);
                // 일부 드라이버는 용지 자체가 가로로 정의돼 Landscape 플래그와 Bounds 방향이 반대다 (PdfiumViewer와 같은 보정)
                bool inverted = false;
                try
                {
                    var bounds = e.PageSettings.Bounds;
                    inverted = (bounds.Width > bounds.Height) != e.PageSettings.Landscape;
                }
                catch
                {
                }
                e.PageSettings.Landscape = inverted ? !wantLandscape : wantLandscape;
                if (_options.Grayscale)
                    e.PageSettings.Color = false;
            }

            protected override void OnPrintPage(PrintPageEventArgs e)
            {
                base.OnPrintPage(e);
                if (_ct.IsCancellationRequested)
                {
                    e.Cancel = true;
                    e.HasMorePages = false;
                    return;
                }

                int pageIndex = _options.PageIndices[_cursor];
                var size = _source.PageSizes[pageIndex];
                var graphics = e.Graphics!;

                // 1/100 인치. 원점 = 용지 왼쪽 위, 인쇄 가능 영역은 물리 여백(HardMargin)부터 VisibleClipBounds 크기.
                double hardX = e.PageSettings.HardMarginX;
                double hardY = e.PageSettings.HardMarginY;
                var clip = graphics.VisibleClipBounds;
                var printable = new PrintRect(hardX, hardY, clip.Width, clip.Height);
                var rect = PrintLayout.Compute(size.Width, size.Height, e.PageBounds.Width, e.PageBounds.Height, printable, _options.ScaleMode);

                // 장치 픽셀 (원점 = 인쇄 가능 영역 왼쪽 위)
                float dpiX = graphics.DpiX;
                float dpiY = graphics.DpiY;
                var bounds = new DrawingRectangle(
                    (int)Math.Round((rect.X - hardX) / 100.0 * dpiX),
                    (int)Math.Round((rect.Y - hardY) / 100.0 * dpiY),
                    Math.Max(1, (int)Math.Round(rect.Width / 100.0 * dpiX)),
                    Math.Max(1, (int)Math.Round(rect.Height / 100.0 * dpiY)));

                lock (PdfService.PdfiumLock)
                {
                    _source.Document.Render(pageIndex, graphics, dpiX, dpiY, bounds, GetRenderFlags(_options));
                }

                _cursor++;
                PrintedPages++;
                _progress?.Report(PrintedPages);
                e.HasMorePages = _cursor < _options.PageIndices.Count;
            }
        }
    }

    /// <summary>프린터 드라이버의 속성 창 (용지·양면·품질 등). 결과는 PrinterSettings의 DEVMODE에 반영.</summary>
    internal static class PrinterPropertiesDialog
    {
        private const int DM_OUT_BUFFER = 2;
        private const int DM_IN_PROMPT = 4;
        private const int DM_IN_BUFFER = 8;
        private const int IDOK = 1;

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter, string pDeviceName, IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GlobalFree(IntPtr hMem);

        /// <returns>확인을 눌렀으면 true (settings·pageSettings 갱신됨)</returns>
        public static bool Show(IntPtr owner, PrinterSettings settings, PageSettings pageSettings)
        {
            IntPtr hDevMode = settings.GetHdevmode(pageSettings);
            IntPtr devModeOut = IntPtr.Zero;
            try
            {
                IntPtr pDevMode = GlobalLock(hDevMode);
                try
                {
                    int size = DocumentProperties(owner, IntPtr.Zero, settings.PrinterName, IntPtr.Zero, IntPtr.Zero, 0);
                    if (size <= 0)
                        return false;
                    devModeOut = Marshal.AllocHGlobal(size);
                    int result = DocumentProperties(owner, IntPtr.Zero, settings.PrinterName, devModeOut, pDevMode, DM_IN_BUFFER | DM_IN_PROMPT | DM_OUT_BUFFER);
                    if (result != IDOK)
                        return false;
                }
                finally
                {
                    GlobalUnlock(hDevMode);
                }

                settings.SetHdevmode(devModeOut);
                pageSettings.SetHdevmode(devModeOut);
                return true;
            }
            finally
            {
                GlobalFree(hDevMode);
                if (devModeOut != IntPtr.Zero)
                    Marshal.FreeHGlobal(devModeOut);
            }
        }
    }
}
