using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace MinsPDFViewer
{
    /// <summary>
    /// PDFium 문서 하나 = 핸들 하나. 렌더링·텍스트·주석·북마크·권한이 모두 이 핸들을 쓴다.
    /// 최근 페이지(FPDF_PAGE + 텍스트 페이지)를 열어 둔 채 재사용해 디코딩된 이미지도 재사용된다.
    /// 모든 멤버는 <see cref="Pdfium.Lock"/> 안에서 호출해야 한다 (Dispose 포함, 내부에서 잡음).
    /// </summary>
    public sealed class PdfiumDocument : IDisposable
    {
        private const int PageCacheCapacity = 12;

        private IntPtr _document;
        private FileStream? _stream;
        private IntPtr _fileAccess;
        private PdfiumNative.GetBlockCallback? _getBlock; // 네이티브가 들고 있는 동안 GC되지 않게
        private GCHandle _pinnedBytes;

        private readonly Dictionary<int, CachedPage> _pages = new();
        private readonly LinkedList<int> _pageOrder = new();

        public string? FilePath { get; }
        public int PageCount { get; }
        public bool IsDisposed => _document == IntPtr.Zero;

        private sealed class CachedPage
        {
            public IntPtr Page;
            public IntPtr TextPage;
        }

        private IntPtr _form;
        private IntPtr _formInfo;
        private readonly bool _flattenForPrint;

        private PdfiumDocument(IntPtr document, string? filePath, bool flattenForPrint)
        {
            _document = document;
            FilePath = filePath;
            PageCount = PdfiumNative.FPDF_GetPageCount(document);
            _flattenForPrint = flattenForPrint;
            if (!flattenForPrint)
                InitForms();
        }

        /// <summary>읽기 전용 양식 환경 (콜백 없음). 위젯(서명 도장·입력 칸)을 그리는 데만 쓴다.</summary>
        private void InitForms()
        {
            _formInfo = Marshal.AllocHGlobal(PdfiumNative.FormFillInfoSize);
            Marshal.Copy(new byte[PdfiumNative.FormFillInfoSize], 0, _formInfo, PdfiumNative.FormFillInfoSize);
            Marshal.WriteInt32(_formInfo, 0, 1); // version 1
            _form = PdfiumNative.FPDFDOC_InitFormFillEnvironment(_document, _formInfo);
            if (_form == IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_formInfo);
                _formInfo = IntPtr.Zero;
            }
        }

        // ── 열기 ─────────────────────────────────────────────────────────

        /// <summary>파일을 스트림으로 연다 (통째로 메모리에 올리지 않음). 실패하면 <see cref="PdfiumException"/>.</summary>
        /// <param name="flattenForPrint">페이지를 열 때 주석·위젯을 내용으로 굳힌다 (인쇄용 별도 문서만, 저장하지 않음).</param>
        public static PdfiumDocument Open(string path, string? password, bool flattenForPrint = false)
        {
            Pdfium.EnsureInitialized();
            var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1, FileOptions.RandomAccess);
            if (stream.Length > uint.MaxValue)
            {
                stream.Dispose();
                throw new PdfiumException(PdfiumNative.FPDF_ERR_FILE, "4GB를 넘는 파일은 열 수 없습니다");
            }

            PdfiumNative.GetBlockCallback? getBlock = null;
            IntPtr access = IntPtr.Zero;
            try
            {
                lock (Pdfium.Lock)
                {
                    byte[] readBuffer = new byte[64 * 1024];
                    var source = stream; // 아래에서 stream 변수를 비우므로 콜백은 따로 잡아 둔다
                    getBlock = (param, position, buffer, size) => ReadBlock(source, readBuffer, (long)position.Value, buffer, (long)size.Value);
                    access = Marshal.AllocHGlobal(Marshal.SizeOf<PdfiumNative.FPDF_FILEACCESS>());
                    Marshal.StructureToPtr(new PdfiumNative.FPDF_FILEACCESS
                    {
                        m_FileLen = new CULong((uint)stream.Length),
                        m_GetBlock = Marshal.GetFunctionPointerForDelegate(getBlock),
                        m_Param = IntPtr.Zero
                    }, access, false);

                    IntPtr doc = PdfiumNative.FPDF_LoadCustomDocument(access, password);
                    if (doc == IntPtr.Zero)
                        throw new PdfiumException(PdfiumNative.FPDF_GetLastError(), "PDF를 열 수 없습니다");

                    var result = new PdfiumDocument(doc, path, flattenForPrint)
                    {
                        _stream = stream,
                        _fileAccess = access,
                        _getBlock = getBlock
                    };
                    access = IntPtr.Zero;
                    stream = null!;
                    return result;
                }
            }
            finally
            {
                if (access != IntPtr.Zero)
                    Marshal.FreeHGlobal(access);
                stream?.Dispose();
                GC.KeepAlive(getBlock);
            }
        }

        /// <summary>메모리의 PDF를 연다 (배열은 문서를 닫을 때까지 고정됨).</summary>
        public static PdfiumDocument OpenBytes(byte[] bytes, string? password, bool flattenForPrint = false)
        {
            Pdfium.EnsureInitialized();
            var handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                lock (Pdfium.Lock)
                {
                    IntPtr doc = PdfiumNative.FPDF_LoadMemDocument64(handle.AddrOfPinnedObject(), (UIntPtr)bytes.Length, password);
                    if (doc == IntPtr.Zero)
                        throw new PdfiumException(PdfiumNative.FPDF_GetLastError(), "PDF를 열 수 없습니다");
                    var result = new PdfiumDocument(doc, null, flattenForPrint) { _pinnedBytes = handle };
                    handle = default;
                    return result;
                }
            }
            finally
            {
                if (handle.IsAllocated)
                    handle.Free();
            }
        }

        private static int ReadBlock(FileStream stream, byte[] readBuffer, long position, IntPtr buffer, long size)
        {
            try
            {
                stream.Position = position;
                long done = 0;
                while (done < size)
                {
                    int chunk = (int)Math.Min(readBuffer.Length, size - done);
                    int read = stream.Read(readBuffer, 0, chunk);
                    if (read <= 0)
                        return 0;
                    Marshal.Copy(readBuffer, 0, IntPtr.Add(buffer, (int)done), read);
                    done += read;
                }
                return 1;
            }
            catch
            {
                return 0;
            }
        }

        private void ThrowIfDisposed()
        {
            if (_document == IntPtr.Zero)
                throw new ObjectDisposedException(nameof(PdfiumDocument));
        }

        // ── 문서 정보 ────────────────────────────────────────────────────

        /// <summary>페이지 크기(포인트). CropBox 기준, 원본 /Rotate 적용 후 = 화면에 보이는 방향.</summary>
        public (double Width, double Height) GetPageSize(int pageIndex)
        {
            ThrowIfDisposed();
            if (PdfiumNative.FPDF_GetPageSizeByIndexF(_document, pageIndex, out var size) == 0)
                return (612, 792);
            return (size.width, size.height);
        }

        /// <summary>PDF 권한 비트. 암호화되지 않은 문서는 0xFFFFFFFF. bit 3(0x4) = 인쇄.</summary>
        public uint Permissions
        {
            get
            {
                ThrowIfDisposed();
                return PdfiumNative.FPDF_GetDocPermissions(_document);
            }
        }

        public int SignatureCount
        {
            get
            {
                ThrowIfDisposed();
                return PdfiumNative.FPDF_GetSignatureCount(_document);
            }
        }

        // ── 페이지 캐시 ──────────────────────────────────────────────────

        internal IntPtr GetPage(int pageIndex)
        {
            ThrowIfDisposed();
            if (pageIndex < 0 || pageIndex >= PageCount)
                return IntPtr.Zero;

            if (_pages.TryGetValue(pageIndex, out var cached))
            {
                _pageOrder.Remove(pageIndex);
                _pageOrder.AddFirst(pageIndex);
                return cached.Page;
            }

            IntPtr page = PdfiumNative.FPDF_LoadPage(_document, pageIndex);
            if (page == IntPtr.Zero)
                return IntPtr.Zero;

            if (_flattenForPrint && PdfiumNative.FPDFPage_Flatten(page, PdfiumNative.FLAT_PRINT) == PdfiumNative.FLATTEN_SUCCESS)
            {
                // 평탄화한 내용은 페이지를 다시 열어야 반영된다 (메모리에서만, 파일은 그대로)
                PdfiumNative.FPDF_ClosePage(page);
                page = PdfiumNative.FPDF_LoadPage(_document, pageIndex);
                if (page == IntPtr.Zero)
                    return IntPtr.Zero;
            }
            if (_form != IntPtr.Zero)
                PdfiumNative.FORM_OnAfterLoadPage(page, _form);

            _pages[pageIndex] = new CachedPage { Page = page };
            _pageOrder.AddFirst(pageIndex);
            while (_pageOrder.Count > PageCacheCapacity)
            {
                int evict = _pageOrder.Last!.Value;
                _pageOrder.RemoveLast();
                if (_pages.Remove(evict, out var old))
                    ClosePage(old);
            }
            return page;
        }

        internal IntPtr GetTextPage(int pageIndex)
        {
            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return IntPtr.Zero;
            var cached = _pages[pageIndex];
            if (cached.TextPage == IntPtr.Zero)
                cached.TextPage = PdfiumNative.FPDFText_LoadPage(page);
            return cached.TextPage;
        }

        private void ClosePage(CachedPage cached)
        {
            if (cached.TextPage != IntPtr.Zero)
                PdfiumNative.FPDFText_ClosePage(cached.TextPage);
            if (cached.Page != IntPtr.Zero)
            {
                if (_form != IntPtr.Zero)
                    PdfiumNative.FORM_OnBeforeClosePage(cached.Page, _form);
                PdfiumNative.FPDF_ClosePage(cached.Page);
            }
            cached.TextPage = IntPtr.Zero;
            cached.Page = IntPtr.Zero;
        }

        // ── 렌더링 ───────────────────────────────────────────────────────

        /// <summary>
        /// 페이지를 fullWidth×fullHeight 크기로 그렸을 때의 region 부분만 BGRA 버퍼로 돌려준다 (타일).
        /// 점진 렌더: shouldCancel이 true를 돌려주면 중간에 멈추고 null.
        /// rotate: 0~3 (시계 방향 90° 단위, 원본 /Rotate에 더해짐).
        /// </summary>
        public byte[]? RenderToBuffer(int pageIndex, int fullWidth, int fullHeight, int rotate, int flags,
            PixelRegion region, Func<bool>? shouldCancel = null)
        {
            ThrowIfDisposed();
            if (region.Width <= 0 || region.Height <= 0 || fullWidth <= 0 || fullHeight <= 0)
                return null;

            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return null;

            int stride = region.Width * 4;
            byte[] buffer = new byte[(long)stride * region.Height];
            var handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
            IntPtr bitmap = IntPtr.Zero;
            try
            {
                bitmap = PdfiumNative.FPDFBitmap_CreateEx(region.Width, region.Height, PdfiumNative.FPDFBitmap_BGRA, handle.AddrOfPinnedObject(), stride);
                if (bitmap == IntPtr.Zero)
                    return null;
                PdfiumNative.FPDFBitmap_FillRect(bitmap, 0, 0, region.Width, region.Height, 0xFFFFFFFF);

                if (shouldCancel == null)
                {
                    PdfiumNative.FPDF_RenderPageBitmap(bitmap, page, -region.X, -region.Y, fullWidth, fullHeight, rotate & 3, flags);
                    DrawForms(bitmap, page, region, fullWidth, fullHeight, rotate, flags);
                    return buffer;
                }

                PdfiumNative.NeedToPauseNowCallback pauseCallback = _ => shouldCancel() ? 1 : 0;
                var pause = new PdfiumNative.IFSDK_PAUSE
                {
                    version = 1,
                    NeedToPauseNow = Marshal.GetFunctionPointerForDelegate(pauseCallback),
                    user = IntPtr.Zero
                };
                try
                {
                    int status = PdfiumNative.FPDF_RenderPageBitmap_Start(bitmap, page, -region.X, -region.Y, fullWidth, fullHeight, rotate & 3, flags, ref pause);
                    while (status == PdfiumNative.FPDF_RENDER_TOBECONTINUED)
                    {
                        if (shouldCancel())
                            return null;
                        status = PdfiumNative.FPDF_RenderPage_Continue(page, ref pause);
                    }
                    if (status != PdfiumNative.FPDF_RENDER_DONE)
                        return null;
                    DrawForms(bitmap, page, region, fullWidth, fullHeight, rotate, flags);
                    return buffer;
                }
                finally
                {
                    PdfiumNative.FPDF_RenderPage_Close(page);
                    GC.KeepAlive(pauseCallback);
                }
            }
            finally
            {
                if (bitmap != IntPtr.Zero)
                    PdfiumNative.FPDFBitmap_Destroy(bitmap);
                handle.Free();
            }
        }

        private void DrawForms(IntPtr bitmap, IntPtr page, PixelRegion region, int fullWidth, int fullHeight, int rotate, int flags)
        {
            if (_form != IntPtr.Zero && (flags & PdfiumNative.FPDF_ANNOT) != 0)
                PdfiumNative.FPDF_FFLDraw(_form, bitmap, page, -region.X, -region.Y, fullWidth, fullHeight, rotate & 3, flags);
        }

        /// <summary>GDI DC(프린터)에 직접 그린다. 좌표는 장치 픽셀 (DC 원점 기준).</summary>
        public void RenderToDC(int pageIndex, IntPtr hdc, int x, int y, int width, int height, int rotate, int flags)
        {
            ThrowIfDisposed();
            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return;
            PdfiumNative.FPDF_RenderPage(hdc, page, x, y, width, height, rotate & 3, flags);
        }

        // ── 좌표 변환 ────────────────────────────────────────────────────

        private const int CoordinatePrecision = 16; // FPDF_PageToDevice는 정수 → 16배로 계산해 1/16 픽셀 정밀도

        /// <summary>PDF 사용자 공간 좌표 → 화면 좌표 (width×height 크기, rotate 적용). CropBox·원본 /Rotate 반영.</summary>
        public (double X, double Y) PageToView(int pageIndex, double width, double height, int rotate, double pageX, double pageY)
        {
            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return (0, 0);
            int w = Math.Max(1, (int)Math.Round(width * CoordinatePrecision));
            int h = Math.Max(1, (int)Math.Round(height * CoordinatePrecision));
            PdfiumNative.FPDF_PageToDevice(page, 0, 0, w, h, rotate & 3, pageX, pageY, out int dx, out int dy);
            return (dx / (double)CoordinatePrecision, dy / (double)CoordinatePrecision);
        }

        /// <summary>화면 좌표 → PDF 사용자 공간 좌표.</summary>
        public (double X, double Y) ViewToPage(int pageIndex, double width, double height, int rotate, double viewX, double viewY)
        {
            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return (0, 0);
            int w = Math.Max(1, (int)Math.Round(width * CoordinatePrecision));
            int h = Math.Max(1, (int)Math.Round(height * CoordinatePrecision));
            PdfiumNative.FPDF_DeviceToPage(page, 0, 0, w, h, rotate & 3,
                (int)Math.Round(viewX * CoordinatePrecision), (int)Math.Round(viewY * CoordinatePrecision), out double px, out double py);
            return (px, py);
        }

        /// <summary>PDF 좌표 사각형(left, top, right, bottom) → 화면 사각형.</summary>
        public System.Windows.Rect PageRectToView(int pageIndex, double width, double height, int rotate, double left, double top, double right, double bottom)
        {
            var a = PageToView(pageIndex, width, height, rotate, left, top);
            var b = PageToView(pageIndex, width, height, rotate, right, bottom);
            return new System.Windows.Rect(new System.Windows.Point(a.X, a.Y), new System.Windows.Point(b.X, b.Y));
        }

        // ── 텍스트 ───────────────────────────────────────────────────────

        /// <summary>검색어가 나오는 곳마다 PDF 좌표 사각형들 (한 결과가 여러 줄이면 여러 개).</summary>
        public List<(double Left, double Top, double Right, double Bottom)> FindText(int pageIndex, string query)
        {
            var results = new List<(double, double, double, double)>();
            if (string.IsNullOrEmpty(query))
                return results;
            IntPtr textPage = GetTextPage(pageIndex);
            if (textPage == IntPtr.Zero)
                return results;

            IntPtr search = PdfiumNative.FPDFText_FindStart(textPage, query, 0, 0);
            if (search == IntPtr.Zero)
                return results;
            try
            {
                while (PdfiumNative.FPDFText_FindNext(search) != 0)
                {
                    int start = PdfiumNative.FPDFText_GetSchResultIndex(search);
                    int count = PdfiumNative.FPDFText_GetSchCount(search);
                    int rects = PdfiumNative.FPDFText_CountRects(textPage, start, count);
                    for (int i = 0; i < rects; i++)
                    {
                        if (PdfiumNative.FPDFText_GetRect(textPage, i, out double l, out double t, out double r, out double b) != 0)
                            results.Add((l, t, r, b));
                    }
                }
            }
            finally
            {
                PdfiumNative.FPDFText_FindClose(search);
            }
            return results;
        }

        /// <summary>PDF 좌표 사각형 안의 텍스트.</summary>
        public string GetBoundedText(int pageIndex, double left, double top, double right, double bottom)
        {
            IntPtr textPage = GetTextPage(pageIndex);
            if (textPage == IntPtr.Zero)
                return string.Empty;
            int length = PdfiumNative.FPDFText_GetBoundedText(textPage, left, top, right, bottom, IntPtr.Zero, 0);
            if (length <= 0)
                return string.Empty;
            IntPtr buffer = Marshal.AllocHGlobal((length + 1) * 2);
            try
            {
                int written = PdfiumNative.FPDFText_GetBoundedText(textPage, left, top, right, bottom, buffer, length + 1);
                return written <= 0 ? string.Empty : (Marshal.PtrToStringUni(buffer, written) ?? string.Empty).TrimEnd('\0').Trim();
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        // ── 주석 ─────────────────────────────────────────────────────────

        public sealed class AnnotationInfo
        {
            public int Subtype { get; init; }
            /// <summary>PDF 좌표</summary>
            public float Left { get; init; }
            public float Top { get; init; }
            public float Right { get; init; }
            public float Bottom { get; init; }
            public string Contents { get; init; } = "";
            public string DefaultAppearance { get; init; } = "";
            public string FieldName { get; init; } = "";
            public (byte R, byte G, byte B, byte A)? Color { get; init; }
            /// <summary>관리형 이미지 도장의 원본 바이트 (있으면)</summary>
            public byte[]? ImageBytes { get; set; }
            internal IntPtr Handle { get; init; }
        }

        /// <summary>
        /// 페이지 주석을 읽는다. visitor는 주석 핸들이 열려 있는 동안 호출된다 (이미지 추출 등).
        /// </summary>
        public List<AnnotationInfo> ReadAnnotations(int pageIndex, Action<AnnotationInfo>? visitor = null)
        {
            var list = new List<AnnotationInfo>();
            IntPtr page = GetPage(pageIndex);
            if (page == IntPtr.Zero)
                return list;

            int count = PdfiumNative.FPDFPage_GetAnnotCount(page);
            for (int i = 0; i < count; i++)
            {
                IntPtr annot = PdfiumNative.FPDFPage_GetAnnot(page, i);
                if (annot == IntPtr.Zero)
                    continue;
                try
                {
                    if (PdfiumNative.FPDFAnnot_GetRect(annot, out var rect) == 0)
                        continue;
                    (byte, byte, byte, byte)? color = null;
                    if (PdfiumNative.FPDFAnnot_GetColor(annot, PdfiumNative.FPDFANNOT_COLORTYPE_Color, out uint r, out uint g, out uint b, out uint a) != 0)
                        color = ((byte)r, (byte)g, (byte)b, (byte)a);

                    var info = new AnnotationInfo
                    {
                        Subtype = PdfiumNative.FPDFAnnot_GetSubtype(annot),
                        Left = rect.left,
                        Top = rect.top,
                        Right = rect.right,
                        Bottom = rect.bottom,
                        Contents = GetAnnotationString(annot, "Contents"),
                        DefaultAppearance = GetAnnotationString(annot, "DA"),
                        FieldName = GetAnnotationString(annot, "T"),
                        Color = color,
                        Handle = annot
                    };
                    visitor?.Invoke(info);
                    list.Add(info);
                }
                finally
                {
                    PdfiumNative.FPDFPage_CloseAnnot(annot);
                }
            }
            return list;
        }

        internal static string GetAnnotationString(IntPtr annot, string key)
        {
            uint length = PdfiumNative.FPDFAnnot_GetStringValue(annot, key, IntPtr.Zero, 0);
            if (length <= 2)
                return string.Empty;
            IntPtr buffer = Marshal.AllocHGlobal((int)length);
            try
            {
                PdfiumNative.FPDFAnnot_GetStringValue(annot, key, buffer, length);
                return Marshal.PtrToStringUni(buffer) ?? string.Empty;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        // ── 북마크 ───────────────────────────────────────────────────────

        public sealed class BookmarkNode
        {
            public string Title { get; init; } = "";
            public int PageIndex { get; init; }
            public List<BookmarkNode> Children { get; } = new();
        }

        public List<BookmarkNode> GetBookmarks()
        {
            ThrowIfDisposed();
            var roots = new List<BookmarkNode>();
            var visited = new HashSet<IntPtr>();
            ReadBookmarkChildren(IntPtr.Zero, roots, visited, 0);
            return roots;
        }

        private void ReadBookmarkChildren(IntPtr parent, List<BookmarkNode> target, HashSet<IntPtr> visited, int depth)
        {
            if (depth > 64)
                return;
            IntPtr child = PdfiumNative.FPDFBookmark_GetFirstChild(_document, parent);
            while (child != IntPtr.Zero && visited.Add(child))
            {
                var node = new BookmarkNode { Title = GetBookmarkTitle(child), PageIndex = GetBookmarkPage(child) };
                ReadBookmarkChildren(child, node.Children, visited, depth + 1);
                target.Add(node);
                child = PdfiumNative.FPDFBookmark_GetNextSibling(_document, child);
            }
        }

        private static string GetBookmarkTitle(IntPtr bookmark)
        {
            uint length = PdfiumNative.FPDFBookmark_GetTitle(bookmark, IntPtr.Zero, 0);
            if (length <= 2)
                return string.Empty;
            IntPtr buffer = Marshal.AllocHGlobal((int)length);
            try
            {
                PdfiumNative.FPDFBookmark_GetTitle(bookmark, buffer, length);
                return Marshal.PtrToStringUni(buffer) ?? string.Empty;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private int GetBookmarkPage(IntPtr bookmark)
        {
            IntPtr dest = PdfiumNative.FPDFBookmark_GetDest(_document, bookmark);
            if (dest == IntPtr.Zero)
            {
                IntPtr action = PdfiumNative.FPDFBookmark_GetAction(bookmark);
                if (action != IntPtr.Zero)
                    dest = PdfiumNative.FPDFAction_GetDest(_document, action);
            }
            if (dest == IntPtr.Zero)
                return -1;
            return PdfiumNative.FPDFDest_GetDestPageIndex(_document, dest);
        }

        // ── 닫기 ─────────────────────────────────────────────────────────

        public void Dispose()
        {
            lock (Pdfium.Lock)
            {
                if (_document == IntPtr.Zero)
                    return;
                foreach (var cached in _pages.Values)
                    ClosePage(cached);
                _pages.Clear();
                _pageOrder.Clear();
                if (_form != IntPtr.Zero)
                {
                    PdfiumNative.FPDFDOC_ExitFormFillEnvironment(_form);
                    _form = IntPtr.Zero;
                }
                PdfiumNative.FPDF_CloseDocument(_document);
                _document = IntPtr.Zero;
            }

            if (_formInfo != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_formInfo);
                _formInfo = IntPtr.Zero;
            }

            if (_fileAccess != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(_fileAccess);
                _fileAccess = IntPtr.Zero;
            }
            _stream?.Dispose();
            _stream = null;
            _getBlock = null;
            if (_pinnedBytes.IsAllocated)
                _pinnedBytes.Free();
        }
    }

    /// <summary>렌더 결과 중 잘라 낼 영역 (페이지 전체 크기 기준 픽셀).</summary>
    public readonly record struct PixelRegion(int X, int Y, int Width, int Height)
    {
        public static PixelRegion Full(int width, int height) => new(0, 0, width, height);
        public bool Contains(PixelRegion other) =>
            other.X >= X && other.Y >= Y && other.X + other.Width <= X + Width && other.Y + other.Height <= Y + Height;
    }
}
