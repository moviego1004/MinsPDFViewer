using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace MinsPDFViewer
{
    /// <summary>
    /// PDFium C API (bblanchon.PDFium.Win32, 버전 고정). 모든 선언은 이 파일 한 곳에 둔다 — 패키지를 올릴 때 헤더와 대조.
    /// PDFium은 스레드 안전하지 않다: 모든 호출은 <see cref="Pdfium.Lock"/> 안에서.
    /// C의 unsigned long은 Windows에서 32비트 → uint (구조체·콜백은 플랫폼 무관하게 CULong).
    /// </summary>
    internal static class PdfiumNative
    {
        private const string Dll = "pdfium.dll";

        // ── 라이브러리 ──
        [StructLayout(LayoutKind.Sequential)]
        public struct FPDF_LIBRARY_CONFIG
        {
            public int version;
            public IntPtr m_pUserFontPaths;
            public IntPtr m_pIsolate;
            public uint m_v8EmbedderSlot;
            public IntPtr m_pPlatform;
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_InitLibraryWithConfig(ref FPDF_LIBRARY_CONFIG config);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint FPDF_GetLastError();

        public const uint FPDF_ERR_SUCCESS = 0;
        public const uint FPDF_ERR_UNKNOWN = 1;
        public const uint FPDF_ERR_FILE = 2;
        public const uint FPDF_ERR_FORMAT = 3;
        public const uint FPDF_ERR_PASSWORD = 4;
        public const uint FPDF_ERR_SECURITY = 5;

        // ── 문서 ──
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int GetBlockCallback(IntPtr param, CULong position, IntPtr buffer, CULong size);

        [StructLayout(LayoutKind.Sequential)]
        public struct FPDF_FILEACCESS
        {
            public CULong m_FileLen;
            public IntPtr m_GetBlock;
            public IntPtr m_Param;
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDF_LoadCustomDocument(IntPtr pFileAccess, [MarshalAs(UnmanagedType.LPUTF8Str)] string? password);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDF_LoadMemDocument64(IntPtr dataBuf, UIntPtr size, [MarshalAs(UnmanagedType.LPUTF8Str)] string? password);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_CloseDocument(IntPtr document);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_GetPageCount(IntPtr document);

        [StructLayout(LayoutKind.Sequential)]
        public struct FS_SIZEF
        {
            public float width;
            public float height;
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_GetPageSizeByIndexF(IntPtr document, int pageIndex, out FS_SIZEF size);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint FPDF_GetDocPermissions(IntPtr document);

        // ── 페이지 ──
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_ClosePage(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern float FPDF_GetPageWidthF(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern float FPDF_GetPageHeightF(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFPage_GetRotation(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_PageToDevice(IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate,
            double pageX, double pageY, out int deviceX, out int deviceY);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_DeviceToPage(IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate,
            int deviceX, int deviceY, out double pageX, out double pageY);

        // ── 렌더링 ──
        public const int FPDF_ANNOT = 0x01;
        public const int FPDF_LCD_TEXT = 0x02;
        public const int FPDF_GRAYSCALE = 0x08;
        public const int FPDF_PRINTING = 0x800;

        public const int FPDFBitmap_BGRA = 4;

        public const int FPDF_RENDER_READY = 0;
        public const int FPDF_RENDER_TOBECONTINUED = 1;
        public const int FPDF_RENDER_DONE = 2;
        public const int FPDF_RENDER_FAILED = 3;

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBitmap_CreateEx(int width, int height, int format, IntPtr firstScan, int stride);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFBitmap_FillRect(IntPtr bitmap, int left, int top, int width, int height, uint color);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFBitmap_Destroy(IntPtr bitmap);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_RenderPageBitmap(IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);

        /// <summary>프린터 등 GDI DC에 직접 (벡터). 좌표는 장치 픽셀.</summary>
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_RenderPage(IntPtr hdc, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        public delegate int NeedToPauseNowCallback(IntPtr pThis);

        [StructLayout(LayoutKind.Sequential)]
        public struct IFSDK_PAUSE
        {
            public int version;
            public IntPtr NeedToPauseNow;
            public IntPtr user;
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_RenderPageBitmap_Start(IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY,
            int rotate, int flags, ref IFSDK_PAUSE pause);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_RenderPage_Continue(IntPtr page, ref IFSDK_PAUSE pause);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_RenderPage_Close(IntPtr page);

        // ── 양식(위젯: 서명 도장·입력 칸) ──
        // FPDF_ANNOT만으로는 위젯이 그려지지 않는다. 양식 환경을 열고 페이지를 그린 뒤 FPDF_FFLDraw로 덧그린다.
        public const int FormFillInfoSize = 512; // FPDF_FORMFILLINFO (version 1, 콜백은 모두 비움 = 읽기 전용)

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFDOC_InitFormFillEnvironment(IntPtr document, IntPtr formInfo);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFDOC_ExitFormFillEnvironment(IntPtr form);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FORM_OnAfterLoadPage(IntPtr page, IntPtr form);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FORM_OnBeforeClosePage(IntPtr page, IntPtr form);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDF_FFLDraw(IntPtr form, IntPtr bitmap, IntPtr page, int startX, int startY, int sizeX, int sizeY, int rotate, int flags);

        // ── 평탄화 (인쇄: 프린터 DC에는 FFLDraw를 못 쓰므로 주석·위젯을 내용으로 굳혀 그린다) ──
        public const int FLAT_PRINT = 1;
        public const int FLATTEN_SUCCESS = 1;

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFPage_Flatten(IntPtr page, int flag);

        // ── 텍스트 ──
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFText_LoadPage(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFText_ClosePage(IntPtr textPage);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFText_FindStart(IntPtr textPage, [MarshalAs(UnmanagedType.LPWStr)] string findWhat, uint flags, int startIndex);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_FindNext(IntPtr handle);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_GetSchResultIndex(IntPtr handle);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_GetSchCount(IntPtr handle);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFText_FindClose(IntPtr handle);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_CountRects(IntPtr textPage, int startIndex, int count);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_GetRect(IntPtr textPage, int rectIndex, out double left, out double top, out double right, out double bottom);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFText_GetBoundedText(IntPtr textPage, double left, double top, double right, double bottom, IntPtr buffer, int bufLen);

        // ── 주석 ──
        public const int FPDF_ANNOT_FREETEXT = 3;
        public const int FPDF_ANNOT_HIGHLIGHT = 9;
        public const int FPDF_ANNOT_UNDERLINE = 10;
        public const int FPDF_ANNOT_WIDGET = 20;
        public const int FPDF_ANNOT_STAMP = 13;
        public const int FPDFANNOT_COLORTYPE_Color = 0;

        [StructLayout(LayoutKind.Sequential)]
        public struct FS_RECTF
        {
            public float left;
            public float top;
            public float right;
            public float bottom;
        }

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFPage_GetAnnotCount(IntPtr page);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFPage_GetAnnot(IntPtr page, int index);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern void FPDFPage_CloseAnnot(IntPtr annot);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFAnnot_GetSubtype(IntPtr annot);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFAnnot_GetRect(IntPtr annot, out FS_RECTF rect);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint FPDFAnnot_GetStringValue(IntPtr annot, [MarshalAs(UnmanagedType.LPStr)] string key, IntPtr buffer, uint bufLen);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFAnnot_GetColor(IntPtr annot, int colorType, out uint r, out uint g, out uint b, out uint a);

        // ── 서명 ──
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDF_GetSignatureCount(IntPtr document);

        // ── 북마크 ──
        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBookmark_GetFirstChild(IntPtr document, IntPtr bookmark);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBookmark_GetNextSibling(IntPtr document, IntPtr bookmark);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern uint FPDFBookmark_GetTitle(IntPtr bookmark, IntPtr buffer, uint bufLen);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBookmark_GetDest(IntPtr document, IntPtr bookmark);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFBookmark_GetAction(IntPtr bookmark);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr FPDFAction_GetDest(IntPtr document, IntPtr action);

        [DllImport(Dll, CallingConvention = CallingConvention.Cdecl)]
        public static extern int FPDFDest_GetDestPageIndex(IntPtr document, IntPtr dest);
    }

    /// <summary>PDFium 초기화와 전역 잠금.</summary>
    internal static class Pdfium
    {
        /// <summary>모든 PDFium 호출을 직렬화하는 전역 잠금 (PdfService.PdfiumLock과 같은 객체).</summary>
        public static readonly object Lock = new();

        private static int _initialized;

        public static void EnsureInitialized()
        {
            if (Volatile.Read(ref _initialized) == 1)
                return;

            lock (Lock)
            {
                if (_initialized == 1)
                    return;
                var config = new PdfiumNative.FPDF_LIBRARY_CONFIG { version = 2 };
                PdfiumNative.FPDF_InitLibraryWithConfig(ref config);
                Volatile.Write(ref _initialized, 1);
            }
        }

        public static string DescribeError(uint code) => code switch
        {
            PdfiumNative.FPDF_ERR_FILE => "파일을 열 수 없음",
            PdfiumNative.FPDF_ERR_FORMAT => "PDF 형식 오류",
            PdfiumNative.FPDF_ERR_PASSWORD => "암호 필요",
            PdfiumNative.FPDF_ERR_SECURITY => "지원하지 않는 보안 방식",
            _ => $"알 수 없는 오류 ({code})"
        };
    }

    internal sealed class PdfiumException : Exception
    {
        public uint ErrorCode { get; }
        public bool IsPasswordError => ErrorCode == PdfiumNative.FPDF_ERR_PASSWORD;

        public PdfiumException(uint errorCode, string message) : base($"{message}: {Pdfium.DescribeError(errorCode)}")
        {
            ErrorCode = errorCode;
        }
    }
}
