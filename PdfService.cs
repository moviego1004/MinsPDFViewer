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

            // Annotation subtypes
            public const int FPDF_ANNOT_UNKNOWN = 0;
            public const int FPDF_ANNOT_TEXT = 1;
            public const int FPDF_ANNOT_LINK = 2;
            public const int FPDF_ANNOT_FREETEXT = 3;
            public const int FPDF_ANNOT_LINE = 4;
            public const int FPDF_ANNOT_SQUARE = 5;
            public const int FPDF_ANNOT_CIRCLE = 6;
            public const int FPDF_ANNOT_POLYGON = 7;
            public const int FPDF_ANNOT_POLYLINE = 8;
            public const int FPDF_ANNOT_HIGHLIGHT = 9;
            public const int FPDF_ANNOT_UNDERLINE = 10;
            public const int FPDF_ANNOT_SQUIGGLY = 11;
            public const int FPDF_ANNOT_STRIKEOUT = 12;
            public const int FPDF_ANNOT_STAMP = 13;
            public const int FPDF_ANNOT_CARET = 14;
            public const int FPDF_ANNOT_INK = 15;
            public const int FPDF_ANNOT_POPUP = 16;
            public const int FPDF_ANNOT_FILEATTACHMENT = 17;
            public const int FPDF_ANNOT_SOUND = 18;
            public const int FPDF_ANNOT_MOVIE = 19;
            public const int FPDF_ANNOT_WIDGET = 20;

            // Color types for FPDFAnnot_GetColor
            public const int FPDFANNOT_COLORTYPE_Color = 0;
            public const int FPDFANNOT_COLORTYPE_InteriorColor = 1;
        }

        // 1. PDF 로드
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

        // 2. 문서 초기화
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

        // 3. 렌더링
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

        // 4. 주석 로드 (Pdfium P/Invoke 사용)
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

                    // Pdfium을 사용하여 주석 로드
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
                                        // Rect 가져오기
                                        var rect = new NativeMethods.FS_RECTF();
                                        if (!NativeMethods.FPDFAnnot_GetRect(annot, ref rect)) continue;

                                        // PDF 좌표계(좌하단 0,0)를 UI 좌표계(좌상단 0,0)로 변환
                                        double uiX = rect.left * (pageVM.Width / pageW);
                                        double uiY = (pageH - rect.top) * (pageVM.Height / pageH);
                                        double uiW = (rect.right - rect.left) * (pageVM.Width / pageW);
                                        double uiH = (rect.top - rect.bottom) * (pageVM.Height / pageH);

                                        // Contents 읽기
                                        string content = GetAnnotationStringValue(annot, "Contents");

                                        // 스타일 파싱
                                        double fSize = 12;
                                        Brush brush = Brushes.Black;

                                        // /DA 문자열 추출 시도
                                        string da = GetAnnotationStringValue(annot, "DA");
                                        if (!string.IsNullOrEmpty(da))
                                        {
                                            // FontSize 파싱
                                            var fsMatch = Regex.Match(da, @"([\d.]+)\s+Tf");
                                            if (fsMatch.Success) double.TryParse(fsMatch.Groups[1].Value, out fSize);

                                            // Color 파싱 (RGB)
                                            var colMatch = Regex.Match(da, @"([\d.]+)\s+([\d.]+)\s+([\d.]+)\s+rg");
                                            if (colMatch.Success)
                                            {
                                                double.TryParse(colMatch.Groups[1].Value, out double r);
                                                double.TryParse(colMatch.Groups[2].Value, out double g);
                                                double.TryParse(colMatch.Groups[3].Value, out double b);
                                                brush = new SolidColorBrush(Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255)));
                                            }
                                        }

                                        // Highlight 배경색 파싱
                                        Brush background = Brushes.Transparent;
                                        if (subtype == NativeMethods.FPDF_ANNOT_HIGHLIGHT)
                                        {
                                            background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)); // 기본 노란색
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
                    // 읽기 실패 시 조용히 무시하거나 로그
                    System.Diagnostics.Debug.WriteLine($"Annotation Load Error: {ex.Message}");
                }
            });
        }

        // Pdfium에서 주석 문자열 값 가져오기 헬퍼
        private string GetAnnotationStringValue(IntPtr annot, string key)
        {
            // 먼저 필요한 버퍼 크기 확인
            ulong len = NativeMethods.FPDFAnnot_GetStringValue(annot, key, IntPtr.Zero, 0);
            if (len == 0) return "";

            // UTF-16LE로 반환되므로 버퍼 할당
            IntPtr buffer = System.Runtime.InteropServices.Marshal.AllocHGlobal((int)len);
            try
            {
                NativeMethods.FPDFAnnot_GetStringValue(annot, key, buffer, len);
                // UTF-16LE 문자열로 변환 (null terminator 제외)
                return System.Runtime.InteropServices.Marshal.PtrToStringUni(buffer) ?? "";
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FreeHGlobal(buffer);
            }
        }

        // 5. 저장 (자체 구현 엔진 - PdfSharp 미사용)
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null) return; // model null 체크

            // 1. 스냅샷 생성
            List<PageSaveData> pagesSnapshot = new List<PageSaveData>();
            string originalFilePath = model.FilePath;
            Exception? collectException = null;

            // UI 스레드에서 데이터 수집 (동기적으로 실행하여 확실히 완료되도록 함)
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine($"SavePdf: Starting data collection. Pages count = {model.Pages.Count}");

                    if (model.Pages.Count == 0)
                    {
                        // 페이지가 없으면 저장할 것도 없음
                        System.Diagnostics.Debug.WriteLine("SavePdf: No pages in document model.");
                        return;
                    }

                    foreach (var p in model.Pages)
                    {
                        var pageData = new PageSaveData
                        {
                            OriginalPageIndex = p.OriginalPageIndex,
                            Width = p.Width,
                            Height = p.Height,
                            PdfPageWidthPoint = p.PdfPageWidthPoint,
                            PdfPageHeightPoint = p.PdfPageHeightPoint,
                            Annotations = new List<AnnotationSaveData>()
                        };

                        // 주석이 없는 페이지도 스냅샷에는 포함시켜야 함 (페이지 구조 유지를 위해)
                        foreach (var a in p.Annotations)
                        {
                            if (a.Type == AnnotationType.SearchHighlight) continue;
                            pageData.Annotations.Add(new AnnotationSaveData
                            {
                                Type = a.Type,
                                X = a.X,
                                Y = a.Y,
                                Width = a.Width,
                                Height = a.Height,
                                TextContent = a.TextContent ?? "",
                                FontSize = a.FontSize,
                                ForeR = (a.Foreground as SolidColorBrush)?.Color.R ?? 0,
                                ForeG = (a.Foreground as SolidColorBrush)?.Color.G ?? 0,
                                ForeB = (a.Foreground as SolidColorBrush)?.Color.B ?? 0,
                                BackR = (a.Background as SolidColorBrush)?.Color.R ?? 255,
                                BackG = (a.Background as SolidColorBrush)?.Color.G ?? 255,
                                BackB = (a.Background as SolidColorBrush)?.Color.B ?? 255,
                                IsHighlight = a.Type == AnnotationType.Highlight
                            });
                        }
                        pagesSnapshot.Add(pageData);
                    }

                    System.Diagnostics.Debug.WriteLine($"SavePdf: Data collection complete. Snapshot count = {pagesSnapshot.Count}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"SavePdf Snapshot Error: {ex.Message}");
                    collectException = ex;
                }
            });

            // 데이터 수집 중 예외 발생 시 다시 throw
            if (collectException != null)
            {
                throw new Exception($"데이터 수집 실패: {collectException.Message}", collectException);
            }

            // 스냅샷이 비어있으면 진행 불가
            if (pagesSnapshot.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("SavePdf: Snapshot is empty after collection.");
                // 원본 모델에 페이지가 있었는데 스냅샷이 비었다면 오류
                // (UI 스레드 외부에서 Pages.Count 접근은 위험할 수 있으나, 여기서는 단순 체크용)
                return; // 페이지가 없으면 저장할 것이 없음
            }

            string tempFilePath = Path.GetTempFileName();
            try
            {
                await Task.Run(() => SavePdfWithIncrementalUpdate(originalFilePath, tempFilePath, pagesSnapshot));
                if (File.Exists(outputPath)) File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}");
            }
            finally { if (File.Exists(tempFilePath)) try { File.Delete(tempFilePath); } catch { } }

            // 저장 후 리로드
            var newModel = await LoadPdfAsync(outputPath);
            if (newModel != null)
            {
                model.PdfDocument?.Dispose();
                model.FileStream?.Dispose();
                model.PdfDocument = newModel.PdfDocument;
                model.FileStream = newModel.FileStream;
                model.FilePath = outputPath;
                model.IsDisposed = false;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var p in model.Pages)
                    {
                        p.Annotations.Clear();
                        p.AnnotationsLoaded = false;
                        p.ImageSource = null;
                    }
                });
            }
        }

        // ==================================================================================
        // CORE: 저수준 PDF 파싱 및 수정 로직 (PdfSharp 대체)
        // ==================================================================================

        private void SavePdfWithIncrementalUpdate(string inputPath, string outputPath, List<PageSaveData> pages)
        {
            byte[] originalBytes = File.ReadAllBytes(inputPath);
            string content = Encoding.Latin1.GetString(originalBytes);

            // 1. 기존 주석들의 Subtype 매핑 (ID -> Subtype 문자열)
            // 목적: 기존 주석 중 지워야 할 것(FreeText/Highlight)과 유지할 것(Widget 등)을 구분
            var annotTypeMap = MapAllAnnotationTypes(content);

            // 2. 페이지 객체 ID 찾기
            var pageObjIds = FindPageObjectIds(content, pages.Count);

            int startxrefPos = content.LastIndexOf("startxref");
            if (startxrefPos < 0) throw new Exception("Invalid PDF: startxref missing");
            var match = Regex.Match(content.Substring(startxrefPos), @"startxref\s+(\d+)");
            long prevXref = long.Parse(match.Groups[1].Value);

            int nextObjNum = FindNextObjectNumber(content);

            using (var output = new FileStream(outputPath, FileMode.Create))
            {
                // 원본 복사 (EOF 제외)
                int eofPos = content.LastIndexOf("%%EOF");
                if (eofPos < 0) eofPos = originalBytes.Length;
                output.Write(originalBytes, 0, eofPos);

                // 줄바꿈용 버퍼
                byte[] newLine = Encoding.Latin1.GetBytes("\r\n");
                var xrefs = new List<(int id, long offset)>();

                // 페이지별 새 주석 참조를 미리 수집 (페이지 객체 수정에 필요)
                var pageAnnotRefs = new Dictionary<int, List<string>>();

                foreach (var page in pages)
                {
                    // 새 주석 객체 생성
                    var newAnnotRefs = new List<string>();
                    foreach (var ann in page.Annotations)
                    {
                        int id = nextObjNum++;

                        // 줄바꿈 먼저 쓰기
                        output.Write(newLine, 0, newLine.Length);

                        // 현재 위치 기록 (정확한 오프셋)
                        long offset = output.Position;
                        xrefs.Add((id, offset));

                        double sx = page.PdfPageWidthPoint / page.Width;
                        double sy = page.PdfPageHeightPoint / page.Height;
                        double px = ann.X * sx;
                        double py = page.PdfPageHeightPoint - (ann.Y * sy) - (ann.Height * sy);
                        double pw = ann.Width * sx;
                        double ph = ann.Height * sy;

                        // 객체 데이터 생성 및 쓰기
                        string objStr = GenerateAnnotObj(id, ann, px, py, pw, ph);
                        byte[] objBytes = Encoding.Latin1.GetBytes(objStr);
                        output.Write(objBytes, 0, objBytes.Length);

                        newAnnotRefs.Add($"{id} 0 R");
                    }

                    int pIdx = page.OriginalPageIndex;
                    if (pIdx < pageObjIds.Count && pageObjIds[pIdx].id > 0)
                    {
                        pageAnnotRefs[pageObjIds[pIdx].id] = newAnnotRefs;
                    }
                }

                // 페이지 객체 업데이트 (/Annots 수정)
                foreach (var page in pages)
                {
                    int pIdx = page.OriginalPageIndex;
                    if (pIdx < pageObjIds.Count && pageObjIds[pIdx].id > 0)
                    {
                        int pid = pageObjIds[pIdx].id;
                        // 기존 /Annots 가져오기
                        var existingRefs = ExtractAnnotRefs(content, pid);

                        // [핵심] 필터링: 기존 주석 중 "FreeText/Highlight"는 제거 (중복 방지), "Widget" 등은 유지
                        var keptRefs = new List<string>();
                        foreach (var refStr in existingRefs)
                        {
                            var refIdMatch = Regex.Match(refStr, @"(\d+)");
                            if (refIdMatch.Success && int.TryParse(refIdMatch.Groups[1].Value, out int refId))
                            {
                                if (annotTypeMap.TryGetValue(refId, out string? subtype))
                                {
                                    // 우리가 관리하는 타입이면 제거(새로 쓸거니까), 아니면 유지
                                    if (subtype != "/FreeText" && subtype != "/Highlight")
                                        keptRefs.Add(refStr);
                                }
                                else keptRefs.Add(refStr); // 타입 모르면 유지
                            }
                        }

                        // 병합
                        if (pageAnnotRefs.TryGetValue(pid, out var newRefs))
                            keptRefs.AddRange(newRefs);

                        string newAnnotsStr = $"/Annots [{string.Join(" ", keptRefs)}]";
                        string modPage = GenerateModifiedPage(content, pid, newAnnotsStr);
                        if (!string.IsNullOrEmpty(modPage))
                        {
                            // 줄바꿈 먼저 쓰기
                            output.Write(newLine, 0, newLine.Length);

                            // 현재 위치 기록 (정확한 오프셋)
                            long offset = output.Position;
                            xrefs.Add((pid, offset));

                            byte[] pageBytes = Encoding.Latin1.GetBytes(modPage);
                            output.Write(pageBytes, 0, pageBytes.Length);
                        }
                    }
                }

                // XREF & Trailer
                WriteXrefAndTrailer(output, prevXref, xrefs);
            }
        }

        private Dictionary<int, string> MapAllAnnotationTypes(string content)
        {
            var map = new Dictionary<int, string>();
            // 모든 객체 중 /Type /Annot 인 것의 Subtype 추출
            // 정규식: "10 0 obj ... /Subtype /Widget ... endobj"
            var matches = Regex.Matches(content, @"(\d+)\s+\d+\s+obj.*?(?:/Type\s*/Annot)?.*?/Subtype\s*(/\w+).*?endobj", RegexOptions.Singleline);
            foreach (Match m in matches)
            {
                int id = int.Parse(m.Groups[1].Value);
                string subtype = m.Groups[2].Value;
                if (!map.ContainsKey(id)) map[id] = subtype;
            }
            return map;
        }

        private List<string> ExtractAnnotRefs(string content, int pageId)
        {
            var refs = new List<string>();
            var match = Regex.Match(content, $@"{pageId}\s+0\s+obj(.*?)endobj", RegexOptions.Singleline);
            if (!match.Success) return refs;

            string body = match.Groups[1].Value;
            var annotMatch = Regex.Match(body, @"/Annots\s*\[([^\]]*)\]");
            if (annotMatch.Success)
            {
                var parts = annotMatch.Groups[1].Value.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < parts.Length; i += 3)
                {
                    if (i + 2 < parts.Length && parts[i + 2] == "R")
                        refs.Add($"{parts[i]} {parts[i + 1]} R");
                }
            }
            return refs;
        }

        private string GenerateModifiedPage(string content, int pageId, string newAnnots)
        {
            var match = Regex.Match(content, $@"{pageId}\s+0\s+obj(.*?)endobj", RegexOptions.Singleline);
            if (!match.Success) return "";
            string body = match.Groups[1].Value;
            // 기존 /Annots 제거
            body = Regex.Replace(body, @"/Annots\s*\[[^\]]*\]", "");
            body = Regex.Replace(body, @"/Annots\s+\d+\s+0\s+R", "");

            // 삽입 위치 (>> 앞)
            int pos = body.LastIndexOf(">>");
            if (pos >= 0) body = body.Insert(pos, "\n" + newAnnots + "\n");

            return $"{pageId} 0 obj{body}endobj\n";
        }

        private string GenerateAnnotObj(int id, AnnotationSaveData a, double x, double y, double w, double h)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{id} 0 obj << /Type /Annot");
            if (a.Type == AnnotationType.FreeText)
            {
                sb.AppendLine($"/Subtype /FreeText /Rect [{x:F2} {y:F2} {x + w:F2} {y + h:F2}]");
                // UTF-16BE Hex String
                var bom = new byte[] { 0xFE, 0xFF };
                var txt = Encoding.BigEndianUnicode.GetBytes(a.TextContent);
                var hex = BitConverter.ToString(bom.Concat(txt).ToArray()).Replace("-", "");
                sb.AppendLine($"/Contents <{hex}>");
                sb.AppendLine($"/DA (/Helv {a.FontSize:F1} Tf {a.ForeR / 255.0:F3} {a.ForeG / 255.0:F3} {a.ForeB / 255.0:F3} rg)");
                sb.AppendLine("/Q 0 /F 4");
            }
            else // Highlight
            {
                sb.AppendLine($"/Subtype /Highlight /Rect [{x:F2} {y:F2} {x + w:F2} {y + h:F2}]");
                sb.AppendLine($"/QuadPoints [{x:F2} {y + h:F2} {x + w:F2} {y + h:F2} {x:F2} {y:F2} {x + w:F2} {y:F2}]");
                sb.AppendLine($"/C [{a.BackR / 255.0:F3} {a.BackG / 255.0:F3} {a.BackB / 255.0:F3}]");
                sb.AppendLine("/F 4");
            }
            sb.AppendLine(">> endobj");
            return sb.ToString();
        }

        private void WriteXrefAndTrailer(Stream s, long prevXref, List<(int id, long offset)> xrefs)
        {
            long startxref = s.Position;
            var sb = new StringBuilder();
            sb.AppendLine("xref");
            // 0번 객체 (dummy)
            sb.AppendLine("0 1");
            sb.AppendLine("0000000000 65535 f ");

            xrefs.Sort((a, b) => a.id.CompareTo(b.id));
            int idx = 0;
            while (idx < xrefs.Count)
            {
                int start = xrefs[idx].id;
                int count = 1;
                while (idx + count < xrefs.Count && xrefs[idx + count].id == start + count) count++;
                sb.AppendLine($"{start} {count}");
                for (int i = 0; i < count; i++)
                    sb.AppendLine($"{xrefs[idx + i].offset:D10} 00000 n ");
                idx += count;
            }
            sb.AppendLine($"trailer << /Prev {prevXref} >>");
            sb.AppendLine("startxref");
            sb.AppendLine(startxref.ToString());
            sb.AppendLine("%%EOF");
            byte[] b = Encoding.Latin1.GetBytes(sb.ToString());
            s.Write(b, 0, b.Length);
        }

        private List<(int id, int gen)> FindPageObjectIds(string content, int count)
        {
            var list = new List<(int id, int gen)>();
            // /Kids [ ... ] 파싱
            var match = Regex.Match(content, @"/Type\s*/Pages.*?/Kids\s*\[([^\]]+)\]", RegexOptions.Singleline);
            if (match.Success)
            {
                var refs = Regex.Matches(match.Groups[1].Value, @"(\d+)\s+(\d+)\s+R");
                foreach (Match r in refs) list.Add((int.Parse(r.Groups[1].Value), int.Parse(r.Groups[2].Value)));
            }
            // fallback
            if (list.Count < count)
            {
                var pages = Regex.Matches(content, @"(\d+)\s+(\d+)\s+obj.*?/Type\s*/Page\b", RegexOptions.Singleline);
                foreach (Match p in pages)
                {
                    int id = int.Parse(p.Groups[1].Value);
                    if (!list.Any(x => x.id == id)) list.Add((id, int.Parse(p.Groups[2].Value)));
                }
            }
            return list;
        }

        private int FindNextObjectNumber(string content)
        {
            var matches = Regex.Matches(content, @"(\d+)\s+\d+\s+obj");
            int max = 0;
            foreach (Match m in matches) max = Math.Max(max, int.Parse(m.Groups[1].Value));
            return max + 1;
        }

        // --- Models for Snapshot ---
        private class PageSaveData
        {
            public int OriginalPageIndex;
            public double Width, Height, PdfPageWidthPoint, PdfPageHeightPoint;
            public List<AnnotationSaveData> Annotations = new List<AnnotationSaveData>();
        }

        private class AnnotationSaveData
        {
            public AnnotationType Type;
            public double X, Y, Width, Height, FontSize;
            public string TextContent = "";
            public byte ForeR, ForeG, ForeB, BackR, BackG, BackB;
            public bool IsHighlight;
        }

        // Bookmarks (기존 유지)
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
