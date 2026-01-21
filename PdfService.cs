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
using PdfSharp.Pdf;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingRectangle = System.Drawing.Rectangle;
using PdfiumDoc = PdfiumViewer.PdfDocument;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private const double RENDER_SCALE = 2.0;

        // 1. PDF 로드
        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            return await Task.Run(() =>
            {
                try
                {
                    byte[] fileBytes = File.ReadAllBytes(filePath);
                    var memoryStream = new MemoryStream(fileBytes);
                    IPdfDocument? pdfDoc = null;

                    lock (PdfiumLock)
                    {
                        pdfDoc = PdfiumDoc.Load(memoryStream);
                    }

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
            if (model.PdfDocument == null || model.IsDisposed)
                return;
            await Task.Run(() =>
            {
                if (model.IsDisposed)
                    return;
                int pageCount = 0;
                lock (PdfiumLock)
                {
                    if (model.PdfDocument == null)
                        return;
                    pageCount = model.PdfDocument.PageCount;
                }
                if (pageCount == 0)
                    return;

                var tempPageList = new List<PdfPageViewModel>();
                for (int i = 0; i < pageCount; i++)
                {
                    if (model.IsDisposed)
                        return;
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
                if (!model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() => { foreach (var p in tempPageList) model.Pages.Add(p); });
                }
            });
        }

        // 3. 렌더링
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (model.IsDisposed || pageVM.IsBlankPage)
                return;
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
                            int dpi = (int)(96.0 * RENDER_SCALE);

                            using (var bitmap = model.PdfDocument.Render(pageVM.OriginalPageIndex, renderW, renderH, dpi, dpi, PdfRenderFlags.Annotations))
                            {
                                bmpSource = ToBitmapSource((DrawingBitmap)bitmap);
                                bmpSource.Freeze();
                            }
                        }
                        catch { }
                    }
                }
                if (bmpSource != null && !model.IsDisposed)
                {
                    Application.Current.Dispatcher.Invoke(() => { pageVM.ImageSource = bmpSource; });
                }
            }
            LoadAnnotationsLazy(model, pageVM);
        }

        public static BitmapSource ToBitmapSource(DrawingBitmap bitmap)
        {
            var rect = new DrawingRectangle(0, 0, bitmap.Width, bitmap.Height);
            var bitmapData = bitmap.LockBits(rect, System.Drawing.Imaging.ImageLockMode.ReadOnly, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            try
            {
                return BitmapSource.Create(bitmap.Width, bitmap.Height, bitmap.HorizontalResolution, bitmap.VerticalResolution, PixelFormats.Bgra32, null, bitmapData.Scan0, bitmapData.Stride * bitmap.Height, bitmapData.Stride);
            }
            finally { bitmap.UnlockBits(bitmapData); }
        }

        // 4. 주석 로드 - PDF 바이트에서 직접 파싱 (PdfSharp 우회)
        // [개선] PdfSharp를 사용하지 않고 직접 바이트 파싱만 사용 (PDF 1.5+ 지원)
        private void LoadAnnotationsLazy(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            // [FIX] 스레드 안전성: 이미 로드된 경우 스킵
            if (pageVM.AnnotationsLoaded)
                return;
            pageVM.AnnotationsLoaded = true;

            Task.Run(() =>
            {
                try
                {
                    if (model.IsDisposed)
                        return;

                    string path = pageVM.OriginalFilePath ?? model.FilePath;
                    if (!File.Exists(path))
                        return;

                    // [변경] PdfSharp 대신 직접 바이트 파싱 사용 (PDF 1.5+ 호환)
                    var extracted = ParseAnnotationsFromBytes(path, pageVM.OriginalPageIndex, pageVM.Width, pageVM.Height);

                    // UI 업데이트
                    if (extracted != null && extracted.Count > 0)
                    {
                        var annotationsToAdd = new List<PdfAnnotation>();
                        foreach (var item in extracted)
                        {
                            var newAnnot = new PdfAnnotation
                            {
                                Type = item.Type,
                                X = item.X,
                                Y = item.Y,
                                Width = item.Width,
                                Height = item.Height,
                                TextContent = item.TextContent,
                                FontSize = item.FontSize,
                                Foreground = item.Foreground ?? Brushes.Black,
                                Background = item.Background ?? Brushes.Transparent
                            };
                            annotationsToAdd.Add(newAnnot);
                        }

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            if (pageVM.Annotations.Count == 0)
                            {
                                foreach (var annot in annotationsToAdd)
                                {
                                    pageVM.Annotations.Add(annot);
                                }
                            }
                        });
                    }
                }
                catch { }
            });
        }

        // [신규] PDF 바이트에서 직접 annotation 파싱 (PdfSharp 실패 시 Fallback)
        private List<MPdfAnnotation> ParseAnnotationsFromBytes(string filePath, int pageIndex, double uiWidth, double uiHeight)
        {
            var result = new List<MPdfAnnotation>();
            try
            {
                string content = File.ReadAllText(filePath, Encoding.Latin1);

                // 페이지 MediaBox에서 크기 추출 (간단한 휴리스틱)
                double pageW = 612, pageH = 792; // 기본값 (Letter 사이즈)
                var mediaBoxMatch = Regex.Match(content, @"/MediaBox\s*\[\s*([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s*\]");
                if (mediaBoxMatch.Success)
                {
                    double.TryParse(mediaBoxMatch.Groups[3].Value, out pageW);
                    double.TryParse(mediaBoxMatch.Groups[4].Value, out pageH);
                }

                double scaleX = uiWidth / pageW;
                double scaleY = uiHeight / pageH;

                // [개선] FreeText annotation 객체 전체를 먼저 찾은 후 내부 파싱
                // 패턴: "숫자 0 obj" ~ "endobj" 범위에서 /Subtype /FreeText 포함된 것
                var annotObjMatches = Regex.Matches(content,
                    @"\d+\s+\d+\s+obj\s*<<(.*?)>>\s*endobj",
                    RegexOptions.Singleline);

                foreach (Match objMatch in annotObjMatches)
                {
                    string objContent = objMatch.Groups[1].Value;

                    // FreeText인지 확인
                    if (!Regex.IsMatch(objContent, @"/Subtype\s*/FreeText"))
                        continue;

                    // /Rect 추출
                    var rectMatch = Regex.Match(objContent, @"/Rect\s*\[\s*([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s+([\d.-]+)\s*\]");
                    if (!rectMatch.Success)
                        continue;

                    double.TryParse(rectMatch.Groups[1].Value, out double rx1);
                    double.TryParse(rectMatch.Groups[2].Value, out double ry1);
                    double.TryParse(rectMatch.Groups[3].Value, out double rx2);
                    double.TryParse(rectMatch.Groups[4].Value, out double ry2);

                    // /Contents 추출 (문자열 또는 hex)
                    string textContent = "";
                    var contentsMatch = Regex.Match(objContent, @"/Contents\s*\(([^)]*)\)");
                    if (contentsMatch.Success)
                    {
                        textContent = UnescapePdfString(contentsMatch.Groups[1].Value);
                    }
                    else
                    {
                        var hexMatch = Regex.Match(objContent, @"/Contents\s*<([^>]*)>");
                        if (hexMatch.Success)
                            textContent = DecodeHexString(hexMatch.Groups[1].Value);
                    }

                    double pdfX = Math.Min(rx1, rx2);
                    double pdfY = Math.Min(ry1, ry2);
                    double pdfW = Math.Abs(rx2 - rx1);
                    double pdfH = Math.Abs(ry2 - ry1);

                    double uiX = pdfX * scaleX;
                    double uiY = (pageH - pdfY - pdfH) * scaleY;
                    double uiW = pdfW * scaleX;
                    double uiH = pdfH * scaleY;

                    result.Add(new MPdfAnnotation
                    {
                        Type = AnnotationType.FreeText,
                        X = uiX, Y = uiY, Width = Math.Max(uiW, 50), Height = Math.Max(uiH, 20),
                        TextContent = textContent,
                        FontSize = 12,
                        Foreground = Brushes.Black,
                        Background = Brushes.Transparent
                    });
                }
            }
            catch { }
            return result;
        }

        private string UnescapePdfString(string escaped)
        {
            if (string.IsNullOrEmpty(escaped)) return "";
            return escaped
                .Replace("\\\\", "\x00")  // 임시 치환
                .Replace("\\(", "(")
                .Replace("\\)", ")")
                .Replace("\\r", "\r")
                .Replace("\\n", "\n")
                .Replace("\x00", "\\");   // 복원
        }

        private string DecodeHexString(string hex)
        {
            try
            {
                hex = hex.Replace(" ", "").Replace("\n", "").Replace("\r", "");
                var bytes = new byte[hex.Length / 2];
                for (int i = 0; i < bytes.Length; i++)
                    bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);

                // UTF-16BE 또는 Latin1 시도
                if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                    return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
                return Encoding.Latin1.GetString(bytes);
            }
            catch { return ""; }
        }

        // 5. 저장 - PDF incremental update 방식으로 주석 추가 (PdfSharp 우회)
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            // UI 스레드에서 안전하게 스냅샷 생성
            List<PageSaveData> pagesSnapshot = new List<PageSaveData>();
            string originalFilePath = model.FilePath;

            Action collectData = () =>
            {
                foreach (var p in model.Pages)
                {
                    var pageData = new PageSaveData
                    {
                        IsBlankPage = p.IsBlankPage,
                        Width = p.Width,
                        Height = p.Height,
                        PdfPageWidthPoint = p.PdfPageWidthPoint,
                        PdfPageHeightPoint = p.PdfPageHeightPoint,
                        OriginalPageIndex = p.OriginalPageIndex,
                        Annotations = new List<AnnotationSaveData>()
                    };

                    foreach (var a in p.Annotations)
                    {
                        // SearchHighlight는 저장하지 않음
                        if (a.Type == AnnotationType.SearchHighlight)
                            continue;

                        var annotData = new AnnotationSaveData
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
                            BackA = (a.Background as SolidColorBrush)?.Color.A ?? 255,
                            IsHighlight = a.Type == AnnotationType.Highlight
                        };
                        pageData.Annotations.Add(annotData);
                    }

                    pagesSnapshot.Add(pageData);
                }
            };

            if (Application.Current.Dispatcher.CheckAccess())
                collectData();
            else
                await Application.Current.Dispatcher.InvokeAsync(collectData);

            string tempFilePath = Path.GetTempFileName();

            try
            {
                await Task.Run(() =>
                {
                    // [개선] PDF incremental update 방식으로 직접 바이트 조작
                    // PdfSharp를 우회하여 PDF 1.5+ Object Streams도 지원
                    SavePdfWithIncrementalUpdate(originalFilePath, tempFilePath, pagesSnapshot);
                });

                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex)
            {
                // 실패 시 원본 복원 시도
                try
                {
                    var restoreModel = await LoadPdfAsync(originalFilePath);
                    if (restoreModel != null)
                    {
                        model.IsDisposed = false;
                        model.PdfDocument = restoreModel.PdfDocument;
                        model.FileStream = restoreModel.FileStream;
                    }
                }
                catch { }

                throw new IOException($"파일 저장 실패: {ex.Message}");
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { }
                }
            }

            // 저장 후 파일 다시 로드
            var newModel = await LoadPdfAsync(outputPath);
            if (newModel != null)
            {
                // [수정] 기존 리소스 해제
                model.PdfDocument?.Dispose();
                model.FileStream?.Dispose();

                model.IsDisposed = false;
                model.PdfDocument = newModel.PdfDocument;
                model.FileStream = newModel.FileStream;
                model.FilePath = outputPath;

                Action refreshUI = () =>
                {
                    foreach (var p in model.Pages)
                    {
                        p.Annotations.Clear();
                        p.AnnotationsLoaded = false;  // [FIX] 저장 후 annotation 다시 로드하도록 리셋
                        p.ImageSource = null;
                    }
                };

                if (Application.Current.Dispatcher.CheckAccess())
                    refreshUI();
                else
                    Application.Current.Dispatcher.Invoke(refreshUI);
            }
        }

        public void LoadBookmarks(PdfDocumentModel model)
        {
            // [FIX] Lock 안에서 데이터 추출 후, Lock 밖에서 UI 업데이트
            System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>? bookmarkVms = null;

            lock (PdfiumLock)
            {
                if (model.PdfDocument == null)
                    return;
                var bookmarks = model.PdfDocument.Bookmarks;
                if (bookmarks != null && bookmarks.Count > 0)
                {
                    bookmarkVms = new System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>();
                    foreach (var bm in bookmarks)
                    {
                        var vm = ConvertToViewModel(bm, null);
                        bookmarkVms.Add(vm);
                    }
                }
            }

            // Lock 밖에서 UI 업데이트
            if (bookmarkVms != null)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    model.Bookmarks.Clear();
                    foreach (var vm in bookmarkVms)
                    {
                        model.Bookmarks.Add(vm);
                    }
                });
            }
        }
        private PdfBookmarkViewModel ConvertToViewModel(PdfBookmark bookmark, PdfBookmarkViewModel? parent)
        {
            var vm = new PdfBookmarkViewModel { Title = bookmark.Title, PageIndex = bookmark.PageIndex, Parent = parent, IsExpanded = true };
            foreach (var child in bookmark.Children)
                vm.Children.Add(ConvertToViewModel(child, vm));
            return vm;
        }

        public class MPdfAnnotation
        {
            public AnnotationType Type
            {
                get; set;
            }
            public double X
            {
                get; set;
            }
            public double Y
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public string TextContent { get; set; } = ""; public double FontSize
            {
                get; set;
            }
            public Brush? Foreground
            {
                get; set;
            }
            public Brush? Background
            {
                get; set;
            }
        }
        private class PageSaveData
        {
            public bool IsBlankPage
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public double PdfPageWidthPoint
            {
                get; set;
            }
            public double PdfPageHeightPoint
            {
                get; set;
            }
            public int OriginalPageIndex
            {
                get; set;
            }
            public List<AnnotationSaveData> Annotations { get; set; } = new List<AnnotationSaveData>();
        }

        // [신규] PDF Incremental Update로 annotation 저장 (PdfSharp 우회)
        private void SavePdfWithIncrementalUpdate(string inputPath, string outputPath, List<PageSaveData> pages)
        {
            // 원본 PDF 바이트 읽기
            byte[] originalBytes = File.ReadAllBytes(inputPath);
            string originalContent = Encoding.Latin1.GetString(originalBytes);

            // 저장할 annotation이 있는지 확인
            bool hasAnnotations = pages.Any(p => p.Annotations.Count > 0);
            if (!hasAnnotations)
            {
                // annotation이 없으면 원본 그대로 복사
                File.Copy(inputPath, outputPath, true);
                return;
            }

            // PDF 구조 파싱: xref 테이블 위치 및 trailer 찾기
            int startxrefPos = originalContent.LastIndexOf("startxref");
            if (startxrefPos < 0)
                throw new Exception("Invalid PDF: startxref not found");

            // 기존 xref 오프셋 추출
            var startxrefMatch = Regex.Match(originalContent.Substring(startxrefPos), @"startxref\s+(\d+)");
            if (!startxrefMatch.Success)
                throw new Exception("Invalid PDF: cannot parse startxref");
            long prevXrefOffset = long.Parse(startxrefMatch.Groups[1].Value);

            // 페이지 객체 ID 찾기 (간단한 휴리스틱)
            var pageObjIds = FindPageObjectIds(originalContent, pages.Count);

            // 다음 사용할 객체 번호 찾기
            int nextObjNum = FindNextObjectNumber(originalContent);

            // Incremental update 데이터 생성
            using (var output = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                // 1. 원본 PDF 그대로 쓰기 (%%EOF 제외)
                int eofPos = originalContent.LastIndexOf("%%EOF");
                if (eofPos < 0) eofPos = originalBytes.Length;
                output.Write(originalBytes, 0, eofPos);

                // 2. 새 annotation 객체들과 수정된 페이지 객체 추가
                long newObjectsStart = output.Position;
                var newXrefEntries = new List<(int objNum, long offset)>();
                var sb = new StringBuilder();
                sb.AppendLine();

                for (int pageIdx = 0; pageIdx < pages.Count; pageIdx++)
                {
                    var pageData = pages[pageIdx];
                    if (pageData.Annotations.Count == 0)
                        continue;

                    // 이 페이지의 annotation 객체들 생성
                    var annotRefs = new List<string>();

                    foreach (var ann in pageData.Annotations)
                    {
                        int annotObjNum = nextObjNum++;
                        long annotOffset = newObjectsStart + Encoding.Latin1.GetByteCount(sb.ToString());

                        // UI 좌표 -> PDF 좌표 변환
                        double scaleX = pageData.PdfPageWidthPoint / pageData.Width;
                        double scaleY = pageData.PdfPageHeightPoint / pageData.Height;

                        double pdfX = ann.X * scaleX;
                        double pdfW = ann.Width * scaleX;
                        double pdfH = ann.Height * scaleY;
                        double pdfY = pageData.PdfPageHeightPoint - (ann.Y * scaleY) - pdfH;

                        string annotObj = GenerateAnnotationObject(annotObjNum, ann, pdfX, pdfY, pdfW, pdfH);
                        sb.Append(annotObj);
                        newXrefEntries.Add((annotObjNum, annotOffset));
                        annotRefs.Add($"{annotObjNum} 0 R");
                    }

                    // 페이지 객체에 /Annots 배열 추가 (기존 Annots가 있으면 병합 필요)
                    if (pageIdx < pageObjIds.Count && pageObjIds[pageIdx].objNum > 0)
                    {
                        int pageObjNum = pageObjIds[pageIdx].objNum;
                        long pageObjOffset = newObjectsStart + Encoding.Latin1.GetByteCount(sb.ToString());

                        // 기존 페이지 객체에서 /Annots 추출 시도
                        string existingAnnots = ExtractExistingAnnots(originalContent, pageObjNum);
                        string allAnnotRefs = existingAnnots;
                        if (!string.IsNullOrEmpty(allAnnotRefs))
                            allAnnotRefs += " " + string.Join(" ", annotRefs);
                        else
                            allAnnotRefs = string.Join(" ", annotRefs);

                        // 수정된 페이지 객체 생성
                        string modifiedPageObj = GenerateModifiedPageObject(originalContent, pageObjNum, allAnnotRefs);
                        if (!string.IsNullOrEmpty(modifiedPageObj))
                        {
                            sb.Append(modifiedPageObj);
                            newXrefEntries.Add((pageObjNum, pageObjOffset));
                        }
                    }
                }

                // 새 객체들 쓰기
                byte[] newObjBytes = Encoding.Latin1.GetBytes(sb.ToString());
                output.Write(newObjBytes, 0, newObjBytes.Length);

                // 3. 새 xref 테이블 작성
                long newXrefOffset = output.Position;
                var xrefSb = new StringBuilder();
                xrefSb.AppendLine("xref");

                // 각 객체별로 xref 엔트리 작성 (연속 범위로 그룹화)
                var sortedEntries = newXrefEntries.OrderBy(e => e.objNum).ToList();
                int i = 0;
                while (i < sortedEntries.Count)
                {
                    int startNum = sortedEntries[i].objNum;
                    int count = 1;

                    // 연속된 객체 번호 찾기
                    while (i + count < sortedEntries.Count && sortedEntries[i + count].objNum == startNum + count)
                        count++;

                    xrefSb.AppendLine($"{startNum} {count}");
                    for (int j = 0; j < count; j++)
                    {
                        long offset = sortedEntries[i + j].offset;
                        xrefSb.AppendLine($"{offset:D10} 00000 n ");
                    }
                    i += count;
                }

                byte[] xrefBytes = Encoding.Latin1.GetBytes(xrefSb.ToString());
                output.Write(xrefBytes, 0, xrefBytes.Length);

                // 4. 새 trailer 작성
                string trailerStr = $"trailer\n<< /Prev {prevXrefOffset} >>\nstartxref\n{newXrefOffset}\n%%EOF\n";
                byte[] trailerBytes = Encoding.Latin1.GetBytes(trailerStr);
                output.Write(trailerBytes, 0, trailerBytes.Length);
            }
        }

        private List<(int objNum, int gen)> FindPageObjectIds(string content, int pageCount)
        {
            var result = new List<(int objNum, int gen)>();

            // /Type /Page 패턴으로 페이지 객체 찾기
            var pageMatches = Regex.Matches(content, @"(\d+)\s+(\d+)\s+obj[^>]*?/Type\s*/Page\b", RegexOptions.Singleline);
            foreach (Match m in pageMatches)
            {
                if (result.Count >= pageCount) break;
                int objNum = int.Parse(m.Groups[1].Value);
                int gen = int.Parse(m.Groups[2].Value);
                result.Add((objNum, gen));
            }

            // 페이지를 못 찾은 경우 빈 항목 추가
            while (result.Count < pageCount)
                result.Add((0, 0));

            return result;
        }

        private int FindNextObjectNumber(string content)
        {
            int maxObjNum = 0;
            var objMatches = Regex.Matches(content, @"(\d+)\s+\d+\s+obj\b");
            foreach (Match m in objMatches)
            {
                int objNum = int.Parse(m.Groups[1].Value);
                if (objNum > maxObjNum) maxObjNum = objNum;
            }
            return maxObjNum + 1;
        }

        private string GenerateAnnotationObject(int objNum, AnnotationSaveData ann, double x, double y, double w, double h)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{objNum} 0 obj");
            sb.AppendLine("<<");
            sb.AppendLine("/Type /Annot");

            if (ann.Type == AnnotationType.FreeText)
            {
                sb.AppendLine("/Subtype /FreeText");
                sb.AppendLine($"/Rect [{x:F2} {y:F2} {(x + w):F2} {(y + h):F2}]");

                // Contents - 텍스트 내용 (특수문자 이스케이프)
                string escapedText = EscapePdfString(ann.TextContent);
                sb.AppendLine($"/Contents ({escapedText})");

                // DA - Default Appearance
                double r = ann.ForeR / 255.0;
                double g = ann.ForeG / 255.0;
                double b = ann.ForeB / 255.0;
                sb.AppendLine($"/DA (/Helv {ann.FontSize:F0} Tf {r:F2} {g:F2} {b:F2} rg)");

                // 추가 속성
                sb.AppendLine("/F 4"); // Print flag
                sb.AppendLine("/Q 0"); // Left-justified
            }
            else if (ann.Type == AnnotationType.Highlight)
            {
                sb.AppendLine("/Subtype /Highlight");
                sb.AppendLine($"/Rect [{x:F2} {y:F2} {(x + w):F2} {(y + h):F2}]");

                // QuadPoints
                sb.AppendLine($"/QuadPoints [{x:F2} {(y + h):F2} {(x + w):F2} {(y + h):F2} {x:F2} {y:F2} {(x + w):F2} {y:F2}]");

                // 색상
                double r = ann.BackR / 255.0;
                double g = ann.BackG / 255.0;
                double b = ann.BackB / 255.0;
                sb.AppendLine($"/C [{r:F2} {g:F2} {b:F2}]");

                sb.AppendLine("/F 4"); // Print flag
            }

            sb.AppendLine(">>");
            sb.AppendLine("endobj");
            return sb.ToString();
        }

        private string EscapePdfString(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\", "\\\\")
                .Replace("(", "\\(")
                .Replace(")", "\\)")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }

        private string ExtractExistingAnnots(string content, int pageObjNum)
        {
            // 해당 페이지 객체에서 기존 /Annots 배열 추출
            var pageObjMatch = Regex.Match(content, $@"{pageObjNum}\s+0\s+obj(.*?)endobj", RegexOptions.Singleline);
            if (!pageObjMatch.Success) return "";

            string pageObjContent = pageObjMatch.Groups[1].Value;
            var annotsMatch = Regex.Match(pageObjContent, @"/Annots\s*\[([^\]]*)\]");
            if (annotsMatch.Success)
                return annotsMatch.Groups[1].Value.Trim();

            // /Annots가 간접 참조인 경우
            var annotsRefMatch = Regex.Match(pageObjContent, @"/Annots\s+(\d+)\s+0\s+R");
            if (annotsRefMatch.Success)
                return annotsRefMatch.Groups[1].Value + " 0 R";

            return "";
        }

        private string GenerateModifiedPageObject(string content, int pageObjNum, string annotRefs)
        {
            // 기존 페이지 객체 찾기
            var pageObjMatch = Regex.Match(content, $@"{pageObjNum}\s+0\s+obj(.*?)endobj", RegexOptions.Singleline);
            if (!pageObjMatch.Success) return "";

            string pageObjContent = pageObjMatch.Groups[1].Value;

            // 기존 /Annots 제거
            pageObjContent = Regex.Replace(pageObjContent, @"/Annots\s*\[[^\]]*\]", "");
            pageObjContent = Regex.Replace(pageObjContent, @"/Annots\s+\d+\s+0\s+R", "");

            // >> 앞에 새 /Annots 추가
            int insertPos = pageObjContent.LastIndexOf(">>");
            if (insertPos < 0) return "";

            string newAnnotsEntry = $"\n/Annots [{annotRefs}]\n";
            pageObjContent = pageObjContent.Insert(insertPos, newAnnotsEntry);

            return $"{pageObjNum} 0 obj{pageObjContent}endobj\n";
        }

        private class AnnotationSaveData
        {
            public AnnotationType Type
            {
                get; set;
            }
            public double X
            {
                get; set;
            }
            public double Y
            {
                get; set;
            }
            public double Width
            {
                get; set;
            }
            public double Height
            {
                get; set;
            }
            public string TextContent { get; set; } = ""; public double FontSize
            {
                get; set;
            }
            public byte ForeR
            {
                get; set;
            }
            public byte ForeG
            {
                get; set;
            }
            public byte ForeB
            {
                get; set;
            }
            public byte BackR
            {
                get; set;
            }
            public byte BackG
            {
                get; set;
            }
            public byte BackB
            {
                get; set;
            }
            public byte BackA
            {
                get; set;
            }
            public bool IsHighlight
            {
                get; set;
            }
        }
    }
}