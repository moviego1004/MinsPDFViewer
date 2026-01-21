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

        // 5. 저장 - PDFsharp를 사용하여 주석 추가
        public async Task SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || model.Pages.Count == 0)
                return;

            // 1. UI 스레드에서 안전하게 스냅샷 생성
            List<PageSaveData> pagesSnapshot = new List<PageSaveData>();
            string originalFilePath = model.FilePath;

            Action collectData = () =>
            {
                int pageIndex = 0;
                foreach (var p in model.Pages)
                {
                    var pageData = new PageSaveData
                    {
                        PageIndex = pageIndex++,
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
                    // 2. PDFsharp를 사용하여 원본 문서 열기
                    using (var inputDoc = PdfReader.Open(originalFilePath, PdfDocumentOpenMode.Import))
                    {
                        // 3. 새로운 출력용 문서 생성
                        using (var outputDoc = new PdfSharp.Pdf.PdfDocument())
                        {
                            outputDoc.Version = inputDoc.Version;
                            if (inputDoc.Info != null)
                            {
                                outputDoc.Info.Title = inputDoc.Info.Title;
                                outputDoc.Info.Author = inputDoc.Info.Author;
                                outputDoc.Info.Subject = inputDoc.Info.Subject;
                            }

                            // 4. 원본 페이지를 새 문서로 복사 (삭제된 페이지 제외)
                            var pageMapping = new Dictionary<int, int>(); // outputIndex -> snapshotIndex
                            int outputPageIndex = 0;

                            for (int i = 0; i < inputDoc.Pages.Count; i++)
                            {
                                // pagesSnapshot에서 이 원본 페이지 인덱스에 해당하는 데이터 찾기
                                var pageData = pagesSnapshot.FirstOrDefault(p => p.OriginalPageIndex == i);

                                // 페이지가 스냅샷에 있고 빈 페이지가 아니면 복사
                                if (pageData != null && !pageData.IsBlankPage)
                                {
                                    outputDoc.AddPage(inputDoc.Pages[i]);
                                    pageMapping[outputPageIndex] = pagesSnapshot.IndexOf(pageData);
                                    outputPageIndex++;
                                }
                            }

                            // 5. 새 문서의 페이지에 주석 추가
                            for (int i = 0; i < outputDoc.Pages.Count; i++)
                            {
                                if (!pageMapping.TryGetValue(i, out int snapshotIndex))
                                    continue;

                                var pageData = pagesSnapshot[snapshotIndex];
                                var pdfsharpPage = outputDoc.Pages[i];

                                if (pageData.Annotations == null || !pageData.Annotations.Any())
                                    continue;

                                foreach (var annData in pageData.Annotations)
                                {
                                    // UI 좌표 -> PDF 좌표 변환 (Y축이 반대)
                                    double scaleX = pdfsharpPage.Width.Point / pageData.Width;
                                    double scaleY = pdfsharpPage.Height.Point / pageData.Height;

                                    double pdfX = annData.X * scaleX;
                                    double pdfW = annData.Width * scaleX;
                                    double pdfH = annData.Height * scaleY;
                                    double pdfY = pdfsharpPage.Height.Point - (annData.Y * scaleY) - pdfH;

                                    // PDFsharp 저수준 API를 사용하여 annotation 생성
                                    if (annData.Type == AnnotationType.FreeText)
                                    {
                                        AddFreeTextAnnotation(pdfsharpPage, pdfX, pdfY, pdfW, pdfH, annData);
                                    }
                                    else if (annData.Type == AnnotationType.Highlight)
                                    {
                                        AddHighlightAnnotation(pdfsharpPage, pdfX, pdfY, pdfW, pdfH, annData);
                                    }
                                }
                            }

                            // 6. 새 문서 저장
                            outputDoc.Save(tempFilePath);
                        }
                    }
                });

                // 임시 파일을 최종 경로로 이동
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                File.Move(tempFilePath, outputPath);
            }
            catch (Exception ex)
            {
                throw new IOException($"파일 저장 실패: {ex.Message}", ex);
            }
            finally
            {
                if (File.Exists(tempFilePath))
                {
                    try { File.Delete(tempFilePath); } catch { }
                }
            }

            // 7. 저장 후 문서 다시 로드
            var newModel = await LoadPdfAsync(outputPath);
            if (newModel != null)
            {
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
                        p.AnnotationsLoaded = false;
                        p.ImageSource = null;
                    }
                };

                if (Application.Current.Dispatcher.CheckAccess())
                    refreshUI();
                else
                    Application.Current.Dispatcher.Invoke(refreshUI);
            }
        }

        // FreeText annotation 추가 (PDFsharp 저수준 API 사용)
        private void AddFreeTextAnnotation(PdfSharp.Pdf.PdfPage page, double x, double y, double w, double h, AnnotationSaveData annData)
        {
            // Annotation 딕셔너리 생성
            var annotDict = new PdfSharp.Pdf.PdfDictionary(page.Owner);
            annotDict.Elements.SetName("/Type", "/Annot");
            annotDict.Elements.SetName("/Subtype", "/FreeText");

            // Rect 설정
            var rect = new PdfSharp.Pdf.PdfArray(page.Owner);
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(x));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(y));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(x + w));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(y + h));
            annotDict.Elements.SetObject("/Rect", rect);

            // Contents 설정 (UTF-16BE 인코딩을 hex string으로)
            var bom = new byte[] { 0xFE, 0xFF };
            var textBytes = Encoding.BigEndianUnicode.GetBytes(annData.TextContent);
            var combinedBytes = bom.Concat(textBytes).ToArray();
            var hexString = "<" + string.Concat(combinedBytes.Select(b => b.ToString("X2"))) + ">";
            annotDict.Elements["/Contents"] = new PdfSharp.Pdf.PdfLiteral(hexString);

            // DA (Default Appearance) 설정
            double r = annData.ForeR / 255.0;
            double g = annData.ForeG / 255.0;
            double b = annData.ForeB / 255.0;
            annotDict.Elements.SetString("/DA", $"/Helv {annData.FontSize:F1} Tf {r:F3} {g:F3} {b:F3} rg");

            // DR (Default Resources) - 폰트 리소스
            var drDict = new PdfSharp.Pdf.PdfDictionary(page.Owner);
            var fontDict = new PdfSharp.Pdf.PdfDictionary(page.Owner);
            var helvDict = new PdfSharp.Pdf.PdfDictionary(page.Owner);
            helvDict.Elements.SetName("/Type", "/Font");
            helvDict.Elements.SetName("/Subtype", "/Type1");
            helvDict.Elements.SetName("/BaseFont", "/Helvetica");
            fontDict.Elements.SetObject("/Helv", helvDict);
            drDict.Elements.SetObject("/Font", fontDict);
            annotDict.Elements.SetObject("/DR", drDict);

            // 추가 속성
            annotDict.Elements.SetInteger("/F", 4); // Print flag
            annotDict.Elements.SetInteger("/Q", 0); // Left-justified

            // 페이지의 Annots 배열에 추가
            AddAnnotationToPage(page, annotDict);
        }

        // Highlight annotation 추가 (PDFsharp 저수준 API 사용)
        private void AddHighlightAnnotation(PdfSharp.Pdf.PdfPage page, double x, double y, double w, double h, AnnotationSaveData annData)
        {
            var annotDict = new PdfSharp.Pdf.PdfDictionary(page.Owner);
            annotDict.Elements.SetName("/Type", "/Annot");
            annotDict.Elements.SetName("/Subtype", "/Highlight");

            // Rect 설정
            var rect = new PdfSharp.Pdf.PdfArray(page.Owner);
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(x));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(y));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(x + w));
            rect.Elements.Add(new PdfSharp.Pdf.PdfReal(y + h));
            annotDict.Elements.SetObject("/Rect", rect);

            // QuadPoints 설정 (Highlight 영역 정의)
            var quadPoints = new PdfSharp.Pdf.PdfArray(page.Owner);
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(x));           // x1
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(y + h));       // y1 (top-left)
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(x + w));       // x2
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(y + h));       // y2 (top-right)
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(x));           // x3
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(y));           // y3 (bottom-left)
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(x + w));       // x4
            quadPoints.Elements.Add(new PdfSharp.Pdf.PdfReal(y));           // y4 (bottom-right)
            annotDict.Elements.SetObject("/QuadPoints", quadPoints);

            // Color 설정
            var color = new PdfSharp.Pdf.PdfArray(page.Owner);
            color.Elements.Add(new PdfSharp.Pdf.PdfReal(annData.BackR / 255.0));
            color.Elements.Add(new PdfSharp.Pdf.PdfReal(annData.BackG / 255.0));
            color.Elements.Add(new PdfSharp.Pdf.PdfReal(annData.BackB / 255.0));
            annotDict.Elements.SetObject("/C", color);

            // Print flag
            annotDict.Elements.SetInteger("/F", 4);

            AddAnnotationToPage(page, annotDict);
        }

        // 페이지에 annotation 추가하는 헬퍼 메서드
        private static void AddAnnotationToPage(PdfSharp.Pdf.PdfPage page, PdfSharp.Pdf.PdfDictionary annotDict)
        {
            // 페이지의 /Annots 배열 가져오기 (없으면 생성)
            PdfSharp.Pdf.PdfArray? annotsArray = null;
            var annotsItem = page.Elements.GetObject("/Annots");

            if (annotsItem is PdfSharp.Pdf.PdfArray existingArray)
            {
                annotsArray = existingArray;
            }
            else if (annotsItem != null)
            {
                // indirect reference인 경우 dereferencing 시도
                var dereferenced = annotsItem;
                if (dereferenced is PdfSharp.Pdf.PdfArray refArray)
                    annotsArray = refArray;
            }

            if (annotsArray == null)
            {
                annotsArray = new PdfSharp.Pdf.PdfArray(page.Owner);
                page.Elements.SetObject("/Annots", annotsArray);
            }

            // annotation을 indirect object로 추가
            page.Owner.Internals.AddObject(annotDict);
            var reference = annotDict.Reference;
            if (reference != null)
                annotsArray.Elements.Add(reference);
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
            public int PageIndex { get; set; }
            public bool IsBlankPage { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public double PdfPageWidthPoint { get; set; }
            public double PdfPageHeightPoint { get; set; }
            public int OriginalPageIndex { get; set; }
            public List<AnnotationSaveData> Annotations { get; set; } = new List<AnnotationSaveData>();
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