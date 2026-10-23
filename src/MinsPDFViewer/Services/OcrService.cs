using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows;
using PdfiumViewer; // Docnet 대체
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace MinsPDFViewer
{
    public class OcrService
    {
        private OcrEngine? _ocrEngine;

        public OcrService()
        {
            InitializeOcrEngine();
        }

        private void InitializeOcrEngine()
        {
            try
            {
                _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR"))
                             ?? OcrEngine.TryCreateFromLanguage(new Language("ko"))
                             ?? OcrEngine.TryCreateFromUserProfileLanguages();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OCR 엔진 초기화 오류: {ex.Message}");
            }
        }

        public bool IsAvailable => _ocrEngine != null;
        public string CurrentLanguage => _ocrEngine?.RecognizerLanguage.LanguageTag ?? "알 수 없음";

        /// <summary>OCR 대상 페이지 수 (빈 페이지 제외). 진행률 최대값으로 쓴다.</summary>
        public static int CountTargetPages(PdfDocumentModel document) => document.Pages.Count(p => !p.IsBlankPage);

        /// <summary>
        /// 화면 목록 순서대로 각 페이지의 원본(<see cref="PdfPageViewModel.OriginalPageIndex"/>)을 그려 OCR한다.
        /// 단어 좌표는 줌과 무관한 PDF 포인트(원본 페이지 화면 기준, 왼쪽 위 원점)로 저장한다.
        /// </summary>
        public async Task RunOcrAsync(PdfDocumentModel document, IProgress<int> progress)
        {
            if (_ocrEngine == null || document == null)
                return;

            // 호출 스레드(UI)에서 목록을 고정: OCR 중 페이지 이동·삭제가 있어도 대상이 흔들리지 않게.
            var targets = document.Pages.Where(p => !p.IsBlankPage).ToList();
            uint maxSide = Math.Max(1000u, OcrEngine.MaxImageDimension);

            await Task.Run(async () =>
            {
                int done = 0;
                foreach (var pageVM in targets)
                {
                    if (document.IsDisposed)
                        break;

                    byte[]? rawBytes = null;
                    int w = 0, h = 0;
                    double viewW = pageVM.SourceViewWidthPoint;
                    double viewH = pageVM.SourceViewHeightPoint;

                    try
                    {
                        if (viewW <= 0 || viewH <= 0)
                        {
                            progress?.Report(++done);
                            continue;
                        }

                        // 216 dpi (포인트의 3배). Windows OCR은 2.5~3배 입력에서 정확도가 좋다. 엔진 한계를 넘지 않게 줄인다.
                        double renderScale = Math.Min(3.0, maxSide / Math.Max(viewW, viewH));
                        w = Math.Max(1, (int)Math.Round(viewW * renderScale));
                        h = Math.Max(1, (int)Math.Round(viewH * renderScale));
                        float dpi = (float)(72.0 * renderScale);

                        lock (PdfService.PdfiumLock)
                        {
                            if (document.PdfDocument == null || document.IsDisposed)
                                break;
                            if (pageVM.OriginalPageIndex < 0 || pageVM.OriginalPageIndex >= document.PdfDocument.PageCount)
                            {
                                Debug.WriteLine($"OCR skipped: invalid source page {pageVM.OriginalPageIndex}");
                                rawBytes = null;
                            }
                            else
                            {
                                using var bitmap = document.PdfDocument.Render(pageVM.OriginalPageIndex, w, h, dpi, dpi, PdfRenderFlags.Annotations);
                                using var ms = new MemoryStream();
                                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                rawBytes = ms.ToArray();
                            }
                        }

                        if (rawBytes == null || rawBytes.Length == 0)
                        {
                            progress?.Report(++done);
                            continue;
                        }

                        // 픽셀 → 포인트
                        double scaleX = w / viewW;
                        double scaleY = h / viewH;

                        using var imageStream = new MemoryStream(rawBytes);
                        var decoder = await BitmapDecoder.CreateAsync(imageStream.AsRandomAccessStream());
                        using var softwareBitmap = await decoder.GetSoftwareBitmapAsync(
                            BitmapPixelFormat.Bgra8,
                            BitmapAlphaMode.Premultiplied);

                        var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);
                        var wordList = new List<OcrWordInfo>();

                        foreach (var line in ocrResult.Lines)
                        {
                            foreach (var word in line.Words)
                            {
                                double normX = word.BoundingRect.X / scaleX;
                                double normY = word.BoundingRect.Y / scaleY;
                                double normW = word.BoundingRect.Width / scaleX;
                                double normH = word.BoundingRect.Height / scaleY;
                                wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(normX, normY, normW, normH) });
                            }
                        }

                        pageVM.OcrWords = wordList;
                        Debug.WriteLine($"OCR page {pageVM.PageNumber} (source {pageVM.OriginalPageIndex + 1}): {wordList.Count} words");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"OCR page {pageVM.PageNumber} failed: {ex}");
                        throw;
                    }

                    progress?.Report(++done);
                }
            });
        }
    }
}
