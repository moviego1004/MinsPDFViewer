using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using PdfiumViewer; // Docnet 대체
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Security.Cryptography;

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

        public async Task RunOcrAsync(PdfDocumentModel document, IProgress<int> progress)
        {
            if (_ocrEngine == null || document == null)
                return;

            await Task.Run(async () =>
            {
                // OCR용 별도 로드를 하지 않고 기존 모델의 PdfDocument 사용 (스레드 안전 주의)
                int pageCount = 0;
                lock (PdfService.PdfiumLock)
                {
                    if (document.PdfDocument == null)
                        return;
                    pageCount = document.PdfDocument.PageCount;
                }

                for (int i = 0; i < pageCount; i++)
                {
                    if (document.IsDisposed)
                        break;
                    var pageVM = document.Pages[i];
                    byte[]? rawBytes = null;
                    int w = 0, h = 0;

                    // 3.0배 고해상도 렌더링
                    lock (PdfService.PdfiumLock)
                    {
                        if (document.PdfDocument == null)
                            break;
                        w = (int)(pageVM.Width * 3.0);
                        h = (int)(pageVM.Height * 3.0);
                        // 96 DPI * 3 = 288
                        using (var bitmap = document.PdfDocument.Render(i, w, h, 288, 288, PdfRenderFlags.None))
                        using (var ms = new MemoryStream())
                        {
                            // BGRA32 포맷으로 저장
                            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                            rawBytes = ms.ToArray();
                            // BMP 헤더가 포함되어 있음. OCR 엔진은 SoftwareBitmap 필요.
                            // Windows Runtime Buffer로 변환
                        }
                    }

                    if (rawBytes != null && rawBytes.Length > 0)
                    {
                        // 좌표 역보정 비율
                        double scaleFactor = (double)w / pageVM.Width;

                        // BMP 바이트에서 SoftwareBitmap 생성 (헤더 자동 처리 안됨, 이미지 디코더 사용 권장)
                        // 여기서는 간편하게 MemoryStream -> BitmapDecoder 사용
                        using (var ms = new MemoryStream(rawBytes))
                        {
                            var decoder = await BitmapDecoder.CreateAsync(ms.AsRandomAccessStream());
                            using (var softwareBitmap = await decoder.GetSoftwareBitmapAsync())
                            {
                                var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);
                                var wordList = new List<OcrWordInfo>();

                                foreach (var line in ocrResult.Lines)
                                {
                                    foreach (var word in line.Words)
                                    {
                                        double normX = word.BoundingRect.X / scaleFactor;
                                        double normY = word.BoundingRect.Y / scaleFactor;
                                        double normW = word.BoundingRect.Width / scaleFactor;
                                        double normH = word.BoundingRect.Height / scaleFactor;
                                        wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(normX, normY, normW, normH) });
                                    }
                                }
                                pageVM.OcrWords = wordList;
                            }
                        }
                    }
                    progress?.Report(i + 1);
                }
            });
        }
    }
}