using System;
using System.Collections.Generic;
using System.IO;
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

        public async Task RunOcrAsync(PdfDocumentModel document, IProgress<int> progress)
        {
            if (_ocrEngine == null || document == null)
                return;

            await Task.Run(async () =>
            {
                int pageCount = 0;
                lock (PdfService.PdfiumLock)
                {
                    if (document.PdfDocument == null)
                        return;
                    pageCount = Math.Min(document.PdfDocument.PageCount, document.Pages.Count);
                }

                for (int i = 0; i < pageCount; i++)
                {
                    if (document.IsDisposed)
                        break;

                    var pageVM = document.Pages[i];
                    byte[]? rawBytes = null;
                    int w = 0, h = 0;

                    try
                    {
                        // Windows OCR accuracy improves noticeably with 2.5x-3x input.
                        const double renderScale = 3.0;
                        w = Math.Max(1, (int)Math.Round(pageVM.Width * renderScale));
                        h = Math.Max(1, (int)Math.Round(pageVM.Height * renderScale));

                        lock (PdfService.PdfiumLock)
                        {
                            if (document.PdfDocument == null)
                                break;

                            using var bitmap = document.PdfDocument.Render(i, w, h, 288, 288, PdfRenderFlags.Annotations);
                            using var ms = new MemoryStream();
                            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                            rawBytes = ms.ToArray();
                        }

                        if (rawBytes == null || rawBytes.Length == 0)
                            continue;

                        double scaleX = w / pageVM.Width;
                        double scaleY = h / pageVM.Height;

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
                        Debug.WriteLine($"OCR page {i + 1}/{pageCount}: {wordList.Count} words");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"OCR page {i + 1} failed: {ex}");
                        throw;
                    }

                    progress?.Report(i + 1);
                }
            });
        }
    }
}
