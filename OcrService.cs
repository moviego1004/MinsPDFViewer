using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Security.Cryptography;
using Docnet.Core;
using Docnet.Core.Models;

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
                // 한국어 우선 시도, 실패 시 사용자 프로필 언어 사용
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

        /// <summary>
        /// 문서 전체에 대해 고해상도 OCR을 수행하고 결과를 PageViewModel에 채워넣습니다.
        /// </summary>
        public async Task RunOcrAsync(PdfDocumentModel document, IProgress<int> progress)
        {
            if (_ocrEngine == null || document == null) return;

            string filePath = document.FilePath;

            await Task.Run(async () =>
            {
                // [핵심 1] OCR 인식률 향상을 위해 3.0배 고해상도로 별도 렌더링
                using (var ocrDocReader = DocLib.Instance.GetDocReader(filePath, new PageDimensions(3.0)))
                {
                    for (int i = 0; i < document.Pages.Count; i++)
                    {
                        var pageVM = document.Pages[i];

                        using (var pageReader = ocrDocReader.GetPageReader(i))
                        {
                            var rawBytes = pageReader.GetImage();
                            var w = pageReader.GetPageWidth();
                            var h = pageReader.GetPageHeight();

                            // [핵심 2] 투명 배경을 하얀색으로 변환 (Compositing)
                            // PDFium은 배경을 투명(0,0,0,0)으로 렌더링하지만, 
                            // OCR 엔진은 흰 배경(255,255,255,255)을 선호합니다.
                            // 투명값이 섞인 픽셀을 흰색 배경과 합성합니다.
                            MakeBackgroundWhite(rawBytes);

                            // 좌표 역보정 비율 계산 (고해상도 Width / 원본 PageViewModel Width)
                            double scaleFactor = (double)w / pageVM.Width;

                            using (var stream = new MemoryStream(rawBytes))
                            {
                                // UWP/WinRT용 버퍼 변환
                                var ibuffer = CryptographicBuffer.CreateFromByteArray(rawBytes);
                                using (var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied))
                                {
                                    softwareBitmap.CopyFromBuffer(ibuffer);

                                    // OCR 수행
                                    var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);

                                    // 결과 변환 및 좌표 보정
                                    var wordList = new List<OcrWordInfo>();
                                    foreach (var line in ocrResult.Lines)
                                    {
                                        foreach (var word in line.Words)
                                        {
                                            // 3.0배 확대된 좌표를 다시 1.0배 기준으로 축소
                                            double normX = word.BoundingRect.X / scaleFactor;
                                            double normY = word.BoundingRect.Y / scaleFactor;
                                            double normW = word.BoundingRect.Width / scaleFactor;
                                            double normH = word.BoundingRect.Height / scaleFactor;

                                            wordList.Add(new OcrWordInfo
                                            {
                                                Text = word.Text,
                                                BoundingBox = new Rect(normX, normY, normW, normH)
                                            });
                                        }
                                    }

                                    // ViewModel에 결과 저장
                                    pageVM.OcrWords = wordList;
                                }
                            }
                        }
                        // 진행률 보고
                        progress?.Report(i + 1);
                    }
                }
            });
        }

        // BGRA 바이트 배열에서 투명한 픽셀을 흰색으로 합성하는 헬퍼 메서드
        private void MakeBackgroundWhite(byte[] data)
        {
            // BGRA 32bit: 4바이트씩 순회
            for (int i = 0; i < data.Length; i += 4)
            {
                byte b = data[i];
                byte g = data[i + 1];
                byte r = data[i + 2];
                byte a = data[i + 3];

                if (a < 255) // 완전히 불투명하지 않은 경우
                {
                    // 공식: Output = Source + (1 - Alpha) * Background(White=255)
                    // (Docnet/PDFium은 보통 Pre-multiplied alpha를 반환하므로 단순히 더해주면 흰색과 합성됨)
                    var alphaFactor = 255 - a;
                    
                    data[i]     = (byte)Math.Min(255, b + alphaFactor);
                    data[i + 1] = (byte)Math.Min(255, g + alphaFactor);
                    data[i + 2] = (byte)Math.Min(255, r + alphaFactor);
                    data[i + 3] = 255; // 알파를 100%로 강제
                }
            }
        }
    }
}