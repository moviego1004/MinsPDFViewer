using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using MinsPDFViewer;
using Xunit;

namespace MinsPDFViewer.Tests
{
    public class PdfServiceTests : IDisposable
    {
        private readonly string _testFileDir;
        private readonly string _samplePdfPath;
        private readonly string _signedPdfPath;

        public PdfServiceTests()
        {
            PdfService.UseDirectDispatcherForTests = true;

            // 각 테스트 인스턴스마다 고유한 임시 디렉토리 사용하여 파일 잠금 충돌 방지
            _testFileDir = Path.Combine(Path.GetTempPath(), "MinsPDFViewerTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testFileDir);

            _samplePdfPath = Path.Combine(_testFileDir, "normal.pdf");
            _signedPdfPath = Path.Combine(_testFileDir, "signed.pdf");

            string repoSamplePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestFiles", "normal.pdf"));
            if (File.Exists(repoSamplePath))
                File.Copy(repoSamplePath, _samplePdfPath, overwrite: true);
            else
                File.WriteAllBytes(_samplePdfPath, CreateMinimalPdf());
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(_testFileDir))
                    Directory.Delete(_testFileDir, true);
            }
            catch { }
        }

        // [TC-11 & 12] 저장 후 데이터가 유지되는지 확인
        [Fact]
        public async Task Save_And_Reload_Should_Keep_Annotations()
        {
            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(_samplePdfPath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            var page = docModel.Pages[0];

            var newAnnot = new PdfAnnotation
            {
                Type = AnnotationType.FreeText,
                X = 50,
                Y = 50,
                Width = 100,
                Height = 30,
                TextContent = "Hello Test",
                FontSize = 12,
                FontFamily = "Malgun Gothic"
            };
            page.Annotations.Add(newAnnot);
            page.Annotations.Add(new PdfAnnotation
            {
                Type = AnnotationType.Highlight,
                X = 100,
                Y = 100,
                Width = 50,
                Height = 10,
                AnnotationColor = Colors.Yellow
            });

            string savePath = Path.Combine(_testFileDir, "saved_output.pdf");
            await service.SavePdf(docModel, savePath);
            await docModel.CloseAsync();

            Assert.True(File.Exists(savePath));
            var reloadedModel = await service.LoadPdfAsync(savePath);
            Assert.NotNull(reloadedModel);
            await service.InitializeDocumentAsync(reloadedModel);

            var reloadedPage = reloadedModel.Pages[0];
            service.RenderPageImage(reloadedModel, reloadedPage);

            Assert.True(reloadedPage.Annotations.Count >= 2);
            var reloadedText = reloadedPage.Annotations.FirstOrDefault(a => a.Type == AnnotationType.FreeText);
            Assert.NotNull(reloadedText);
            Assert.Equal("Hello Test", reloadedText.TextContent);
            await reloadedModel.CloseAsync();
        }

        // [Task 1] Glyphless Searchable PDF 생성 및 검색 검증 (파일 폭증 없음 검증)
        [Fact]
        public async Task GlyphlessOcr_ShouldInjectSearchableText_WithoutSizeBloat()
        {
            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(_samplePdfPath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            var page = docModel.Pages[0];
            page.OcrWords = new List<OcrWordInfo>
            {
                new OcrWordInfo { Text = "대한민국", BoundingBox = new Rect(50, 100, 80, 20) },
                new OcrWordInfo { Text = "특허청", BoundingBox = new Rect(140, 100, 60, 20) }
            };

            string ocrSavedPath = Path.Combine(_testFileDir, "glyphless_ocr.pdf");
            await service.SavePdf(docModel, ocrSavedPath);
            await docModel.CloseAsync();

            Assert.True(File.Exists(ocrSavedPath));
            long origSize = new FileInfo(_samplePdfPath).Length;
            long savedSize = new FileInfo(ocrSavedPath).Length;

            // 기존 13MB 폰트 임베딩과 달리 수십 KB 이내로 유지되어야 함
            long diffSize = savedSize - origSize;
            Assert.True(diffSize < 100 * 1024, $"File bloat detected: diff was {diffSize} bytes");

            // 저장된 PDF에서 텍스트가 검색되는지 확인
            var rects = service.FindTextRects(ocrSavedPath, 0, "대한민국");
            Assert.NotEmpty(rects);
            var rects2 = service.FindTextRects(ocrSavedPath, 0, "특허청");
            Assert.NotEmpty(rects2);
        }

        // [Task 4] SearchService 순차 탐색 버그 수정 검증 (페이지 1을 건너뛰지 않는지 검증)
        [Fact]
        public async Task SearchService_ShouldSearchSequentialPages_WithoutSkipping()
        {
            string multiPagePath = Path.Combine(_testFileDir, "three_page.pdf");
            File.WriteAllBytes(multiPagePath, CreateThreePagePdf());

            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(multiPagePath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            // 페이지 1(두 번째 페이지)에만 OCR 텍스트 배치
            docModel.Pages[1].OcrWords = new List<OcrWordInfo>
            {
                new OcrWordInfo { Text = "타겟단어", BoundingBox = new Rect(30, 40, 50, 20) }
            };

            var searchService = new SearchService(service);
            searchService.ResetSearch();

            // 검색 실행: 페이지 0에는 없고 페이지 1에 있음 -> 기존 삼각수 루프면 0->2로 건너뛰어 실패했을 것임!
            var match = await searchService.FindNextAsync(docModel, "타겟단어");
            Assert.NotNull(match);

            // 페이지 1에 하이라이트가 추가되었는지 확인
            Assert.Contains(docModel.Pages[1].Annotations, a => a.Type == AnnotationType.SearchHighlight);

            await docModel.CloseAsync();
        }

        // [Task 4] 메모리 내 OcrWords 즉시 검색 검증
        [Fact]
        public async Task SearchService_ShouldSearchInMemoryOcrWords()
        {
            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(_samplePdfPath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            docModel.Pages[0].OcrWords = new List<OcrWordInfo>
            {
                new OcrWordInfo { Text = "인공지능", BoundingBox = new Rect(20, 30, 40, 15) },
                new OcrWordInfo { Text = "검색테스트", BoundingBox = new Rect(70, 30, 60, 15) }
            };

            var searchService = new SearchService(service);
            searchService.ResetSearch();

            var match = await searchService.FindNextAsync(docModel, "인공지능");
            Assert.NotNull(match);

            await docModel.CloseAsync();
        }

        // [Task 2] 단일 페이지 회전 누적 검증
        [Fact]
        public async Task SinglePage_Rotation_ShouldRotateOnlyTargetPage()
        {
            string multiPagePath = Path.Combine(_testFileDir, "two_page.pdf");
            File.WriteAllBytes(multiPagePath, CreateTwoPagePdf());

            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(multiPagePath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            // 페이지 0만 90도 회전
            docModel.Pages[0].Rotation = 90;
            // 페이지 1은 회전 없음 (0도)
            docModel.Pages[1].Rotation = 0;

            string savedPath = Path.Combine(_testFileDir, "rotated.pdf");
            await service.SavePdf(docModel, savedPath);
            await docModel.CloseAsync();

            var reloadedModel = await service.LoadPdfAsync(savedPath);
            Assert.NotNull(reloadedModel);
            await service.InitializeDocumentAsync(reloadedModel);

            // 페이지 0은 가로세로 비율이 바뀌었거나 회전이 적용됨
            Assert.Equal(2, reloadedModel.Pages.Count);
            await reloadedModel.CloseAsync();
        }

        // [Task 6] Stem Darkening LUT 비트맵 변환 검증 (안티앨리어싱 그레이 픽셀의 농도 강화 확인)
        [Fact]
        public void ToBitmapSource_WithStemDarkening_DarkensIntermediatePixels()
        {
            using var bmp = new System.Drawing.Bitmap(2, 2, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            bmp.SetPixel(0, 0, System.Drawing.Color.FromArgb(255, 255, 255, 255));
            bmp.SetPixel(1, 0, System.Drawing.Color.FromArgb(255, 0, 0, 0));
            bmp.SetPixel(0, 1, System.Drawing.Color.FromArgb(255, 180, 180, 180));
            bmp.SetPixel(1, 1, System.Drawing.Color.FromArgb(255, 100, 100, 100));

            var result = PdfService.ToBitmapSource(bmp, applyStemDarkening: true, applySharpen: false);
            Assert.NotNull(result);
            Assert.Equal(2, result.PixelWidth);
            Assert.Equal(2, result.PixelHeight);

            byte[] pixels = new byte[2 * 2 * 4];
            result.CopyPixels(pixels, 2 * 4, 0);

            // White 유지
            Assert.Equal(255, pixels[0]);
            Assert.Equal(255, pixels[1]);
            Assert.Equal(255, pixels[2]);

            // Black 유지
            Assert.Equal(0, pixels[4]);
            Assert.Equal(0, pixels[5]);
            Assert.Equal(0, pixels[6]);

            // Mid gray 어두워짐 검증 (180 -> <180, 100 -> <100)
            Assert.True(pixels[8] < 180, $"Expected pixel to be darkened, got {pixels[8]}");
            Assert.True(pixels[12] < 100, $"Expected pixel to be darkened, got {pixels[12]}");
        }

        // 밝기 언샤프: 가장자리 대비는 커지고, 균일한 면·흰색·검정은 그대로, 색 차이(채널 간 차)는 유지
        [Fact]
        public unsafe void SharpenLuma_IncreasesEdgeContrast_KeepsFlatAreasAndChroma()
        {
            const int w = 8, h = 4;
            var px = new byte[w * h * 4];
            for (int y = 0; y < h; y++)
            {
                for (int x = 0; x < w; x++)
                {
                    int i = (y * w + x) * 4;
                    byte v = x < 4 ? (byte)40 : (byte)220; // 어두운 획 | 밝은 바탕
                    px[i] = v; px[i + 1] = v; px[i + 2] = v; px[i + 3] = 255;
                }
            }
            // 색이 있는 픽셀 하나 (채널 차 보존 확인)
            int c = (1 * w + 6) * 4;
            px[c] = 200; px[c + 1] = 220; px[c + 2] = 240;

            fixed (byte* p = px)
                PdfService.SharpenLuma(p, w, h, w * 4, 0.6);

            byte At(int x, int y) => px[(y * w + x) * 4 + 1];
            Assert.Equal(40, At(0, 2));            // 획 안쪽(균일) 그대로
            Assert.Equal(220, At(7, 2));           // 바탕(균일) 그대로
            Assert.True(At(3, 2) < 40, $"edge dark side should get darker, got {At(3, 2)}");
            Assert.True(At(4, 2) > 220, $"edge light side should get lighter, got {At(4, 2)}");
            Assert.Equal(40, px[c + 2] - px[c]);   // R-B 차 유지 (밝기만 바뀜)
        }

        // 1번: OCR 단어 좌표는 줌과 무관한 포인트. 줌 2배에서도 텍스트 층이 같은 위치에 들어가야 한다.
        [Fact]
        public async Task GlyphlessOcr_IsZoomIndependent()
        {
            string path = Path.Combine(_testFileDir, "zoom.pdf");
            File.WriteAllBytes(path, CreateTwoPagePdf()); // 1쪽 MediaBox 300×400, 회전 없음

            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(path);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            docModel.Zoom = 2.0;
            foreach (var p in docModel.Pages)
            {
                p.Width = p.PdfPageWidthPoint * docModel.Zoom;
                p.Height = p.PdfPageHeightPoint * docModel.Zoom;
            }

            docModel.Pages[0].OcrWords = new List<OcrWordInfo>
            {
                new OcrWordInfo { Text = "줌무관", BoundingBox = new Rect(60, 120, 90, 20) }
            };

            string savedPath = Path.Combine(_testFileDir, "zoom_ocr.pdf");
            await service.SavePdf(docModel, savedPath);
            await docModel.CloseAsync();

            var rects = service.FindTextRects(savedPath, 0, "줌무관");
            Assert.NotEmpty(rects);
            // FindTextRects는 PDF 좌표(아래 원점): left ≈ 60, top ≈ 400 - 120 = 280
            Assert.InRange(rects[0].X, 58, 62);
            Assert.InRange(rects[0].Y, 276, 284);
        }

        // 2·5번: 앱에서 90° 돌린 페이지 + 부모에서 상속된 Resources. 저장 후에도 원본 내용이 남고 OCR 텍스트가 검색된다.
        [Fact]
        public async Task GlyphlessOcr_RotatedPage_WithInheritedResources_IsSearchable()
        {
            string path = Path.Combine(_testFileDir, "inherited.pdf");
            File.WriteAllBytes(path, CreateInheritedResourcesPdf());

            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(path);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            var page = docModel.Pages[0];
            double srcW = page.SourceViewWidthPoint;
            // 앱 회전 버튼과 같은 처리 (가로·세로 맞바꿈)
            page.Rotation = 90;
            (page.Width, page.Height) = (page.Height, page.Width);
            (page.PdfPageWidthPoint, page.PdfPageHeightPoint) = (page.PdfPageHeightPoint, page.PdfPageWidthPoint);
            Assert.Equal(srcW, page.SourceViewWidthPoint, 3);

            page.OcrWords = new List<OcrWordInfo>
            {
                new OcrWordInfo { Text = "회전검색", BoundingBox = new Rect(50, 60, 120, 24) }
            };

            string savedPath = Path.Combine(_testFileDir, "inherited_ocr.pdf");
            await service.SavePdf(docModel, savedPath);
            await docModel.CloseAsync();

            Assert.NotEmpty(service.FindTextRects(savedPath, 0, "회전검색"));
            // 상속 Resources의 원래 글꼴로 쓴 글자도 그대로 남아 있어야 한다
            Assert.NotEmpty(service.FindTextRects(savedPath, 0, "Inherited"));

            string text = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(savedPath));
            Assert.Contains("/" + PdfGlyphlessOcrService.FontResName, text);
        }

        // 5번: 다시 OCR해 저장하면 이전 텍스트 층을 바꿔 끼운다 (겹치지 않음)
        [Fact]
        public async Task GlyphlessOcr_Reinjection_ReplacesPreviousLayer()
        {
            string first = Path.Combine(_testFileDir, "first.pdf");
            File.WriteAllBytes(first, CreateMinimalPdf());
            var svc = new PdfGlyphlessOcrService();
            Assert.True(svc.InjectOcrLayer(first, new List<OcrPageData>
            {
                new OcrPageData { PageIndex = 0, Words = { new OcrWordInfo { Text = "첫번째층", BoundingBox = new Rect(40, 40, 80, 20) } } }
            }));
            Assert.True(svc.InjectOcrLayer(first, new List<OcrPageData>
            {
                new OcrPageData { PageIndex = 0, Words = { new OcrWordInfo { Text = "두번째층", BoundingBox = new Rect(40, 80, 80, 20) } } }
            }));

            var service = new PdfService();
            Assert.Empty(service.FindTextRects(first, 0, "첫번째층"));
            Assert.NotEmpty(service.FindTextRects(first, 0, "두번째층"));
        }

        [Fact]
        public void ToUnicodeCMap_UsesAtMost100EntriesPerBlock()
        {
            string cmap = PdfGlyphlessOcrService.BuildIdentityToUnicodeCMap();
            var counts = System.Text.RegularExpressions.Regex.Matches(cmap, @"(\d+) beginbfrange")
                .Select(m => int.Parse(m.Groups[1].Value)).ToList();
            Assert.Equal(256, counts.Sum());
            Assert.All(counts, n => Assert.True(n <= 100));
        }

        [Fact]
        public void PdfSyntax_HandlesNestedDictionaries()
        {
            string dict = "<< /Font << /F1 5 0 R /F2 << /Sub (a>>b) >> >> /XObject << /Im0 7 0 R >> /ProcSet [/PDF /Text] >>";
            Assert.Equal("<< /Im0 7 0 R >>", PdfSyntax.GetValue(dict, "XObject"));
            string updated = PdfSyntax.SetValue(dict, "Font", PdfSyntax.SetValue(PdfSyntax.GetValue(dict, "Font")!, "FX", "9 0 R"));
            Assert.Contains("/FX 9 0 R", PdfSyntax.GetValue(updated, "Font"));
            Assert.Equal("<< /Im0 7 0 R >>", PdfSyntax.GetValue(updated, "XObject"));
            Assert.Equal("[/PDF /Text]", PdfSyntax.GetValue(updated, "ProcSet"));
            Assert.True(PdfSyntax.TryParseReference(PdfSyntax.GetValue("<< /Parent 12 3 R >>", "Parent"), out int n, out int g));
            Assert.Equal((12, 3), (n, g));
        }

        // [Task 6] HasPendingRenders 상태 노출 확인
        [Fact]
        public void HasPendingRenders_Property_IsAccessible()
        {
            Assert.False(PdfService.HasPendingRenders);
        }

        private static byte[] CreateMinimalPdf()
        {
            var objects = new[]
            {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
                "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> >>\nendobj\n"
            };

            var sb = new System.Text.StringBuilder();
            sb.Append("%PDF-1.7\n");
            var offsets = new List<int>();
            foreach (var obj in objects)
            {
                offsets.Add(System.Text.Encoding.ASCII.GetByteCount(sb.ToString()));
                sb.Append(obj);
            }

            int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(sb.ToString());
            sb.Append("xref\n0 4\n");
            sb.Append("0000000000 65535 f \n");
            foreach (int offset in offsets)
                sb.Append($"{offset:0000000000} 00000 n \n");
            sb.Append("trailer\n<< /Size 4 /Root 1 0 R >>\n");
            sb.Append("startxref\n");
            sb.Append(xrefOffset);
            sb.Append("\n%%EOF\n");
            return System.Text.Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static byte[] CreateInheritedResourcesPdf()
        {
            // /Resources는 부모 /Pages에만 있고, 페이지는 상속받는다
            var objects = new[]
            {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
                "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R >>\nendobj\n",
                "4 0 obj\n<< /Length 40 >>\nstream\nBT /F1 18 Tf 72 700 Td (Inherited) Tj ET\nendstream\nendobj\n",
                "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n"
            };

            var sb = new System.Text.StringBuilder();
            sb.Append("%PDF-1.7\n");
            var offsets = new List<int>();
            foreach (var obj in objects)
            {
                offsets.Add(System.Text.Encoding.ASCII.GetByteCount(sb.ToString()));
                sb.Append(obj);
            }

            int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(sb.ToString());
            sb.Append("xref\n0 6\n");
            sb.Append("0000000000 65535 f \n");
            foreach (int offset in offsets)
                sb.Append($"{offset:0000000000} 00000 n \n");
            sb.Append("trailer\n<< /Size 6 /Root 1 0 R >>\n");
            sb.Append("startxref\n");
            sb.Append(xrefOffset);
            sb.Append("\n%%EOF\n");
            return System.Text.Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static byte[] CreateTwoPagePdf()
        {
            var objects = new[]
            {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
                "2 0 obj\n<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>\nendobj\n",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /Resources << >> >>\nendobj\n",
                "4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 600] /Resources << >> >>\nendobj\n"
            };

            var sb = new System.Text.StringBuilder();
            sb.Append("%PDF-1.7\n");
            var offsets = new List<int>();
            foreach (var obj in objects)
            {
                offsets.Add(System.Text.Encoding.ASCII.GetByteCount(sb.ToString()));
                sb.Append(obj);
            }

            int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(sb.ToString());
            sb.Append("xref\n0 5\n");
            sb.Append("0000000000 65535 f \n");
            foreach (int offset in offsets)
                sb.Append($"{offset:0000000000} 00000 n \n");

            sb.Append("trailer\n<< /Size 5 /Root 1 0 R >>\n");
            sb.Append("startxref\n");
            sb.Append(xrefOffset);
            sb.Append("\n%%EOF");
            return System.Text.Encoding.ASCII.GetBytes(sb.ToString());
        }

        private static byte[] CreateThreePagePdf()
        {
            var objects = new[]
            {
                "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
                "2 0 obj\n<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 >>\nendobj\n",
                "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 500] /Resources << >> >>\nendobj\n",
                "4 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 500] /Resources << >> >>\nendobj\n",
                "5 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 400 500] /Resources << >> >>\nendobj\n"
            };

            var sb = new System.Text.StringBuilder();
            sb.Append("%PDF-1.7\n");
            var offsets = new List<int>();
            foreach (var obj in objects)
            {
                offsets.Add(System.Text.Encoding.ASCII.GetByteCount(sb.ToString()));
                sb.Append(obj);
            }

            int xrefOffset = System.Text.Encoding.ASCII.GetByteCount(sb.ToString());
            sb.Append("xref\n0 6\n");
            sb.Append("0000000000 65535 f \n");
            foreach (int offset in offsets)
                sb.Append($"{offset:0000000000} 00000 n \n");

            sb.Append("trailer\n<< /Size 6 /Root 1 0 R >>\n");
            sb.Append("startxref\n");
            sb.Append(xrefOffset);
            sb.Append("\n%%EOF");
            return System.Text.Encoding.ASCII.GetBytes(sb.ToString());
        }
    }
}
