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
