using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using MinsPDFViewer;
using Xunit;

namespace MinsPDFViewer.Tests
{
    /// <summary>
    /// 자체 PDFium 래퍼(bblanchon.PDFium) 검증: 한글 경로, 타일 렌더 = 전체 렌더의 잘라낸 부분, 회전 방향,
    /// 점진 렌더 취소, 검색 좌표가 실제 글자 위에 오는지, OCR 단어 좌표 → 화면 좌표.
    /// </summary>
    public class PdfiumEngineTests : IDisposable
    {
        private readonly string _dir;

        public PdfiumEngineTests()
        {
            PdfService.UseDirectDispatcherForTests = true;
            _dir = Path.Combine(Path.GetTempPath(), "MinsPdfiumTests_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_dir, true);
            }
            catch
            {
            }
        }

        private string CopyFixture(string name, string? targetName = null)
        {
            string? testFiles = PdfServiceTests.FindTestFilesDirectory();
            Assert.NotNull(testFiles);
            string target = Path.Combine(_dir, targetName ?? name);
            File.Copy(Path.Combine(testFiles!, name), target, true);
            return target;
        }

        [Fact]
        public void Open_KoreanPath_TileEqualsCropOfFullRender()
        {
            string path = CopyFixture("normal.pdf", "한글 경로 문서.pdf");
            using var doc = PdfiumDocument.Open(path, null);
            Assert.True(doc.PageCount > 0);

            lock (Pdfium.Lock)
            {
                var (w, h) = doc.GetPageSize(0);
                int fw = (int)(w * 2), fh = (int)(h * 2);
                int flags = PdfiumNative.FPDF_ANNOT | PdfiumNative.FPDF_LCD_TEXT;
                byte[] full = doc.RenderToBuffer(0, fw, fh, 0, flags, PixelRegion.Full(fw, fh))!;
                var region = new PixelRegion(fw / 3, fh / 4, 200, 150);
                byte[] tile = doc.RenderToBuffer(0, fw, fh, 0, flags, region)!;
                Assert.NotNull(full);
                Assert.NotNull(tile);

                for (int y = 0; y < region.Height; y++)
                {
                    for (int x = 0; x < region.Width * 4; x++)
                    {
                        byte expected = full[((region.Y + y) * fw + region.X) * 4 + x];
                        Assert.Equal(expected, tile[y * region.Width * 4 + x]);
                    }
                }

                // 점진 렌더(취소 안 함)도 같은 결과
                byte[] progressive = doc.RenderToBuffer(0, fw, fh, 0, flags, PixelRegion.Full(fw, fh), () => false)!;
                Assert.True(progressive.SequenceEqual(full));
            }
        }

        [Fact]
        public void Rotate1_IsClockwiseQuarterTurn()
        {
            string path = CopyFixture("normal.pdf");
            using var doc = PdfiumDocument.Open(path, null);
            lock (Pdfium.Lock)
            {
                var (w, h) = doc.GetPageSize(0);
                int fw = (int)w, fh = (int)h;
                byte[] upright = doc.RenderToBuffer(0, fw, fh, 0, PdfiumNative.FPDF_ANNOT, PixelRegion.Full(fw, fh))!;
                byte[] turned = doc.RenderToBuffer(0, fh, fw, 1, PdfiumNative.FPDF_ANNOT, PixelRegion.Full(fh, fw))!;
                double diff = 0;
                int samples = 0;
                for (int y = 0; y < fh; y += 17)
                {
                    for (int x = 0; x < fw; x += 17)
                    {
                        // 시계 방향 90°: (x, y) → (fh - 1 - y, x)
                        diff += Math.Abs(upright[(y * fw + x) * 4] - turned[(x * fh + (fh - 1 - y)) * 4]);
                        samples++;
                    }
                }
                Assert.True(diff / samples < 8, $"average difference {diff / samples:0.0}");
            }
        }

        [Fact]
        public void ProgressiveRender_Cancel_ReturnsNull()
        {
            string path = CopyFixture("normal.pdf");
            using var doc = PdfiumDocument.Open(path, null);
            lock (Pdfium.Lock)
            {
                var (w, h) = doc.GetPageSize(0);
                int fw = (int)(w * 4), fh = (int)(h * 4);
                Assert.Null(doc.RenderToBuffer(0, fw, fh, 0, PdfiumNative.FPDF_ANNOT, PixelRegion.Full(fw, fh), () => true));
                // 취소 뒤에도 같은 페이지를 정상으로 그릴 수 있어야 한다
                Assert.NotNull(doc.RenderToBuffer(0, fw / 4, fh / 4, 0, PdfiumNative.FPDF_ANNOT, PixelRegion.Full(fw / 4, fh / 4)));
            }
        }

        [Fact]
        public async Task FindTextRectsInView_LandsOnInk_AtZoomAndRotation()
        {
            string path = CopyFixture("print_restricted.pdf");
            var service = new PdfService();
            var model = await service.LoadPdfAsync(path);
            Assert.NotNull(model);
            await service.InitializeDocumentAsync(model!);
            var page = model!.Pages[0];

            foreach (int rotation in new[] { 0, 90 })
            {
                if (rotation == 90)
                {
                    page.Rotation = 90;
                    (page.PdfPageWidthPoint, page.PdfPageHeightPoint) = (page.PdfPageHeightPoint, page.PdfPageWidthPoint);
                }
                page.Width = page.PdfPageWidthPoint * 1.5;
                page.Height = page.PdfPageHeightPoint * 1.5;

                var rects = service.FindTextRectsInView(model, page, "restricted");
                Assert.NotEmpty(rects);
                Rect r = rects[0];

                int pw = (int)Math.Round(page.Width), ph = (int)Math.Round(page.Height);
                byte[] image;
                lock (Pdfium.Lock)
                {
                    image = model.Document!.RenderToBuffer(page.OriginalPageIndex, pw, ph, PdfService.ToPdfiumRotate(page.Rotation),
                        PdfiumNative.FPDF_ANNOT, PixelRegion.Full(pw, ph))!;
                }
                int dark = 0, total = 0;
                for (int y = (int)r.Top; y < (int)r.Bottom; y++)
                {
                    for (int x = (int)r.Left; x < (int)r.Right; x++)
                    {
                        if (x < 0 || y < 0 || x >= pw || y >= ph)
                            continue;
                        total++;
                        if (image[(y * pw + x) * 4] < 128)
                            dark++;
                    }
                }
                Assert.True(total > 0 && dark > total / 20, $"rotation {rotation}: rect {r} dark {dark}/{total}");

                // 같은 영역의 글자 추출
                string text = service.ExtractTextInRect(model, page, new Rect(r.X - 2, r.Y - 2, r.Width + 4, r.Height + 4));
                Assert.Contains("restricted", text);
            }

            await model.CloseAsync();
        }

        [Fact]
        public void OcrRectToView_FollowsAppRotation()
        {
            // 원본 화면 600×800pt, 앱에서 시계 방향 90° 회전, 줌 1
            var page = new PdfPageViewModel
            {
                Rotation = 90,
                PdfPageWidthPoint = 800,
                PdfPageHeightPoint = 600,
                Width = 800,
                Height = 600
            };
            var r = SearchService.OcrRectToView(page, new Rect(10, 20, 30, 40));
            Assert.Equal(740, r.X, 3);
            Assert.Equal(10, r.Y, 3);
            Assert.Equal(40, r.Width, 3);
            Assert.Equal(30, r.Height, 3);

            // 회전 없음, 줌 2
            var flat = new PdfPageViewModel { PdfPageWidthPoint = 600, PdfPageHeightPoint = 800, Width = 1200, Height = 1600 };
            var f = SearchService.OcrRectToView(flat, new Rect(10, 20, 30, 40));
            Assert.Equal(new Rect(20, 40, 60, 80), f);
        }
    }
}
