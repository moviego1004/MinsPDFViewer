using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using MinsPDFViewer;
using Xunit;

namespace MinsPDFViewer.Tests
{
    public class PrintServiceTests : IDisposable
    {
        private readonly string _dir;

        public PrintServiceTests()
        {
            PdfService.UseDirectDispatcherForTests = true;
            _dir = Path.Combine(Path.GetTempPath(), "MinsPrintTests_" + Guid.NewGuid().ToString("N"));
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

        private string CopyFixture(string name)
        {
            string? testFiles = PdfServiceTests.FindTestFilesDirectory();
            Assert.NotNull(testFiles);
            string target = Path.Combine(_dir, name);
            File.Copy(Path.Combine(testFiles!, name), target, true);
            return target;
        }

        private static async Task<(PdfService Service, PdfDocumentModel Model)> OpenAsync(string path)
        {
            var service = new PdfService();
            var model = await service.LoadPdfAsync(path);
            Assert.NotNull(model);
            await service.InitializeDocumentAsync(model!);
            return (service, model!);
        }

        // ── 범위·배치 ──

        [Theory]
        [InlineData("1-3, 5, 8-", 10, "1,2,3,5,8,9,10")]
        [InlineData("-4", 10, "1,2,3,4")]
        [InlineData("5-3", 10, "5,4,3")]
        [InlineData("2", 3, "2")]
        [InlineData("1~2;3", 5, "1,2,3")]
        public void PageRange_Parses(string text, int count, string expected)
        {
            Assert.True(PrintPageRange.TryParse(text, count, out var pages, out _));
            Assert.Equal(expected, string.Join(",", pages.Select(p => p + 1)));
        }

        [Theory]
        [InlineData("0", 3)]
        [InlineData("4", 3)]
        [InlineData("a", 3)]
        [InlineData("", 3)]
        [InlineData("2-9", 3)]
        public void PageRange_RejectsInvalid(string text, int count)
        {
            Assert.False(PrintPageRange.TryParse(text, count, out _, out string? error));
            Assert.False(string.IsNullOrEmpty(error));
        }

        [Fact]
        public void Layout_FitShrinkActual()
        {
            var printable = new PrintRect(25, 25, 800, 1050); // Letter, 0.25" 여백
            // A4(595×842pt) 맞춤: 높이가 꽉 차고 가로 가운데
            var fit = PrintLayout.Compute(595, 842, 850, 1100, printable, PrintScaleMode.Fit);
            Assert.Equal(1050, fit.Height, 3);
            Assert.Equal(425, fit.X + fit.Width / 2, 3);

            // 작은 페이지(300×400pt)는 '큰 페이지만 축소'에서 100%
            var small = PrintLayout.Compute(300, 400, 850, 1100, printable, PrintScaleMode.ShrinkOversized);
            Assert.Equal(300 * PrintLayout.HundredthsPerPoint, small.Width, 3);

            // 실제 크기는 용지 가운데, 넘치면 음수 좌표(잘림)
            var actual = PrintLayout.Compute(595, 842, 850, 1100, printable, PrintScaleMode.ActualSize);
            Assert.Equal(595 * PrintLayout.HundredthsPerPoint, actual.Width, 3);
            Assert.True(actual.Y < 0);
        }

        [Fact]
        public void Layout_AutoOrientation()
        {
            Assert.True(PrintLayout.IsLandscape(842, 595, PrintOrientationMode.Auto));
            Assert.False(PrintLayout.IsLandscape(595, 842, PrintOrientationMode.Auto));
            Assert.True(PrintLayout.IsLandscape(595, 842, PrintOrientationMode.Landscape));
            Assert.False(PrintLayout.IsLandscape(842, 595, PrintOrientationMode.Portrait));
        }

        [Theory]
        [InlineData(842, 595, 850, 1100, PrintOrientationMode.Auto, true)]   // 가로 페이지인데 세로 용지가 옴 → 돌림
        [InlineData(842, 595, 1100, 850, PrintOrientationMode.Auto, false)]  // 가로 용지가 옴 → 그대로
        [InlineData(595, 842, 850, 1100, PrintOrientationMode.Auto, false)]
        [InlineData(595, 842, 850, 1100, PrintOrientationMode.Landscape, true)]
        [InlineData(600, 600, 850, 1100, PrintOrientationMode.Auto, false)]  // 정사각형 페이지는 돌리지 않음
        public void ContentRotation_WhenDriverIgnoresOrientation(double w, double h, double paperW, double paperH, PrintOrientationMode mode, bool expected)
        {
            Assert.Equal(expected, PrintLayout.NeedsContentRotation(w, h, paperW, paperH, mode));
        }

        // ── 현재 상태 판단 ──

        [Fact]
        public async Task PendingChanges_DetectsEditsAndIgnoresZoom()
        {
            string path = CopyFixture("normal.pdf");
            var (service, model) = await OpenAsync(path);
            int sourceCount = model.Pages.Count;

            // 주석을 파일에서 읽은 상태 = 변경 없음
            foreach (var page in model.Pages)
                service.RenderPageImage(model, page);
            Assert.False(PrintService.HasPendingChanges(model, sourceCount, out string reason), reason);

            // 줌만 바꾸면(주석 좌표가 같이 늘어남) 여전히 변경 없음
            double scale = 1.5;
            foreach (var page in model.Pages)
            {
                page.Width *= scale;
                page.Height *= scale;
                foreach (var a in page.Annotations)
                {
                    a.X *= scale;
                    a.Y *= scale;
                    a.Width *= scale;
                    a.Height *= scale;
                }
            }
            Assert.False(PrintService.HasPendingChanges(model, sourceCount, out reason), reason);

            // 형광펜 추가 → 변경
            model.Pages[0].Annotations.Add(new PdfAnnotation { Type = AnnotationType.Highlight, X = 10, Y = 10, Width = 40, Height = 10, Background = Brushes.Yellow });
            Assert.True(PrintService.HasPendingChanges(model, sourceCount, out _));
            model.Pages[0].Annotations.RemoveAt(model.Pages[0].Annotations.Count - 1);
            Assert.False(PrintService.HasPendingChanges(model, sourceCount, out reason), reason);

            // 검색 하이라이트는 인쇄와 무관
            model.Pages[0].Annotations.Add(new PdfAnnotation { Type = AnnotationType.SearchHighlight, X = 1, Y = 1, Width = 5, Height = 5 });
            Assert.False(PrintService.HasPendingChanges(model, sourceCount, out reason), reason);

            // 회전 → 변경
            model.Pages[0].Rotation = 90;
            Assert.True(PrintService.HasPendingChanges(model, sourceCount, out _));
            model.Pages[0].Rotation = 0;

            // 순서 바꾸기 → 변경
            if (model.Pages.Count > 1)
            {
                model.Pages.Move(0, 1);
                Assert.True(PrintService.HasPendingChanges(model, sourceCount, out _));
                model.Pages.Move(1, 0);
            }

            await model.CloseAsync();
        }

        [Fact]
        public async Task Prepare_WithoutEdits_UsesOriginal_WithEdits_ReflectsThem()
        {
            string path = CopyFixture("print_restricted.pdf");
            var (service, model) = await OpenAsync(path);
            var printService = new PrintService(service);

            using (var source = await printService.PrepareAsync(model))
            {
                Assert.False(source.FromPendingEdits);
                Assert.Equal(2, source.PageCount);
            }

            // 1쪽(세로)을 앱에서 90° 돌리고 텍스트 주석 추가 → 임시 저장본으로 인쇄
            var page = model.Pages[0];
            page.Rotation = 90;
            (page.Width, page.Height) = (page.Height, page.Width);
            (page.PdfPageWidthPoint, page.PdfPageHeightPoint) = (page.PdfPageHeightPoint, page.PdfPageWidthPoint);
            model.Pages[1].Annotations.Add(new PdfAnnotation { Type = AnnotationType.FreeText, X = 50, Y = 50, Width = 200, Height = 40, TextContent = "인쇄메모", FontSize = 14 });

            using (var source = await printService.PrepareAsync(model))
            {
                Assert.True(source.FromPendingEdits);
                Assert.Equal(2, source.PageCount);
                Assert.True(source.PageSizes[0].Width > source.PageSizes[0].Height, "rotated page should print landscape");
            }

            // 원본 파일은 그대로
            Assert.Equal(File.ReadAllBytes(PdfServiceTestsPath("print_restricted.pdf")), File.ReadAllBytes(path));
            await model.CloseAsync();
        }

        // ── 인쇄 제한 무시 ──

        [Fact]
        public async Task PrintRestrictedDocument_IsDetected_AndStillRenders()
        {
            string path = CopyFixture("print_restricted.pdf");
            var (service, model) = await OpenAsync(path);
            var printService = new PrintService(service);

            using var source = await printService.PrepareAsync(model);
            Assert.True(source.PrintRestricted);

            // 인쇄 경로(ForPrinting)로 그려진다 = 제한을 무시
            var bitmap = PrintService.RenderPreviewPage(source, 0, 200, 283, includeAnnotations: true, grayscale: false);
            Assert.NotNull(bitmap);
            Assert.Equal(200, bitmap!.PixelWidth);

            await model.CloseAsync();
        }

        [Fact]
        public async Task UnrestrictedDocument_IsNotFlagged()
        {
            string path = CopyFixture("normal.pdf");
            var (service, model) = await OpenAsync(path);
            using var source = await new PrintService(service).PrepareAsync(model);
            Assert.False(source.PrintRestricted);
            await model.CloseAsync();
        }

        private static string PdfServiceTestsPath(string name) => Path.Combine(PdfServiceTests.FindTestFilesDirectory()!, name);
    }
}
