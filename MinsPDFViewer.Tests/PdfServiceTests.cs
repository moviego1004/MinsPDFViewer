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
    public class PdfServiceTests
    {
        private readonly string _testFileDir;
        private readonly string _samplePdfPath;
        private readonly string _signedPdfPath;

        public PdfServiceTests()
        {
            // 테스트 파일들이 저장될 임시 폴더
            _testFileDir = Path.Combine(AppContext.BaseDirectory, "TestFiles");
            if (!Directory.Exists(_testFileDir))
                Directory.CreateDirectory(_testFileDir);

            _samplePdfPath = Path.Combine(_testFileDir, "normal.pdf");
            _signedPdfPath = Path.Combine(_testFileDir, "signed.pdf");

            string repoSamplePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestFiles", "normal.pdf"));
            if (File.Exists(repoSamplePath))
                File.Copy(repoSamplePath, _samplePdfPath, overwrite: true);
            else
                File.WriteAllBytes(_samplePdfPath, CreateMinimalPdf());
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
                X = 70,
                Y = 120,
                Width = 140,
                Height = 18,
                AnnotationColor = Colors.Lime,
                Background = new SolidColorBrush(Color.FromArgb(80, Colors.Lime.R, Colors.Lime.G, Colors.Lime.B))
            });
            page.Annotations.Add(new PdfAnnotation
            {
                Type = AnnotationType.Underline,
                X = 70,
                Y = 165,
                Width = 140,
                Height = 2,
                AnnotationColor = Colors.Black,
                Background = Brushes.Black
            });

            string savePath = Path.Combine(_testFileDir, "tc11_result.pdf");

            await service.SavePdf(docModel, savePath);

            string savedText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(savePath));
            Assert.Contains("MINS_FREETEXT_V2:", savedText);

            var reloaded = await service.LoadPdfAsync(savePath);
            Assert.NotNull(reloaded);
            await service.InitializeDocumentAsync(reloaded);
            var reloadedPage = reloaded.Pages[0];
            service.RenderPageImage(reloaded, reloadedPage);

            Assert.Contains(reloadedPage.Annotations, a =>
                a.Type == AnnotationType.FreeText &&
                a.TextContent == "Hello Test" &&
                Math.Abs(a.X - 50) < 1 &&
                Math.Abs(a.Y - 50) < 1 &&
                Math.Abs(a.Width - 100) < 1 &&
                Math.Abs(a.Height - 30) < 1);
            Assert.Contains(reloadedPage.Annotations, a => a.Type == AnnotationType.Highlight);
            Assert.Contains(reloadedPage.Annotations, a => a.Type == AnnotationType.Underline);
        }

        // [TC-18] 서명된 파일 저장 시 서명 데이터 보존 확인
        [Fact]
        public async Task Save_Signed_Pdf_Should_Preserve_Signature()
        {
            if (!File.Exists(_signedPdfPath))
            {
                return; // 파일 없으면 패스
            }

            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(_signedPdfPath);
            Assert.NotNull(docModel);

            string savePath = Path.Combine(_testFileDir, "tc18_result_signed.pdf");

            await service.SavePdf(docModel, savePath);

            string savedText = System.Text.Encoding.Latin1.GetString(File.ReadAllBytes(savePath));
            bool isStillSigned = savedText.Contains("/Subtype /Widget") &&
                                 savedText.Contains("/FT /Sig") &&
                                 savedText.Contains("/V ");

            Assert.True(isStillSigned, "저장 후에도 서명 데이터(/V)가 유지되어야 합니다.");
        }

        [Fact]
        public void Verify_Signature_Should_Keep_Ber_EndOfContent_Bytes()
        {
            string samplePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestFiles", "테스트_부산_OCR_signed.pdf"));
            if (!File.Exists(samplePath))
                return;

            var service = new SignatureVerificationService();
            var result = service.VerifySignatureAtPoint(samplePath, 0, 470, 772);

            Assert.NotNull(result);
            Assert.True(result!.IsValid, result.Message);
            Assert.False(result.IsDocumentModified, result.Message);
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
    }
}
