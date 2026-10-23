using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using MinsPDFViewer;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using Xunit;

namespace MinsPDFViewer.Tests
{
    /// <summary>
    /// 증분 기록기(북마크·서명·OCR 층)와 서명 검증: 교차 참조 스트림·선형화 문서(normal.pdf), 기존 서명 유지,
    /// 주석 좌표 왕복(CropBox 오프셋 + /Rotate 270 페이지).
    /// </summary>
    public class IncrementalWriterTests : IDisposable
    {
        private readonly string _dir;

        public IncrementalWriterTests()
        {
            PdfService.UseDirectDispatcherForTests = true;
            _dir = Path.Combine(Path.GetTempPath(), "MinsIncTests_" + Guid.NewGuid().ToString("N"));
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

        private string CopyFixture(string name, string? target = null)
        {
            string? testFiles = PdfServiceTests.FindTestFilesDirectory();
            Assert.NotNull(testFiles);
            string path = Path.Combine(_dir, target ?? name);
            File.Copy(Path.Combine(testFiles!, name), path, true);
            return path;
        }

        [Fact]
        public void ObjectStore_ReadsXrefStreamAndLinearizedFile()
        {
            string path = CopyFixture("normal.pdf");
            using var store = PdfObjectStore.Open(path);
            Assert.True(store.LastXrefIsStream);
            Assert.Equal(5, store.GetPageRefs().Count);
            Assert.All(store.ObjectNumbers, n => Assert.NotNull(store.GetObject(n, 0)));
        }

        [Fact]
        public async Task Bookmarks_OnXrefStreamPdf_AreWrittenAndReadable()
        {
            string path = CopyFixture("normal.pdf");
            byte[] before = File.ReadAllBytes(path);
            await new PdfBookmarkRewriteService().RewriteBookmarksAsync(path, new List<BookmarkSaveData>
            {
                new() { Title = "첫 장", PageIndex = 0, Children = { new() { Title = "하위 (괄호)", PageIndex = 1 } } },
                new() { Title = "끝", PageIndex = 4 }
            });

            byte[] after = File.ReadAllBytes(path);
            Assert.True(after.AsSpan(0, before.Length).SequenceEqual(before)); // 증분: 원본 바이트 그대로

            using var doc = PdfiumDocument.Open(path, null);
            List<PdfiumDocument.BookmarkNode> bookmarks;
            lock (Pdfium.Lock)
                bookmarks = doc.GetBookmarks();
            Assert.Equal(new[] { "첫 장", "끝" }, bookmarks.Select(b => b.Title));
            Assert.Equal(4, bookmarks[1].PageIndex);
            Assert.Equal("하위 (괄호)", bookmarks[0].Children.Single().Title);
            Assert.Equal(1, bookmarks[0].Children.Single().PageIndex);
        }

        private static SignatureConfig CreateTestSignatureConfig()
        {
            var generator = new RsaKeyPairGenerator();
            generator.Init(new KeyGenerationParameters(new SecureRandom(), 2048));
            var keys = generator.GenerateKeyPair();
            var name = new X509Name("CN=Mins Test,O=Mins,C=KR");
            var certGenerator = new X509V3CertificateGenerator();
            certGenerator.SetSerialNumber(BigInteger.ValueOf(DateTime.Now.Ticks));
            certGenerator.SetIssuerDN(name);
            certGenerator.SetSubjectDN(name);
            certGenerator.SetNotBefore(DateTime.UtcNow.AddDays(-1));
            certGenerator.SetNotAfter(DateTime.UtcNow.AddYears(1));
            certGenerator.SetPublicKey(keys.Public);
            var cert = certGenerator.Generate(new Asn1SignatureFactory("SHA256WITHRSA", keys.Private));
            return new SignatureConfig
            {
                Certificate = new NpkiCertificateModel { Subject = cert.SubjectDN.ToString(), Certificate = cert, NotAfter = cert.NotAfter },
                PrivateKey = keys.Private,
                Reason = "검토",
                Location = "서울"
            };
        }

        [Fact]
        public void Sign_XrefStreamPdf_Twice_BothSignaturesValid()
        {
            string source = CopyFixture("normal.pdf");
            string once = Path.Combine(_dir, "once.pdf");
            string twice = Path.Combine(_dir, "twice.pdf");
            var service = new PdfSignatureService();
            var config = CreateTestSignatureConfig();

            service.SignPdf(source, once, config, 0, new SignaturePdfRect(100, 100, 60, 120));
            service.SignPdf(once, twice, config, 1, null);

            var verifier = new SignatureVerificationService();
            using (var doc = PdfiumDocument.Open(twice, null))
            {
                lock (Pdfium.Lock)
                    Assert.Equal(2, doc.SignatureCount);
            }

            var first = verifier.VerifySignatureAtPoint(twice, 0, 130, 160);
            Assert.NotNull(first);
            Assert.True(first!.IsValid, first.Message);
            Assert.False(first.CoversWholeDocument); // 두 번째 서명이 뒤에 붙음

            var firstAlone = verifier.VerifySignatureAtPoint(once, 0, 130, 160);
            Assert.NotNull(firstAlone);
            Assert.True(firstAlone!.IsValid, firstAlone.Message);
            Assert.True(firstAlone.CoversWholeDocument);
            Assert.Equal("Mins Test", firstAlone.SignerName);

            // 서명 범위 안의 바이트를 바꾸면 변조로 나와야 한다
            string tampered = Path.Combine(_dir, "tampered.pdf");
            byte[] bytes = File.ReadAllBytes(once);
            bytes[5000] ^= 1;
            File.WriteAllBytes(tampered, bytes);
            var bad = verifier.VerifySignatureAtPoint(tampered, 0, 130, 160);
            Assert.NotNull(bad);
            Assert.False(bad!.IsValid);
        }

        [Fact]
        public async Task SavingAnnotations_OnSignedPdf_KeepsExistingSignature()
        {
            string path = CopyFixture("signed.pdf");
            var service = new PdfService();
            var model = await service.LoadPdfAsync(path);
            Assert.NotNull(model);
            await service.InitializeDocumentAsync(model!);
            var page = model!.Pages[1];
            service.RenderPageImage(model, page);
            page.Annotations.Add(new PdfAnnotation
            {
                Type = AnnotationType.Highlight,
                X = 50, Y = 50, Width = 80, Height = 12,
                AnnotationColor = Colors.Yellow,
                Background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0))
            });

            string output = Path.Combine(_dir, "signed_annotated.pdf");
            await service.SavePdf(model, output);
            await model.CloseAsync();

            byte[] original = File.ReadAllBytes(path);
            byte[] saved = File.ReadAllBytes(output);
            Assert.True(saved.Length > original.Length);
            Assert.True(saved.AsSpan(0, original.Length).SequenceEqual(original)); // 증분 저장

            var verifier = new SignatureVerificationService();
            var result = verifier.VerifySignatureForAnnotation(output, 0, "Signature8064c35a", default);
            Assert.NotNull(result);
            Assert.True(result!.IsValid, result.Message);
            Assert.False(result.CoversWholeDocument);
        }

        [Fact]
        public async Task Annotations_RoundTrip_OnCroppedRotatedPage()
        {
            // normal.pdf 1쪽: CropBox가 원점에서 떨어져 있고 /Rotate 270
            string path = CopyFixture("normal.pdf");
            var service = new PdfService();
            var model = await service.LoadPdfAsync(path);
            Assert.NotNull(model);
            await service.InitializeDocumentAsync(model!);
            var page = model!.Pages[0];
            service.RenderPageImage(model, page);
            int existing = page.Annotations.Count;

            page.Annotations.Add(new PdfAnnotation
            {
                Type = AnnotationType.Highlight,
                X = 40, Y = 60, Width = 120, Height = 20,
                AnnotationColor = Colors.Yellow,
                Background = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0))
            });
            page.Annotations.Add(new PdfAnnotation
            {
                Type = AnnotationType.FreeText,
                X = 200, Y = 300, Width = 160, Height = 40,
                TextContent = "회전 페이지",
                FontSize = 16,
                FontFamily = "Malgun Gothic"
            });

            string output = Path.Combine(_dir, "rotated_roundtrip.pdf");
            await service.SavePdf(model, output);
            await model.CloseAsync();

            var reloaded = await service.LoadPdfAsync(output);
            Assert.NotNull(reloaded);
            await service.InitializeDocumentAsync(reloaded!);
            var reloadedPage = reloaded!.Pages[0];
            service.RenderPageImage(reloaded, reloadedPage);
            Assert.Equal(existing + 2, reloadedPage.Annotations.Count);

            var highlight = reloadedPage.Annotations.Single(a => a.Type == AnnotationType.Highlight);
            Assert.InRange(highlight.X, 39, 41);
            Assert.InRange(highlight.Y, 59, 61);
            Assert.InRange(highlight.Width, 119, 121);
            Assert.InRange(highlight.Height, 19, 21);

            var text = reloadedPage.Annotations.Single(a => a.Type == AnnotationType.FreeText);
            Assert.Equal("회전 페이지", text.TextContent);
            Assert.InRange(text.X, 199, 201);
            Assert.InRange(text.Y, 299, 301);
            Assert.InRange(text.Width, 159, 161);
            Assert.InRange(text.Height, 39, 41);
            Assert.InRange(text.FontSize, 15.5, 16.5);
            await reloaded.CloseAsync();
        }

        [Fact]
        public void PageSize_HundredPercent_IsActualSize()
        {
            // 100% = 실제 크기: A4 폭 595pt = 8.26인치 = 793 DIP
            Assert.Equal(793.3, 595 * PdfDocumentModel.DipPerPoint, 1);
        }
    }
}
