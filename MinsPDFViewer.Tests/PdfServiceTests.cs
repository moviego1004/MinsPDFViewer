using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using MinsPDFViewer;
using PdfSharp.Fonts; // 폰트 설정용 네임스페이스
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
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
            // [폰트 설정] 테스트 실행 시 폰트 리졸버가 없으면 등록
            if (GlobalFontSettings.FontResolver == null)
            {
                GlobalFontSettings.FontResolver = new TestFontResolver();
            }

            // 테스트 파일들이 저장될 임시 폴더
            _testFileDir = Path.Combine(AppContext.BaseDirectory, "TestFiles");
            if (!Directory.Exists(_testFileDir))
                Directory.CreateDirectory(_testFileDir);

            _samplePdfPath = Path.Combine(_testFileDir, "normal.pdf");
            _signedPdfPath = Path.Combine(_testFileDir, "signed.pdf");

            // 샘플 PDF 생성
            if (!File.Exists(_samplePdfPath))
            {
                using (var doc = new PdfDocument())
                {
                    doc.AddPage();
                    doc.Save(_samplePdfPath);
                }
            }
        }

        // [TC-11 & 12] 저장 후 데이터가 유지되는지 확인
        [Fact]
        public async Task Save_And_Reload_Should_Keep_Annotations()
        {
            // 1. 문서 로드
            var service = new PdfService();
            var docModel = await service.LoadPdfAsync(_samplePdfPath);
            Assert.NotNull(docModel);
            await service.InitializeDocumentAsync(docModel);

            var page = docModel.Pages[0];

            // 2. 텍스트 박스 추가 (Malgun Gothic 사용)
            var newAnnot = new MPdfAnnotation
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

            string savePath = Path.Combine(_testFileDir, "tc11_result.pdf");

            // 3. 저장 수행 (여기서 폰트 리졸버가 작동해야 함)
            await service.SavePdf(docModel, savePath);

            // 4. 검증: 저장된 파일을 직접 열어서 주석 확인
            bool annotationFound = false;
            using (var doc = PdfReader.Open(savePath, PdfDocumentOpenMode.Import))
            {
                var savedPage = doc.Pages[0];
                if (savedPage.Annotations != null)
                {
                    foreach (var ann in savedPage.Annotations)
                    {
                        var dict = ann as PdfSharp.Pdf.PdfDictionary;
                        if (dict == null && ann is PdfSharp.Pdf.Advanced.PdfReference r)
                            dict = r.Value as PdfSharp.Pdf.PdfDictionary;

                        if (dict != null && dict.Elements.GetString("/Subtype") == "/FreeText")
                        {
                            if (dict.Elements.GetString("/Contents") == "Hello Test")
                            {
                                annotationFound = true;
                                break;
                            }
                        }
                    }
                }
            }

            Assert.True(annotationFound, "저장된 PDF 파일 내에 텍스트 박스(FreeText)가 존재해야 합니다.");
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

            Assert.True(docModel.IsReadOnlyMode, "서명된 파일은 IsReadOnlyMode여야 합니다.");

            string savePath = Path.Combine(_testFileDir, "tc18_result_signed.pdf");

            await service.SavePdf(docModel, savePath);

            bool isStillSigned = false;
            using (var doc = PdfReader.Open(savePath, PdfDocumentOpenMode.Import))
            {
                foreach (var page in doc.Pages)
                {
                    if (page.Annotations != null)
                    {
                        foreach (var item in page.Annotations)
                        {
                            var dict = item as PdfSharp.Pdf.PdfDictionary;
                            if (dict == null && item is PdfSharp.Pdf.Advanced.PdfReference r)
                                dict = r.Value as PdfSharp.Pdf.PdfDictionary;

                            if (dict != null &&
                                dict.Elements.GetString("/Subtype") == "/Widget" &&
                                dict.Elements.GetString("/FT") == "/Sig" &&
                                dict.Elements.ContainsKey("/V"))
                            {
                                isStillSigned = true;
                                break;
                            }
                        }
                    }
                    if (isStillSigned)
                        break;
                }
            }

            Assert.True(isStillSigned, "저장 후에도 서명 데이터(/V)가 유지되어야 합니다.");
        }
    }

    // [폰트 리졸버 구현] 테스트 환경에서 폰트 파일을 찾아주는 도우미 클래스
    public class TestFontResolver : IFontResolver
    {
        public string DefaultFontName => "Arial";

        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // 한글 폰트 요청 시 맑은 고딕으로 매핑
            if (familyName.Equals("Malgun Gothic", StringComparison.OrdinalIgnoreCase))
            {
                return new FontResolverInfo("Malgun Gothic");
            }
            // 그 외에는 Arial로 매핑
            return new FontResolverInfo("Arial");
        }

        public byte[] GetFont(string faceName)
        {
            string fontPath = "";

            // 윈도우 폰트 경로에서 파일 찾기
            if (faceName.Contains("Malgun", StringComparison.OrdinalIgnoreCase))
            {
                fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "malgun.ttf");
            }
            else
            {
                fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");
            }

            if (File.Exists(fontPath))
            {
                return File.ReadAllBytes(fontPath);
            }

            return new byte[0]; // 폰트 못 찾음
        }
    }
}