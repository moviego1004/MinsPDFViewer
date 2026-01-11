using System;
using System.IO;
using System.Linq;
using PdfSharp.Fonts;

namespace MinsPDFViewer
{
    // PdfSharp이 시스템 폰트를 사용할 수 있게 해주는 해결사 클래스
    public class WindowsFontResolver : IFontResolver
    {
        public string DefaultFontName => "Malgun Gothic"; // 기본 폰트: 맑은 고딕

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            // 폰트 이름이 무엇이든 요청이 오면 처리
            // 실제로는 familyName에 따라 다른 파일을 지정할 수 있음

            // "Malgun Gothic" 또는 한글 처리를 위해 기본적으로 맑은 고딕으로 매핑
            if (familyName.Contains("Malgun", StringComparison.OrdinalIgnoreCase) ||
                familyName.Equals("맑은 고딕", StringComparison.OrdinalIgnoreCase))
            {
                if (isBold)
                    return new FontResolverInfo("MalgunGothicBold");
                return new FontResolverInfo("MalgunGothic");
            }

            // 그 외 폰트는 Arial 등으로 대체하거나 시스템 폰트 로직 확장 가능
            // 여기서는 편의상 모두 맑은 고딕 또는 Arial로 처리
            if (isBold)
                return new FontResolverInfo("ArialBold");
            return new FontResolverInfo("Arial");
        }

        public byte[]? GetFont(string faceName)
        {
            switch (faceName)
            {
                case "MalgunGothic":
                    return LoadFontData("malgun.ttf");
                case "MalgunGothicBold":
                    return LoadFontData("malgunbd.ttf"); // 맑은 고딕 볼드
                case "Arial":
                    return LoadFontData("arial.ttf");
                case "ArialBold":
                    return LoadFontData("arialbd.ttf");
            }
            return null;
        }

        private byte[]? LoadFontData(string fontFileName)
        {
            // 윈도우 폰트 폴더 경로
            var fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), fontFileName);

            if (File.Exists(fontPath))
            {
                return File.ReadAllBytes(fontPath);
            }
            return null;
        }
    }
}