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
            string lowerName = familyName.ToLowerInvariant();

            // 한글 폰트는 모두 Malgun Gothic으로 매핑 (TTC 파일 이슈 회피)
            if (lowerName.Contains("malgun") || lowerName == "맑은 고딕" ||
                lowerName.Contains("noto") ||
                lowerName.Contains("gulim") || lowerName == "굴림" ||
                lowerName.Contains("dotum") || lowerName == "돋움" ||
                lowerName.Contains("batang") || lowerName == "바탕")
            {
                if (isBold) return new FontResolverInfo("MalgunGothicBold");
                return new FontResolverInfo("MalgunGothic");
            }

            // 그 외 폰트
            if (isBold) return new FontResolverInfo("ArialBold");
            return new FontResolverInfo("Arial");
        }

        public byte[]? GetFont(string faceName)
        {
            switch (faceName)
            {
                case "MalgunGothic": return LoadFontData("malgun.ttf");
                case "MalgunGothicBold": return LoadFontData("malgunbd.ttf");
                // Gulim, Dotum, Batang 요청이 와도 Malgun으로 처리하도록 ResolveTypeface에서 매핑했으므로
                // 여기서는 case가 불릴 일이 없어야 하지만 안전장치로 추가
                case "Gulim": 
                case "Dotum": 
                case "Batang": return LoadFontData("malgun.ttf");
                
                case "Arial": return LoadFontData("arial.ttf");
                case "ArialBold": return LoadFontData("arialbd.ttf");
            }
            return LoadFontData("arial.ttf"); // 기본 Fallback
        }

        private byte[]? LoadFontData(string fontFileName)
        {
            var fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), fontFileName);
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            // 파일이 없으면 맑은 고딕 시도
            if (fontFileName != "malgun.ttf") return LoadFontData("malgun.ttf");
            // 그래도 없으면 Arial
            if (fontFileName != "arial.ttf") return LoadFontData("arial.ttf");
            return null;
        }
    }
}