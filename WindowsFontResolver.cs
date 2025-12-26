using System;
using System.IO;
using System.Linq;
using PdfSharp.Fonts;

namespace MinsPDFViewer
{
    public class WindowsFontResolver : IFontResolver
    {
        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            string fontName = familyName.ToLower();

            // 한글 폰트 매핑
            if (fontName.Contains("malgun") || fontName.Contains("맑은"))
                return new FontResolverInfo("Malgun Gothic", isBold, isItalic);
            if (fontName.Contains("gulim") || fontName.Contains("굴림"))
                return new FontResolverInfo("Gulim", isBold, isItalic);
            if (fontName.Contains("dotum") || fontName.Contains("돋움"))
                return new FontResolverInfo("Dotum", isBold, isItalic);
            if (fontName.Contains("batang") || fontName.Contains("바탕"))
                return new FontResolverInfo("Batang", isBold, isItalic);

            // 기본 폰트
            return new FontResolverInfo("Arial", isBold, isItalic);
        }

        public byte[]? GetFont(string faceName)
        {
            string fontPath = "";
            string lowerFaceName = faceName.ToLower();
            string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

            if (lowerFaceName.Contains("malgun")) 
            {
                // [수정] ttc 우선 확인
                fontPath = Path.Combine(fontsFolder, "malgun.ttc");
                if (!File.Exists(fontPath)) fontPath = Path.Combine(fontsFolder, "malgun.ttf");
            }
            else if (lowerFaceName.Contains("gulim")) fontPath = Path.Combine(fontsFolder, "gulim.ttc");
            else if (lowerFaceName.Contains("dotum")) fontPath = Path.Combine(fontsFolder, "gulim.ttc");
            else if (lowerFaceName.Contains("batang")) fontPath = Path.Combine(fontsFolder, "batang.ttc");
            else fontPath = Path.Combine(fontsFolder, "arial.ttf");

            if (File.Exists(fontPath))
            {
                return File.ReadAllBytes(fontPath);
            }
            return null; 
        }
    }
}