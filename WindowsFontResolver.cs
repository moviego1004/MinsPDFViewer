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
            // 간단한 구현: 윈도우 폰트 폴더에서 해당 폰트를 찾거나 기본 폰트로 매핑
            string fontName = familyName.ToLower();

            // 한글 폰트 매핑 (필요 시 추가)
            if (fontName.Contains("malgun") || fontName.Contains("맑은"))
                return new FontResolverInfo("Malgun Gothic");
            if (fontName.Contains("gulim") || fontName.Contains("굴림"))
                return new FontResolverInfo("Gulim");
            if (fontName.Contains("dotum") || fontName.Contains("돋움"))
                return new FontResolverInfo("Dotum");
            if (fontName.Contains("batang") || fontName.Contains("바탕"))
                return new FontResolverInfo("Batang");

            // 기본 폰트 처리
            if (isBold && isItalic) return new FontResolverInfo("Arial", true, true);
            if (isBold) return new FontResolverInfo("Arial", true, false);
            if (isItalic) return new FontResolverInfo("Arial", false, true);

            return new FontResolverInfo("Arial");
        }

        public byte[]? GetFont(string faceName)
        {
            string fontPath = "";
            string lowerFaceName = faceName.ToLower();

            // 실제 파일 경로 매핑
            if (lowerFaceName.Contains("malgun")) fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "malgun.ttf");
            else if (lowerFaceName.Contains("gulim")) fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "gulim.ttc");
            else if (lowerFaceName.Contains("dotum")) fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "gulim.ttc"); // 굴림/돋움 같은 파일
            else if (lowerFaceName.Contains("batang")) fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "batang.ttc");
            else fontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf");

            if (File.Exists(fontPath))
            {
                return File.ReadAllBytes(fontPath);
            }
            return null; // 폰트 찾기 실패
        }
    }
}