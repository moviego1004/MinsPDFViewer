using System.Collections.Generic;
using System.Windows.Media;

namespace MinsPDFViewer
{
    internal sealed class PageSaveData
    {
        public int OriginalPageIndex { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }
        public List<AnnotationSaveData> Annotations { get; set; } = new();
        public List<OcrWordInfo> OcrWords { get; set; } = new();
    }

    internal sealed class AnnotationSaveData
    {
        public AnnotationType Type { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public string TextContent { get; set; } = "";
        public double FontSize { get; set; }
        public string FontFamily { get; set; } = "Malgun Gothic";
        public bool IsBold { get; set; }
        public Color ForegroundColor { get; set; }
        public Color BackgroundColor { get; set; }
    }

    internal sealed class BookmarkSaveData
    {
        public string Title { get; set; } = "";
        public int OriginalPageIndex { get; set; }
        public List<BookmarkSaveData> Children { get; set; } = new();
    }
}
