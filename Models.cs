using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Media;
using PdfSharp.Pdf;

namespace MinsPDFViewer
{
    public enum AnnotationType { Highlight, Underline, SearchHighlight, FreeText, Other }

    public class OcrWordInfo { public string Text { get; set; } = ""; public Rect BoundingBox { get; set; } }

    public class RawAnnotationInfo { public AnnotationType Type; public Rect Rect; public string Content = ""; public double FontSize; public Color Color; }

    public class PdfAnnotation : INotifyPropertyChanged
    {
        private double _x; public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }
        private double _y; public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }
        private double _width; public double Width { get => _width; set { _width = value; OnPropertyChanged(nameof(Width)); } }
        private double _height; public double Height { get => _height; set { _height = value; OnPropertyChanged(nameof(Height)); } }

        private Brush _background = Brushes.Transparent;
        public Brush Background { get => _background; set { _background = value; OnPropertyChanged(nameof(Background)); } }

        public string TextContent { get; set; } = "";
        public AnnotationType Type { get; set; } = AnnotationType.Other;
        public Color AnnotationColor { get; set; } = Colors.Transparent;

        public double FontSize { get; set; } = 12;
        public string FontFamily { get; set; } = "Malgun Gothic";
        public bool IsBold { get; set; }
        public Brush Foreground { get; set; } = Brushes.Black;

        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        // PDF 좌표 변환용
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }
        public double CropX { get; set; }
        public double CropY { get; set; }
        // [추가] 회전 각도
        public int Rotation { get; set; }

        private ImageSource? _imageSource;
        public ImageSource? ImageSource { get => _imageSource; set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); } }
        
        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();
        public System.Collections.Generic.List<OcrWordInfo> OcrWords { get; set; } = new System.Collections.Generic.List<OcrWordInfo>();

        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }
        private double _selX; public double SelectionX { get => _selX; set { _selX = value; OnPropertyChanged(nameof(SelectionX)); } }
        private double _selY; public double SelectionY { get => _selY; set { _selY = value; OnPropertyChanged(nameof(SelectionY)); } }
        private double _selW; public double SelectionWidth { get => _selW; set { _selW = value; OnPropertyChanged(nameof(SelectionWidth)); } }
        private double _selH; public double SelectionHeight { get => _selH; set { _selH = value; OnPropertyChanged(nameof(SelectionHeight)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfDocumentModel : INotifyPropertyChanged
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public Docnet.Core.Readers.IDocReader? DocReader { get; set; }
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();

        public double SavedVerticalOffset { get; set; }
        public double SavedHorizontalOffset { get; set; }
        
        private double _zoom = 1.0;
        public double Zoom { get => _zoom; set { _zoom = value; OnPropertyChanged(nameof(Zoom)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }

    public class WindowsFontResolver : PdfSharp.Fonts.IFontResolver
    {
        public byte[]? GetFont(string faceName)
        {
            string fontPath = @"C:\Windows\Fonts\malgun.ttf";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\gulim.ttc";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            return null;
        }
        public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new PdfSharp.Fonts.FontResolverInfo("Malgun Gothic");
    }
}