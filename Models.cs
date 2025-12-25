using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using PdfSharp.Pdf;

namespace MinsPDFViewer
{
    // 주석 타입
    public enum AnnotationType { Highlight, Underline, SearchHighlight, FreeText, Other }

    // OCR 정보
    public class OcrWordInfo { public string Text { get; set; } = ""; public Rect BoundingBox { get; set; } }

    // [ViewModel] 주석
    public class PdfAnnotation : INotifyPropertyChanged
    {
        private double _x; public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }
        private double _y; public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }
        private double _width; public double Width { get => _width; set { _width = value; OnPropertyChanged(nameof(Width)); } }
        private double _height; public double Height { get => _height; set { _height = value; OnPropertyChanged(nameof(Height)); } }

        private Brush _background = Brushes.Transparent;
        public Brush Background { get => _background; set { _background = value; OnPropertyChanged(nameof(Background)); } }

        private string _textContent = "";
        public string TextContent { get => _textContent; set { _textContent = value; OnPropertyChanged(nameof(TextContent)); } }

        private AnnotationType _type = AnnotationType.Other;
        public AnnotationType Type 
        { 
            get => _type; 
            set 
            { 
                _type = value; 
                OnPropertyChanged(nameof(Type)); 
                OnPropertyChanged(nameof(IsFreeText)); 
            } 
        }

        public bool IsFreeText => Type == AnnotationType.FreeText;

        private Color _annotationColor = Colors.Transparent;
        public Color AnnotationColor { get => _annotationColor; set { _annotationColor = value; OnPropertyChanged(nameof(AnnotationColor)); } }

        private double _fontSize = 12;
        public double FontSize { get => _fontSize; set { _fontSize = value; OnPropertyChanged(nameof(FontSize)); } }

        private string _fontFamily = "Malgun Gothic";
        public string FontFamily { get => _fontFamily; set { _fontFamily = value; OnPropertyChanged(nameof(FontFamily)); } }

        private bool _isBold;
        public bool IsBold { get => _isBold; set { _isBold = value; OnPropertyChanged(nameof(IsBold)); } }

        private Brush _foreground = Brushes.Black;
        public Brush Foreground { get => _foreground; set { _foreground = value; OnPropertyChanged(nameof(Foreground)); } }

        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // [ViewModel] 페이지
    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }  // 화면(이미지) 너비 (Pixel)
        public double Height { get; set; } // 화면(이미지) 높이 (Pixel)
        
        // 좌표 변환용 정보
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }
        
        // [수정] CropBox 정보 (Point 단위)
        public double CropX { get; set; }
        public double CropY { get; set; }
        public double CropWidthPoint { get; set; }  // 실제 보이는 영역 너비
        public double CropHeightPoint { get; set; } // 실제 보이는 영역 높이
        
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

    // [ViewModel] 문서 모델
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
            if (System.IO.File.Exists(fontPath)) return System.IO.File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\gulim.ttc";
            if (System.IO.File.Exists(fontPath)) return System.IO.File.ReadAllBytes(fontPath);
            return null;
        }
        public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new PdfSharp.Fonts.FontResolverInfo("Malgun Gothic");
    }
}