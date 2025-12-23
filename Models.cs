using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace MinsPDFViewer
{
    // 주석 타입 열거형
    public enum AnnotationType { Highlight, Underline, SearchHighlight, FreeText, Other }

    // OCR 단어 정보
    public class OcrWordInfo { public string Text { get; set; } = ""; public Rect BoundingBox { get; set; } }

    // PDFSharp 원본 주석 정보
    public class RawAnnotationInfo { public AnnotationType Type; public Rect Rect; public string Content = ""; public double FontSize; public Color Color; }

    // [ViewModel] 주석
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

        // 폰트 스타일 속성 생략 가능 (필요시 추가)
        public double FontSize { get; set; } = 12;
        public string FontFamily { get; set; } = "Malgun Gothic";
        public bool IsBold { get; set; }
        public Brush Foreground { get; set; } = Brushes.Black;

        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // [ViewModel] 페이지
    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }  // View(이미지) 기준 너비
        public double Height { get; set; } // View(이미지) 기준 높이
        
        // 원본 PDF 포인트 단위 크기 (좌표 변환용)
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }

        private ImageSource? _imageSource;
        public ImageSource? ImageSource { get => _imageSource; set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); } }
        
        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();
        public System.Collections.Generic.List<OcrWordInfo> OcrWords { get; set; } = new System.Collections.Generic.List<OcrWordInfo>();

        // 드래그 선택 관련 속성
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
}