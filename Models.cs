using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Media;
using PdfSharp.Pdf; // [필수] PdfDocument 참조용

namespace MinsPDFViewer
{
    // 주석 타입
    public enum AnnotationType
    {
        Highlight,
        Underline,
        FreeText,
        SearchHighlight,
        SignatureField // 클릭 가능한 서명 영역
    }

    // 문서 모델
    public class PdfDocumentModel : INotifyPropertyChanged
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        private double _zoom = 1.0;
        public double Zoom
        {
            get => _zoom;
            set { _zoom = value; OnPropertyChanged(nameof(Zoom)); }
        }

        public double SavedVerticalOffset { get; set; }
        public double SavedHorizontalOffset { get; set; }

        public Docnet.Core.IDocLib? DocLib { get; set; }
        public Docnet.Core.Readers.IDocReader? DocReader { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // 페이지 뷰모델
    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; } 
        public double Height { get; set; }
        public ImageSource? ImageSource { get; set; }
        
        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();

        // 원본 PDF 페이지 크기 (Point 단위)
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }
        
        // CropBox 정보
        public double CropX { get; set; }
        public double CropY { get; set; }
        public double CropWidthPoint { get; set; }
        public double CropHeightPoint { get; set; }
        
        // 드래그 선택 관련
        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }

        private double _selectionX;
        public double SelectionX { get => _selectionX; set { _selectionX = value; OnPropertyChanged(nameof(SelectionX)); } }
        private double _selectionY;
        public double SelectionY { get => _selectionY; set { _selectionY = value; OnPropertyChanged(nameof(SelectionY)); } }
        private double _selectionWidth;
        public double SelectionWidth { get => _selectionWidth; set { _selectionWidth = value; OnPropertyChanged(nameof(SelectionWidth)); } }
        private double _selectionHeight;
        public double SelectionHeight { get => _selectionHeight; set { _selectionHeight = value; OnPropertyChanged(nameof(SelectionHeight)); } }

        // OCR 데이터
        public List<OcrWordInfo>? OcrWords { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // 주석 뷰모델
    public class PdfAnnotation : INotifyPropertyChanged
    {
        private double _x, _y, _width, _height;
        public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }
        public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }
        public double Width { get => _width; set { _width = value; OnPropertyChanged(nameof(Width)); } }
        public double Height { get => _height; set { _height = value; OnPropertyChanged(nameof(Height)); } }

        public AnnotationType Type { get; set; }
        public Color AnnotationColor { get; set; } = Colors.Yellow; 
        public Brush Background { get; set; } = Brushes.Transparent;

        // FreeText용 속성
        public string TextContent { get; set; } = "";
        public double FontSize { get; set; } = 12;
        // [주의] 여기서 FontFamily는 string 타입입니다.
        public string FontFamily { get; set; } = "Malgun Gothic"; 
        public Brush Foreground { get; set; } = Brushes.Black;
        public bool IsBold { get; set; } = false;

        // 서명 필드용 데이터
        public object? SignatureData { get; set; } 

        private bool _isSelected;
        public bool IsSelected { get => _isSelected; set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class OcrWordInfo
    {
        public string Text { get; set; } = "";
        public System.Windows.Rect BoundingBox { get; set; }
    }

    // [중요] PdfSignatureService에서 사용하는 클래스 정의 추가
    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }
}