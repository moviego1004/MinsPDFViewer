using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Media;
using PdfSharp.Pdf;

namespace MinsPDFViewer
{

    public enum AnnotationType
    {
        Highlight,
        Underline,
        FreeText,
        SearchHighlight,
        SignatureField,
        SignaturePlaceholder
    }

    public class PdfDocumentModel : INotifyPropertyChanged
    {
        // [필수 추가] 멀티스레드 접근 충돌 방지용 자물쇠
        public object SyncRoot { get; } = new object();

        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";

        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();

        private double _zoom = 1.0;
        public double Zoom
        {
            get => _zoom;
            set
            {
                _zoom = value;
                OnPropertyChanged(nameof(Zoom));
            }
        }

        public double SavedVerticalOffset
        {
            get; set;
        }
        public double SavedHorizontalOffset
        {
            get; set;
        }

        public Docnet.Core.IDocLib? DocLib
        {
            get; set;
        }
        public Docnet.Core.Readers.IDocReader? DocReader
        {
            get; set;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex
        {
            get; set;
        }
        public double Width
        {
            get; set;
        }
        public double Height
        {
            get; set;
        }
        public ImageSource? ImageSource
        {
            get; set;
        }

        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();

        public double PdfPageWidthPoint
        {
            get; set;
        }
        public double PdfPageHeightPoint
        {
            get; set;
        }

        public double CropX
        {
            get; set;
        }
        public double CropY
        {
            get; set;
        }
        public double CropWidthPoint
        {
            get; set;
        }
        public double CropHeightPoint
        {
            get; set;
        }

        private bool _isSelecting;
        public bool IsSelecting
        {
            get => _isSelecting; set
            {
                _isSelecting = value;
                OnPropertyChanged(nameof(IsSelecting));
            }
        }

        private double _selectionX;
        public double SelectionX
        {
            get => _selectionX; set
            {
                _selectionX = value;
                OnPropertyChanged(nameof(SelectionX));
            }
        }
        private double _selectionY;
        public double SelectionY
        {
            get => _selectionY; set
            {
                _selectionY = value;
                OnPropertyChanged(nameof(SelectionY));
            }
        }
        private double _selectionWidth;
        public double SelectionWidth
        {
            get => _selectionWidth; set
            {
                _selectionWidth = value;
                OnPropertyChanged(nameof(SelectionWidth));
            }
        }
        private double _selectionHeight;
        public double SelectionHeight
        {
            get => _selectionHeight; set
            {
                _selectionHeight = value;
                OnPropertyChanged(nameof(SelectionHeight));
            }
        }

        public List<OcrWordInfo>? OcrWords
        {
            get; set;
        }

        public bool HasSignature
        {
            get; set;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfAnnotation : INotifyPropertyChanged
    {
        private double _x, _y, _width, _height;
        public double X
        {
            get => _x; set
            {
                _x = value;
                OnPropertyChanged(nameof(X));
            }
        }
        public double Y
        {
            get => _y; set
            {
                _y = value;
                OnPropertyChanged(nameof(Y));
            }
        }
        public double Width
        {
            get => _width; set
            {
                _width = value;
                OnPropertyChanged(nameof(Width));
            }
        }
        public double Height
        {
            get => _height; set
            {
                _height = value;
                OnPropertyChanged(nameof(Height));
            }
        }

        public AnnotationType Type
        {
            get; set;
        }
        public Color AnnotationColor { get; set; } = Colors.Yellow;
        public Brush Background { get; set; } = Brushes.Transparent;

        private string _textContent = "";
        public string TextContent
        {
            get => _textContent;
            set
            {
                _textContent = value;
                OnPropertyChanged(nameof(TextContent));
            }
        }

        private double _fontSize = 12;
        public double FontSize
        {
            get => _fontSize;
            set
            {
                _fontSize = value;
                OnPropertyChanged(nameof(FontSize));
            }
        }

        private string _fontFamily = "Malgun Gothic";
        public string FontFamily
        {
            get => _fontFamily;
            set
            {
                _fontFamily = value;
                OnPropertyChanged(nameof(FontFamily));
            }
        }

        private Brush _foreground = Brushes.Black;
        public Brush Foreground
        {
            get => _foreground;
            set
            {
                _foreground = value;
                OnPropertyChanged(nameof(Foreground));
            }
        }

        private bool _isBold = false;
        public bool IsBold
        {
            get => _isBold;
            set
            {
                _isBold = value;
                OnPropertyChanged(nameof(IsBold));
            }
        }

        // [필수 추가] 서명 필드 이름
        public string FieldName { get; set; } = "";

        public object? SignatureData
        {
            get; set;
        }

        private string? _visualStampPath;
        public string? VisualStampPath
        {
            get => _visualStampPath;
            set
            {
                _visualStampPath = value;
                OnPropertyChanged(nameof(VisualStampPath));
            }
        }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected; set
            {
                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class OcrWordInfo
    {
        public string Text { get; set; } = "";
        public System.Windows.Rect BoundingBox
        {
            get; set;
        }
    }

    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }
}