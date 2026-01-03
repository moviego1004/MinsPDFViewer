using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using Docnet.Core.Readers;
using PdfSharp.Pdf;

namespace MinsPDFViewer
{
    public enum AnnotationType
    {
        Highlight, Underline, FreeText, SearchHighlight, SignatureField, SignaturePlaceholder
    }

    public class PdfDocumentModel : INotifyPropertyChanged, IDisposable
    {
        public object SyncRoot { get; } = new object();
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();

        private double _zoom = 1.0;
        public double Zoom
        {
            get => _zoom; set
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
        public IDocReader? DocReader
        {
            get; set;
        }
        public IDocReader? CleanDocReader
        {
            get; set;
        }

        // [신규] 책갈피
        public ObservableCollection<PdfBookmarkViewModel> Bookmarks { get; set; } = new ObservableCollection<PdfBookmarkViewModel>();

        public void Dispose()
        {
            lock (PdfService.PdfiumLock)
            {
                if (CleanDocReader != null)
                {
                    try
                    {
                        CleanDocReader.Dispose();
                    }
                    catch { }
                    CleanDocReader = null;
                }
                if (DocReader != null)
                {
                    try
                    {
                        DocReader.Dispose();
                    }
                    catch { }
                    DocReader = null;
                }
            }
            if (Pages != null)
                Application.Current.Dispatcher.Invoke(() => Pages.Clear());
            if (Bookmarks != null)
                Application.Current.Dispatcher.Invoke(() => Bookmarks.Clear());
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        // =========================================================
        // [신규] 페이지 편집 속성 (회전, 삭제, 재조립용)
        // =========================================================
        public string? OriginalFilePath
        {
            get; set;
        } // 원본 경로 (빈 페이지면 null 가능)
        public int OriginalPageIndex
        {
            get; set;
        }
        public bool IsBlankPage
        {
            get; set;
        }

        private int _rotation;
        public int Rotation
        {
            get => _rotation;
            set
            {
                _rotation = value;
                OnPropertyChanged(nameof(Rotation));
            }
        }

        // =========================================================
        // [복구] 기존 OCR 및 검색, 메모리 관리 속성
        // =========================================================
        public List<OcrWordInfo>? OcrWords
        {
            get; set;
        } // OCR 결과
        public string MediaBoxInfo { get; set; } = "";   // 디버깅용

        // 메모리 해제 메서드
        public void Unload()
        {
            ImageSource = null;
            // 필요하다면 여기서 OcrWords 등도 정리 가능
        }

        // =========================================================
        // 기존 뷰어 속성
        // =========================================================
        public int PageIndex
        {
            get; set;
        }

        private double _width;
        public double Width
        {
            get => _width; set
            {
                _width = value;
                OnPropertyChanged(nameof(Width));
            }
        }

        private double _height;
        public double Height
        {
            get => _height; set
            {
                _height = value;
                OnPropertyChanged(nameof(Height));
            }
        }

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
        public bool HasSignature
        {
            get; set;
        }

        private ImageSource? _imageSource;
        public ImageSource? ImageSource
        {
            get => _imageSource; set
            {
                _imageSource = value;
                OnPropertyChanged(nameof(ImageSource));
            }
        }

        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();

        // 텍스트 선택용
        public bool IsSelecting
        {
            get; set;
        }
        public double SelectionX
        {
            get; set;
        }
        public double SelectionY
        {
            get; set;
        }
        public double SelectionWidth
        {
            get; set;
        }
        public double SelectionHeight
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
            get => _textContent; set
            {
                _textContent = value;
                OnPropertyChanged(nameof(TextContent));
            }
        }

        private double _fontSize = 12;
        public double FontSize
        {
            get => _fontSize; set
            {
                _fontSize = value;
                OnPropertyChanged(nameof(FontSize));
            }
        }

        private string _fontFamily = "Malgun Gothic";
        public string FontFamily
        {
            get => _fontFamily; set
            {
                _fontFamily = value;
                OnPropertyChanged(nameof(FontFamily));
            }
        }

        private Brush _foreground = Brushes.Black;
        public Brush Foreground
        {
            get => _foreground; set
            {
                _foreground = value;
                OnPropertyChanged(nameof(Foreground));
            }
        }

        private bool _isBold;
        public bool IsBold
        {
            get => _isBold; set
            {
                _isBold = value;
                OnPropertyChanged(nameof(IsBold));
            }
        }

        public string FieldName { get; set; } = "";
        public object? SignatureData
        {
            get; set;
        }

        private string? _visualStampPath;
        public string? VisualStampPath
        {
            get => _visualStampPath; set
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
        public Rect BoundingBox
        {
            get; set;
        }
    }

    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }

    public class PdfBookmarkViewModel : INotifyPropertyChanged
    {
        private string _title = "";
        public string Title
        {
            get => _title; set
            {
                _title = value;
                OnPropertyChanged(nameof(Title));
            }
        }

        private int _pageIndex;
        public int PageIndex
        {
            get => _pageIndex; set
            {
                _pageIndex = value;
                OnPropertyChanged(nameof(PageIndex));
            }
        }

        private bool _isExpanded = true;
        public bool IsExpanded
        {
            get => _isExpanded; set
            {
                _isExpanded = value;
                OnPropertyChanged(nameof(IsExpanded));
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

        public ObservableCollection<PdfBookmarkViewModel> Children { get; set; } = new ObservableCollection<PdfBookmarkViewModel>();
        public PdfBookmarkViewModel? Parent
        {
            get; set;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}