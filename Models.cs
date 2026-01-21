using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using PdfiumViewer;
// [충돌 방지] PdfSharp의 PdfDocument는 명시적으로 별칭 사용
using PdfSharpDoc = PdfSharp.Pdf.PdfDocument;

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

        // [PdfiumViewer] 문서 객체
        public IPdfDocument? PdfDocument
        {
            get; set;
        }
        public System.IO.MemoryStream? FileStream
        {
            get; set;
        }

        public volatile bool IsDisposed = false;

        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        public ObservableCollection<PdfBookmarkViewModel> Bookmarks { get; set; } = new ObservableCollection<PdfBookmarkViewModel>();

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (IsDisposed)
                return;
            IsDisposed = true;

            if (disposing)
            {
                // [FIX] Lock 순서 변경: PdfiumLock을 먼저 잡고, 그 안에서 Dispatcher 호출
                lock (PdfService.PdfiumLock)
                {
                    try
                    {
                        if (Application.Current != null)
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    PdfDocument?.Dispose();
                                    PdfDocument = null;
                                    FileStream?.Close();
                                    FileStream?.Dispose();
                                    FileStream = null;
                                }
                                catch { }
                            });
                        }
                        else
                        {
                            // Application.Current가 null인 경우 (앱 종료 시)
                            PdfDocument?.Dispose();
                            PdfDocument = null;
                            FileStream?.Close();
                            FileStream?.Dispose();
                            FileStream = null;
                        }
                    }
                    catch { }
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public string? OriginalFilePath
        {
            get; set;
        }
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
            get => _rotation; set
            {
                _rotation = value;
                OnPropertyChanged(nameof(Rotation));
            }
        }

        public List<OcrWordInfo>? OcrWords
        {
            get; set;
        }
        public string MediaBoxInfo { get; set; } = "";
        public void Unload()
        {
            ImageSource = null;
        }
        public int PageIndex
        {
            get; set;
        }

        private double _width; public double Width
        {
            get => _width; set
            {
                _width = value;
                OnPropertyChanged(nameof(Width));
            }
        }
        private double _height; public double Height
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
        public bool AnnotationsLoaded
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

        // 템플릿 선택용 속성
        public bool IsFreeText => Type == AnnotationType.FreeText;
        public bool IsHighlight => Type == AnnotationType.Highlight || Type == AnnotationType.Underline || Type == AnnotationType.SearchHighlight;

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

    // [수정] PdfDocument 명확화
    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfSharpDoc document) : base(document) { }
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

        private bool _isEditing;
        public bool IsEditing
        {
            get => _isEditing; set
            {
                _isEditing = value;
                OnPropertyChanged(nameof(IsEditing));
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