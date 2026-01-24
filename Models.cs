using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading;
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

        public async Task CloseAsync()
        {
            if (IsDisposed) return;
            IsDisposed = true;

            await Task.Run(() =>
            {
                lock (PdfService.PdfiumLock)
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
                }
            });
        }

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
                // [FIX] Avoid blocking UI thread or deadlocks during Dispose
                // Do not lock PdfiumLock here if possible, or ensure it's safe.
                // Pdfium operations must be locked, but Dispose of PdfDocument might need it.
                
                // Offload to background or handle carefully. 
                // Since PdfDocument.Dispose() might access native resources, we lock.
                
                Action disposeAction = () =>
                {
                    lock (PdfService.PdfiumLock)
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
                    }
                };

                if (Application.Current != null)
                {
                    if (Application.Current.Dispatcher.CheckAccess())
                    {
                         // Already on UI thread, just run (but be careful of lock)
                         // Ideally, dispose native resources in background?
                         // PdfiumViewer.PdfDocument.Dispose() is not thread-bound usually.
                         // Let's run it directly or in Task.Run to avoid UI freeze if lock is held by renderer.
                         Task.Run(disposeAction);
                    }
                    else
                    {
                        disposeAction();
                    }
                }
                else
                {
                    disposeAction();
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
            CancelRender();
            ImageSource = null;
            IsHighResRendered = false;
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

        public bool IsHighResRendered { get; set; }

        private CancellationTokenSource? _renderCts;
        public CancellationTokenSource? RenderCts
        {
            get => _renderCts;
            set => _renderCts = value;
        }

        public void CancelRender()
        {
            _renderCts?.Cancel();
            _renderCts?.Dispose();
            _renderCts = null;
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

        private bool _isSelecting;
        public bool IsSelecting
        {
            get => _isSelecting;
            set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); }
        }

        private bool _isHighlighting;
        public bool IsHighlighting
        {
            get => _isHighlighting;
            set { _isHighlighting = value; OnPropertyChanged(nameof(IsHighlighting)); }
        }

        private double _selectionX;
        public double SelectionX
        {
            get => _selectionX;
            set { _selectionX = value; OnPropertyChanged(nameof(SelectionX)); }
        }

        private double _selectionY;
        public double SelectionY
        {
            get => _selectionY;
            set { _selectionY = value; OnPropertyChanged(nameof(SelectionY)); }
        }

        private double _selectionWidth;
        public double SelectionWidth
        {
            get => _selectionWidth;
            set { _selectionWidth = value; OnPropertyChanged(nameof(SelectionWidth)); }
        }

        private double _selectionHeight;
        public double SelectionHeight
        {
            get => _selectionHeight;
            set { _selectionHeight = value; OnPropertyChanged(nameof(SelectionHeight)); }
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

        private AnnotationType _type;
        public AnnotationType Type
        {
            get => _type;
            set
            {
                _type = value;
                OnPropertyChanged(nameof(Type));
                OnPropertyChanged(nameof(IsFreeText));
                OnPropertyChanged(nameof(IsHighlight));
            }
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
                OnPropertyChanged(nameof(PageNumber));
            }
        }

        public int PageNumber => PageIndex + 1;

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