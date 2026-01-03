using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Media;
using Docnet.Core.Readers; // 추가
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

    public class PdfDocumentModel : INotifyPropertyChanged, IDisposable
    {
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
        public IDocReader? DocReader
        {
            get; set;
        }      // 원본 (텍스트/검색용)
        public IDocReader? CleanDocReader
        {
            get; set;
        } // [추가] 렌더링용 (주석 제거됨)

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // [핵심 수정] 리소스 해제 시 충돌 방지
        public void Dispose()
        {
            // PdfService의 전역 락을 사용하여, 렌더링 중인 스레드와 충돌 방지
            lock (PdfService.PdfiumLock)
            {
                if (DocReader != null)
                {
                    DocReader.Dispose();
                    DocReader = null;
                }

                if (CleanDocReader != null)
                {
                    CleanDocReader.Dispose();
                    CleanDocReader = null;
                }
            }

            // DocLib는 싱글톤(Instance)이므로 여기서 Dispose하지 않음 (프로그램 종료 시 자동 해제됨)
        }

        private bool _isSelecting;
        public bool IsSelecting
        {
            get => _isSelecting;
            set
            {
                _isSelecting = value;
                OnPropertyChanged(nameof(IsSelecting));
            }
        }

        private double _selectionX;
        public double SelectionX
        {
            get => _selectionX;
            set
            {
                _selectionX = value;
                OnPropertyChanged(nameof(SelectionX));
            }
        }

        private double _selectionY;
        public double SelectionY
        {
            get => _selectionY;
            set
            {
                _selectionY = value;
                OnPropertyChanged(nameof(SelectionY));
            }
        }

        private double _selectionWidth;
        public double SelectionWidth
        {
            get => _selectionWidth;
            set
            {
                _selectionWidth = value;
                OnPropertyChanged(nameof(SelectionWidth));
            }
        }

        private double _selectionHeight;
        public double SelectionHeight
        {
            get => _selectionHeight;
            set
            {
                _selectionHeight = value;
                OnPropertyChanged(nameof(SelectionHeight));
            }
        }

        // [추가] 책갈피 목록 (트리 구조의 최상위 노드들)
        public ObservableCollection<PdfBookmarkViewModel> Bookmarks { get; set; } = new ObservableCollection<PdfBookmarkViewModel>();

    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {

        // [추가] 디버깅용 회전 및 원본 크기 정보
        public int Rotation
        {
            get; set;
        }
        public string MediaBoxInfo { get; set; } = "";

        // ... (기존 속성들: PageIndex, Width, Height 등 유지) ...
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

        private ImageSource? _imageSource;
        public ImageSource? ImageSource
        {
            get => _imageSource;
            set
            {
                _imageSource = value;
                OnPropertyChanged(nameof(ImageSource));
            }
        }

        // ... (기존 Annotations, CropX, 등등 속성 유지) ...
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
            get => _isSelecting;
            set
            {
                _isSelecting = value;
                OnPropertyChanged(nameof(IsSelecting));
            }
        }

        private double _selectionX;
        public double SelectionX
        {
            get => _selectionX;
            set
            {
                _selectionX = value;
                OnPropertyChanged(nameof(SelectionX));
            }
        }

        private double _selectionY;
        public double SelectionY
        {
            get => _selectionY;
            set
            {
                _selectionY = value;
                OnPropertyChanged(nameof(SelectionY));
            }
        }

        private double _selectionWidth;
        public double SelectionWidth
        {
            get => _selectionWidth;
            set
            {
                _selectionWidth = value;
                OnPropertyChanged(nameof(SelectionWidth));
            }
        }

        private double _selectionHeight;
        public double SelectionHeight
        {
            get => _selectionHeight;
            set
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

        // [추가] 메모리 해제 메서드
        public void Unload()
        {
            ImageSource = null; // 이미지 메모리 해제
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ... (PdfAnnotation, OcrWordInfo, GenericPdfAnnotation 클래스는 기존과 동일 유지) ...
    // Models.cs -> PdfAnnotation 클래스 (전체 교체)

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

        // [수정] 아래 속성들도 화면 갱신을 위해 OnPropertyChanged 추가
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

        // [중요] 선택 상태 알림 복구 (서명/텍스트 이동 핸들 표시에 필수)
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
        public string Text { get; set; } = ""; public System.Windows.Rect BoundingBox
        {
            get; set;
        }
    }
    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }
}

// [신규 클래스] 책갈피 뷰모델 (파일 맨 아래에 추가)
public class PdfBookmarkViewModel : INotifyPropertyChanged
{
    private string _title = "";
    public string Title
    {
        get => _title;
        set
        {
            _title = value;
            OnPropertyChanged(nameof(Title));
        }
    }

    private int _pageIndex;
    public int PageIndex
    {
        get => _pageIndex;
        set
        {
            _pageIndex = value;
            OnPropertyChanged(nameof(PageIndex));
        }
    }

    private bool _isExpanded = true; // 기본적으로 펼침
    public bool IsExpanded
    {
        get => _isExpanded;
        set
        {
            _isExpanded = value;
            OnPropertyChanged(nameof(IsExpanded));
        }
    }

    private bool _isSelected;
    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            _isSelected = value;
            OnPropertyChanged(nameof(IsSelected));
        }
    }

    // 하위 책갈피 (계층 구조)
    public ObservableCollection<PdfBookmarkViewModel> Children { get; set; } = new ObservableCollection<PdfBookmarkViewModel>();

    // 부모 참조 (트리 이동/삭제 시 필요)
    public PdfBookmarkViewModel? Parent
    {
        get; set;
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}