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

        // [추가] 리소스 정리
        public void Dispose()
        {
            DocReader?.Dispose();
            CleanDocReader?.Dispose();
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
        public bool IsSelecting
        {
            get; set;
        } // Notify 구현 생략 (기존 유지)
        public double SelectionX
        {
            get; set;
        } // Notify 구현 생략
        public double SelectionY
        {
            get; set;
        } // Notify 구현 생략
        public double SelectionWidth
        {
            get; set;
        } // Notify 구현 생략
        public double SelectionHeight
        {
            get; set;
        } // Notify 구현 생략
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
    public class PdfAnnotation : INotifyPropertyChanged
    {
        // 기존 내용 유지 (FieldName 포함)
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
        public string TextContent { get; set; } = ""; // Notify 생략
        public double FontSize { get; set; } = 12; // Notify 생략
        public string FontFamily { get; set; } = "Malgun Gothic"; // Notify 생략
        public Brush Foreground { get; set; } = Brushes.Black; // Notify 생략
        public bool IsBold
        {
            get; set;
        }
        public string FieldName { get; set; } = "";
        public object? SignatureData
        {
            get; set;
        }
        public string? VisualStampPath
        {
            get; set;
        } // Notify 생략
        public bool IsSelected
        {
            get; set;
        } // Notify 생략

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