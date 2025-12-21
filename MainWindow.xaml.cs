using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Text;
using System.Runtime.InteropServices.WindowsRuntime;

using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace MinsPDFViewer
{
    // [PDFSharp 6.x 호환용]
    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }

    // [시스템 폰트 리졸버]
    public class WindowsFontResolver : IFontResolver
    {
        public byte[]? GetFont(string faceName)
        {
            string fontPath = @"C:\Windows\Fonts\malgun.ttf";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\gulim.ttc";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\batang.ttc"; 
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\arial.ttf";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            return null;
        }
        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new FontResolverInfo("Malgun Gothic");
    }

    // 주석 데이터 모델
    public enum AnnotationType { Highlight, Underline, SearchHighlight, FreeText, Other }

    public class PdfAnnotation : INotifyPropertyChanged
    {
        private double _x; public double X { get => _x; set { _x = value; OnPropertyChanged(nameof(X)); } }
        private double _y; public double Y { get => _y; set { _y = value; OnPropertyChanged(nameof(Y)); } }
        private double _width; public double Width { get => _width; set { _width = value; OnPropertyChanged(nameof(Width)); } }
        private double _height; public double Height { get => _height; set { _height = value; OnPropertyChanged(nameof(Height)); } }

        private Brush _background = Brushes.Transparent;
        public Brush Background { get => _background; set { _background = value; OnPropertyChanged(nameof(Background)); } }

        private Brush _borderBrush = Brushes.Transparent;
        public Brush BorderBrush { get => _borderBrush; set { _borderBrush = value; OnPropertyChanged(nameof(BorderBrush)); } }

        private Thickness _borderThickness = new Thickness(0);
        public Thickness BorderThickness { get => _borderThickness; set { _borderThickness = value; OnPropertyChanged(nameof(BorderThickness)); } }

        private string _textContent = "";
        public string TextContent { get => _textContent; set { _textContent = value; OnPropertyChanged(nameof(TextContent)); OnPropertyChanged(nameof(HasText)); } }

        public bool HasText => !string.IsNullOrEmpty(TextContent) || Type == AnnotationType.FreeText;

        public AnnotationType Type { get; set; } = AnnotationType.Other;
        public Color AnnotationColor { get; set; } = Colors.Transparent;

        // 스타일 속성
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

        public bool IsFreeText => Type == AnnotationType.FreeText;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class OcrWordInfo { public string Text { get; set; } = ""; public Rect BoundingBox { get; set; } }
    
    // [수정] 필드 초기화로 경고 제거
    public class RawAnnotationInfo 
    { 
        public AnnotationType Type; 
        public XRect Rect; 
        public string Content = ""; 
        public double FontSize; 
        public Color Color; 
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        private ImageSource? _imageSource;
        public ImageSource? ImageSource { get => _imageSource; set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); } }
        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();
        public List<OcrWordInfo> OcrWords { get; set; } = new List<OcrWordInfo>();
        public double PdfPageWidthPoint { get; set; }
        public double PdfPageHeightPoint { get; set; }
        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }
        private double _selX; public double SelectionX { get => _selX; set { _selX = value; OnPropertyChanged(nameof(SelectionX)); } }
        private double _selY; public double SelectionY { get => _selY; set { _selY = value; OnPropertyChanged(nameof(SelectionY)); } }
        private double _selW; public double SelectionWidth { get => _selW; set { _selW = value; OnPropertyChanged(nameof(SelectionWidth)); } }
        private double _selH; public double SelectionHeight { get => _selH; set { _selH = value; OnPropertyChanged(nameof(SelectionHeight)); } }
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();

        private double _renderScale = 2.0;
        private double _currentZoom = 1.0;
        private string _currentFilePath = "";
        private byte[] _originalFileBytes = Array.Empty<byte>();
        private OcrEngine? _ocrEngine;

        private string _currentTool = "CURSOR";
        private PdfAnnotation? _selectedAnnotation = null;

        private string _defaultFontFamily = "Malgun Gothic"; 
        private double _defaultFontSize = 14;
        private Color _defaultFontColor = Colors.Red;
        private bool _defaultIsBold = false;

        private Point _dragStartPoint;
        private int _activePageIndex = -1;
        private string _selectedTextBuffer = "";
        private List<Docnet.Core.Models.Character> _selectedChars = new List<Docnet.Core.Models.Character>();
        private int _selectedPageIndex = -1;

        private bool _isDraggingAnnotation = false;
        private Point _annotationDragStartOffset;

        private List<PdfAnnotation> _searchResults = new List<PdfAnnotation>();
        private int _currentSearchIndex = -1;
        private string _lastSearchQuery = "";
        
        private bool _isUpdatingUiFromSelection = false;

        public MainWindow()
        {
            InitializeComponent();
            try { if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new WindowsFontResolver(); } catch { }
            _docLib = DocLib.Instance;
            PdfListView.ItemsSource = Pages;

            try { _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR")) ?? OcrEngine.TryCreateFromUserProfileLanguages(); } catch { }

            CbFont.ItemsSource = new string[] { "Malgun Gothic", "Gulim", "Dotum", "Batang" };
            CbFont.SelectedIndex = 0;
            CbSize.ItemsSource = new double[] { 10, 12, 14, 16, 18, 24, 32, 48 };
            CbSize.SelectedIndex = 2;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e) { var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" }; if (dlg.ShowDialog() == true) LoadPdf(dlg.FileName); }

        private void CheckToolbarVisibility()
        {
            bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null && _selectedAnnotation.IsFreeText);
            if (TextStyleToolbar != null)
                TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed;
        }

        // [헬퍼] PdfItem에서 숫자 값 추출 (Real 또는 Integer)
        private double GetPdfNumber(PdfItem item)
        {
            if (item is PdfReal r) return r.Value;
            if (item is PdfInteger i) return i.Value;
            return 0.0;
        }

        private void LoadPdf(string path)
        {
            try
            {
                _currentFilePath = path;
                _originalFileBytes = File.ReadAllBytes(path);
                Pages.Clear();

                var extractedRawData = new Dictionary<int, List<RawAnnotationInfo>>();
                var pdfPageSizes = new Dictionary<int, XSize>(); 

                using (var msInput = new MemoryStream(_originalFileBytes))
                using (var doc = PdfReader.Open(msInput, PdfDocumentOpenMode.Modify))
                {
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        pdfPageSizes[i] = new XSize(page.Width.Point, page.Height.Point);
                        extractedRawData[i] = new List<RawAnnotationInfo>();

                        if (page.Annotations != null)
                        {
                            var annotsToRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for (int k = 0; k < page.Annotations.Count; k++)
                            {
                                var ann = page.Annotations[k];
                                var subtype = ann.Elements.GetString("/Subtype");

                                if (subtype == "/FreeText")
                                {
                                    var rect = ann.Rectangle.ToXRect();
                                    string da = ann.Elements.GetString("/DA");
                                    double fSize = 14; Color fColor = Colors.Red; 
                                    if (!string.IsNullOrEmpty(da))
                                    {
                                        var parts = da.Split(' ');
                                        for (int p = 0; p < parts.Length; p++)
                                        {
                                            if (parts[p] == "Tf" && p >= 2) { if (double.TryParse(parts[p-1], out double s)) fSize = s; }
                                            else if (parts[p] == "rg" && p >= 3)
                                            {
                                                double.TryParse(parts[p-3], out double r); double.TryParse(parts[p-2], out double g); double.TryParse(parts[p-1], out double b);
                                                fColor = Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255));
                                            }
                                        }
                                    }
                                    extractedRawData[i].Add(new RawAnnotationInfo { Type = AnnotationType.FreeText, Rect = rect, Content = ann.Contents, FontSize = fSize, Color = fColor });
                                    annotsToRemove.Add(ann); 
                                }
                                else if (subtype == "/Highlight" || subtype == "/Underline")
                                {
                                    var rect = ann.Rectangle.ToXRect();
                                    Color cColor = Colors.Yellow; 
                                    var cArray = ann.Elements.GetArray("/C");
                                    
                                    // [수정] PdfArray 요소 접근 시 타입 캐스팅 방식 변경 (PdfNumber -> PdfReal/PdfInteger)
                                    if (cArray != null && cArray.Elements.Count >= 3)
                                    {
                                        // 안전하게 값 추출: Real이 아니면 Integer 시도, 없으면 기본값
                                        double r = (cArray.Elements[0] as PdfReal)?.Value ?? (cArray.Elements[0] as PdfInteger)?.Value ?? 1.0;
                                        double g = (cArray.Elements[1] as PdfReal)?.Value ?? (cArray.Elements[1] as PdfInteger)?.Value ?? 1.0;
                                        double b = (cArray.Elements[2] as PdfReal)?.Value ?? (cArray.Elements[2] as PdfInteger)?.Value ?? 0.0;
                                        
                                        cColor = Color.FromRgb((byte)(r * 255), (byte)(g * 255), (byte)(b * 255));
                                    }
                                    
                                    AnnotationType type = (subtype == "/Highlight") ? AnnotationType.Highlight : AnnotationType.Underline;
                                    extractedRawData[i].Add(new RawAnnotationInfo { Type = type, Rect = rect, Color = cColor });
                                    annotsToRemove.Add(ann);
                                }
                            }
                            foreach (var item in annotsToRemove) page.Annotations.Remove(item);
                        }
                    }
                    var cleanStream = new MemoryStream(); doc.Save(cleanStream);
                    _docReader?.Dispose(); _docReader = _docLib.GetDocReader(cleanStream.ToArray(), new PageDimensions(_renderScale));
                }

                int pc = _docReader.GetPageCount();
                for (int i = 0; i < pc; i++)
                {
                    using (var r = _docReader.GetPageReader(i))
                    {
                        double viewW = r.GetPageWidth(); double viewH = r.GetPageHeight();
                        var pvm = new PdfPageViewModel { 
                            PageIndex = i, Width = viewW, Height = viewH,
                            PdfPageWidthPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Width : viewW,
                            PdfPageHeightPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Height : viewH
                        };
                        double scaleX = viewW / pvm.PdfPageWidthPoint; double scaleY = viewH / pvm.PdfPageHeightPoint;

                        if (extractedRawData.ContainsKey(i))
                        {
                            foreach (var raw in extractedRawData[i])
                            {
                                double pdfTopY = raw.Rect.Y + raw.Rect.Height; 
                                double viewY = (pvm.PdfPageHeightPoint - pdfTopY) * scaleY;
                                
                                var ann = new PdfAnnotation {
                                    Type = raw.Type, X = raw.Rect.X * scaleX, Y = viewY, Width = raw.Rect.Width * scaleX, Height = raw.Rect.Height * scaleY,
                                    TextContent = raw.Content, FontSize = raw.FontSize, FontFamily = "Malgun Gothic", Foreground = new SolidColorBrush(raw.Color), AnnotationColor = raw.Color
                                };

                                if (raw.Type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, raw.Color.R, raw.Color.G, raw.Color.B));
                                else if (raw.Type == AnnotationType.Underline) { ann.Background = new SolidColorBrush(raw.Color); ann.Y = viewY + (raw.Rect.Height * scaleY) - 2; ann.Height = 2; }
                                else ann.Background = Brushes.Transparent;

                                pvm.Annotations.Add(ann);
                            }
                        }
                        Pages.Add(pvm);
                    }
                }
                Task.Run(() => RenderAllPagesAsync(pc)); FitWidth();
            }
            catch (Exception ex) { MessageBox.Show($"열기 실패: {ex.Message}"); }
        }

        private void RenderAllPagesAsync(int pageCount)
        {
            Parallel.For(0, pageCount, new ParallelOptions { MaxDegreeOfParallelism = 4 }, i =>
            {
                if (_docReader == null) return;
                using (var r = _docReader.GetPageReader(i))
                {
                    var bytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight();
                    Application.Current.Dispatcher.Invoke(() => { if (i < Pages.Count) Pages[i].ImageSource = RawBytesToBitmapImage(bytes, w, h); });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        private void SavePdf(string savePath)
        {
            try
            {
                using (var ms = new MemoryStream(_originalFileBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    if (doc.Version < 14) doc.Version = 14;
                    for (int i = 0; i < doc.PageCount && i < Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i]; var pageVM = Pages[i];
                        double scaleX = pdfPage.Width.Point / pageVM.Width; double scaleY = pdfPage.Height.Point / pageVM.Height;
                        var annotsToRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                        if (pdfPage.Annotations != null) {
                            for (int k = 0; k < pdfPage.Annotations.Count; k++) {
                                var a = pdfPage.Annotations[k]; 
                                string st = a.Elements.GetString("/Subtype");
                                if (st == "/FreeText" || st == "/Highlight" || st == "/Underline") annotsToRemove.Add(a);
                            }
                            foreach (var a in annotsToRemove) pdfPage.Annotations.Remove(a);
                        }
                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0) {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage)) {
                                foreach (var word in pageVM.OcrWords) {
                                    double x = word.BoundingBox.X * scaleX; double y = word.BoundingBox.Y * scaleY; double w = word.BoundingBox.Width * scaleX; double h = word.BoundingBox.Height * scaleY;
                                    double fSize = h * 0.75; if(fSize < 1) fSize = 1;
                                    gfx.DrawString(word.Text, new XFont("Malgun Gothic", fSize), XBrushes.Transparent, new XRect(x, y, w, h), XStringFormats.Center);
                                }
                            }
                        }
                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;
                            double ax = ann.X * scaleX; double ay = ann.Y * scaleY; double aw = ann.Width * scaleX; double ah = ann.Height * scaleY;
                            double pdfY_BottomUp = pdfPage.Height.Point - (ay + ah);

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                using (var gfx = XGraphics.FromPdfPage(pdfPage)) {
                                    var font = new XFont("Malgun Gothic", ann.FontSize, XFontStyleEx.Regular);
                                    var size = gfx.MeasureString(ann.TextContent, font);
                                    int lineCount = ann.TextContent.Split('\n').Length;
                                    double estimatedHeight = size.Height * lineCount * 1.2; 
                                    if (estimatedHeight > ah) { double diff = estimatedHeight - ah; ah = estimatedHeight; pdfY_BottomUp -= diff; }
                                }
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                string da = $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg";
                                pdfAnnot.Elements.SetString("/DA", da);
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else 
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                if (ann.Type == AnnotationType.Underline) pdfY_BottomUp = pdfPage.Height.Point - (ay + 2);
                                var rect = new XRect(ax, pdfY_BottomUp, aw, ah);
                                pdfAnnot.Rectangle = new PdfRectangle(rect);
                                pdfAnnot.Elements.SetName("/Subtype", subtype);
                                double qLeft = rect.X; double qRight = rect.X + rect.Width; double qBottom = rect.Y; double qTop = rect.Y + rect.Height;
                                var quadPoints = new PdfArray(doc, new PdfReal(qLeft), new PdfReal(qTop), new PdfReal(qRight), new PdfReal(qTop), new PdfReal(qLeft), new PdfReal(qBottom), new PdfReal(qRight), new PdfReal(qBottom));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);
                                double r = ann.AnnotationColor.R / 255.0; double g = ann.AnnotationColor.G / 255.0; double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }

        private void Page_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; }

            var canvas = sender as Canvas; if (canvas == null) return;
            _activePageIndex = (int)canvas.Tag; _dragStartPoint = e.GetPosition(canvas);

            if (_currentTool == "TEXT")
            {
                var pageVM = Pages[_activePageIndex];
                var newAnnot = new PdfAnnotation {
                    Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50,
                    FontSize = _defaultFontSize, FontFamily = _defaultFontFamily,
                    Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold,
                    TextContent = "", // 입력 유도
                    IsSelected = true
                };
                pageVM.Annotations.Add(newAnnot);
                _selectedAnnotation = newAnnot;
                
                _currentTool = "CURSOR"; RbCursor.IsChecked = true;
                UpdateToolbarFromAnnotation(_selectedAnnotation);
                CheckToolbarVisibility(); 
                e.Handled = true;
                return;
            }

            if (_currentTool == "CURSOR")
            {
                foreach (var p in Pages) { p.IsSelecting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                SelectionPopup.IsOpen = false;
                var pageVM = Pages[_activePageIndex];
                pageVM.IsSelecting = true;
                pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y;
                pageVM.SelectionWidth = 0; pageVM.SelectionHeight = 0;
                canvas.CaptureMouse();
                e.Handled = true;
            }
            CheckToolbarVisibility();
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1) return; var canvas = sender as Canvas; if (canvas == null) return;
            if (_currentTool == "CURSOR" && Pages[_activePageIndex].IsSelecting) {
                var pt = e.GetPosition(canvas); var p = Pages[_activePageIndex];
                double x = Math.Min(_dragStartPoint.X, pt.X); double y = Math.Min(_dragStartPoint.Y, pt.Y);
                double w = Math.Abs(pt.X - _dragStartPoint.X); double h = Math.Abs(pt.Y - _dragStartPoint.Y);
                p.SelectionX = x; p.SelectionY = y; p.SelectionWidth = w; p.SelectionHeight = h;
            }
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var canvas = sender as Canvas; if (canvas == null || _activePageIndex == -1) return;
            canvas.ReleaseMouseCapture(); var p = Pages[_activePageIndex];
            
            if (p.IsSelecting && _currentTool == "CURSOR") {
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5) {
                    var rect = new Rect(p.SelectionX, p.SelectionY, p.SelectionWidth, p.SelectionHeight);
                    CheckTextInSelection(_activePageIndex, rect);
                    SelectionPopup.PlacementTarget = canvas; SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(canvas).X, e.GetPosition(canvas).Y + 10, 0, 0); SelectionPopup.IsOpen = true;
                    _selectedPageIndex = _activePageIndex;
                    if (_selectedChars.Count > 0 || !string.IsNullOrEmpty(_selectedTextBuffer)) TxtStatus.Text = "텍스트 선택됨"; else TxtStatus.Text = "영역 선택됨";
                } else { 
                    p.IsSelecting = false; TxtStatus.Text = "준비"; 
                    if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; }
                }
            }
            _activePageIndex = -1; e.Handled = true;
            CheckToolbarVisibility(); 
        }

        private void Annotation_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool != "CURSOR") return;
            var element = sender as FrameworkElement;
            if (element?.DataContext is PdfAnnotation ann && ann.Type == AnnotationType.FreeText) {
                if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true;
                _isDraggingAnnotation = true; _annotationDragStartOffset = e.GetPosition(element);
                UpdateToolbarFromAnnotation(ann); 
                CheckToolbarVisibility(); 
            }
        }
        private void Annotation_PreviewMouseMove(object sender, MouseEventArgs e) { }
        private void Annotation_PreviewMouseUp(object sender, MouseButtonEventArgs e) {
            if (_isDraggingAnnotation) { _isDraggingAnnotation = false; (sender as FrameworkElement)?.ReleaseMouseCapture(); }
        }
        
        private void AnnotationTextBox_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann && ann.IsSelected) {
                tb.Focus();
                CheckToolbarVisibility();
            }
        }
        private void AnnotationTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann) {
                _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true;
                UpdateToolbarFromAnnotation(ann);
                CheckToolbarVisibility();
            }
        }

        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e) {
            var thumb = sender as Thumb; if (thumb?.DataContext is PdfAnnotation ann) { ann.Width = Math.Max(50, ann.Width + e.HorizontalChange); ann.Height = Math.Max(30, ann.Height + e.VerticalChange); }
        }

        private void UpdateToolbarFromAnnotation(PdfAnnotation ann)
        {
            _isUpdatingUiFromSelection = true;
            try {
                CbFont.SelectedItem = ann.FontFamily; CbSize.SelectedItem = ann.FontSize; BtnBold.IsChecked = ann.IsBold;
                if (ann.Foreground is SolidColorBrush brush) {
                    foreach (ComboBoxItem item in CbColor.Items) {
                        string cn = item.Tag?.ToString() ?? ""; Color c = Colors.Black;
                        if (cn == "Red") c = Colors.Red; else if (cn == "Blue") c = Colors.Blue; else if (cn == "Green") c = Colors.Green; else if (cn == "Orange") c = Colors.Orange;
                        if (brush.Color == c) { CbColor.SelectedItem = item; break; }
                    }
                }
            } finally { _isUpdatingUiFromSelection = false; }
        }

        private void StyleChanged(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _isUpdatingUiFromSelection) return;
            if (CbFont.SelectedItem != null) _defaultFontFamily = CbFont.SelectedItem.ToString();
            if (CbSize.SelectedItem != null) _defaultFontSize = (double)CbSize.SelectedItem;
            _defaultIsBold = BtnBold.IsChecked == true;
            if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null) {
                string cn = item.Tag.ToString();
                if (cn == "Black") _defaultFontColor = Colors.Black; else if (cn == "Red") _defaultFontColor = Colors.Red; else if (cn == "Blue") _defaultFontColor = Colors.Blue; else if (cn == "Green") _defaultFontColor = Colors.Green; else if (cn == "Orange") _defaultFontColor = Colors.Orange;
            }
            if (_selectedAnnotation != null) {
                _selectedAnnotation.FontFamily = _defaultFontFamily; _selectedAnnotation.FontSize = _defaultFontSize;
                _selectedAnnotation.IsBold = _defaultIsBold; _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor);
            }
        }

        private void BtnDeleteAnnot_Click(object sender, RoutedEventArgs e) {
            if (_selectedAnnotation != null) { foreach(var p in Pages) { if (p.Annotations.Contains(_selectedAnnotation)) { p.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; CheckToolbarVisibility(); break; } } }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; }
            // [수정] 더미 객체 전달로 경고 회피
            else if (e.Key == Key.Delete) BtnDeleteAnnot_Click(this, new RoutedEventArgs());
            else if (e.Key == Key.Escape) { 
                if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; }
                CheckToolbarVisibility();
            }
        }
        
        private void Tool_Click(object sender, RoutedEventArgs e) { 
            if (RbCursor.IsChecked == true) _currentTool = "CURSOR"; 
            else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT"; 
            else if (RbText.IsChecked == true) _currentTool = "TEXT"; 
            CheckToolbarVisibility();
        }

        private void BtnOCR_Click(object sender, RoutedEventArgs e) { if (_ocrEngine == null) { MessageBox.Show("OCR 미지원"); return; } if (Pages.Count == 0) return; BtnOCR.IsEnabled = false; PbStatus.Visibility = Visibility.Visible; PbStatus.Maximum = Pages.Count; PbStatus.Value = 0; TxtStatus.Text = "OCR 분석 중..."; Task.Run(async () => { try { for (int i = 0; i < Pages.Count; i++) { var pageVM = Pages[i]; if (_docReader == null) continue; using (var r = _docReader.GetPageReader(i)) { var rawBytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight(); using (var stream = new MemoryStream(rawBytes)) { var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied); softwareBitmap.CopyFromBuffer(rawBytes.AsBuffer()); var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap); var wordList = new List<OcrWordInfo>(); foreach (var line in ocrResult.Lines) { foreach (var word in line.Words) { wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(word.BoundingRect.X, word.BoundingRect.Y, word.BoundingRect.Width, word.BoundingRect.Height) }); } } pageVM.OcrWords = wordList; } } Application.Current.Dispatcher.Invoke(() => { PbStatus.Value = i + 1; }); } Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "OCR 완료. 저장하세요."); } catch (Exception ex) { Application.Current.Dispatcher.Invoke(() => MessageBox.Show($"OCR 오류: {ex.Message}")); } finally { Application.Current.Dispatcher.Invoke(() => { BtnOCR.IsEnabled = true; PbStatus.Visibility = Visibility.Collapsed; }); } }); }
        private void CheckTextInSelection(int pageIndex, Rect uiRect) { _selectedTextBuffer = ""; _selectedChars.Clear(); if (_docReader == null) return; var sb = new StringBuilder(); using (var reader = _docReader.GetPageReader(pageIndex)) { var chars = reader.GetCharacters().ToList(); foreach (var c in chars) { var r = new Rect(Math.Min(c.Box.Left, c.Box.Right), Math.Min(c.Box.Top, c.Box.Bottom), Math.Abs(c.Box.Right - c.Box.Left), Math.Abs(c.Box.Bottom - c.Box.Top)); if (uiRect.IntersectsWith(r)) { _selectedChars.Add(c); sb.Append(c.Char); } } } var pageVM = Pages[pageIndex]; if (pageVM.OcrWords != null) { foreach (var word in pageVM.OcrWords) { if (uiRect.IntersectsWith(word.BoundingBox)) sb.Append(word.Text + " "); } } _selectedTextBuffer = sb.ToString(); }
        private void PerformSearch(string query) { if (string.IsNullOrWhiteSpace(query)) return; _lastSearchQuery = query; _searchResults.Clear(); foreach (var p in Pages) { var tr = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList(); foreach (var i in tr) p.Annotations.Remove(i); } for (int i = 0; i < Pages.Count; i++) { var p = Pages[i]; if (p.OcrWords != null) { foreach (var word in p.OcrWords) { if (word.Text.Contains(query, StringComparison.OrdinalIgnoreCase)) { var a = new PdfAnnotation { X = word.BoundingBox.X, Y = word.BoundingBox.Y, Width = word.BoundingBox.Width, Height = word.BoundingBox.Height, Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)), Type = AnnotationType.SearchHighlight }; p.Annotations.Add(a); _searchResults.Add(a); } } } if (_docReader != null) { using (var r = _docReader.GetPageReader(i)) { string t = r.GetText(); var cs = r.GetCharacters().ToList(); int idx = 0; while ((idx = t.IndexOf(query, idx, StringComparison.OrdinalIgnoreCase)) != -1) { double mx = double.MaxValue, my = double.MaxValue, Mx = double.MinValue, My = double.MinValue; for (int c = 0; c < query.Length; c++) { if (idx + c < cs.Count) { var b = cs[idx + c].Box; mx = Math.Min(mx, b.Left); my = Math.Min(my, b.Top); Mx = Math.Max(Mx, b.Right); My = Math.Max(My, b.Bottom); } } var a = new PdfAnnotation { X = mx, Y = my, Width = Mx - mx, Height = My - my, Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)), Type = AnnotationType.SearchHighlight }; Pages[i].Annotations.Add(a); _searchResults.Add(a); idx += query.Length; } } } } TxtStatus.Text = $"검색: {_searchResults.Count}건"; if (_searchResults.Count > 0) { _currentSearchIndex = -1; NavigateSearchResult(true); } }
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) { if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift) NavigateSearchResult(false); else { string query = TxtSearch.Text; if (query == _lastSearchQuery && _searchResults.Count > 0) NavigateSearchResult(true); else PerformSearch(query); } e.Handled = true; } }
        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (!string.IsNullOrEmpty(_currentFilePath)) SavePdf(_currentFilePath); }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) { if (string.IsNullOrEmpty(_currentFilePath)) return; var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(_currentFilePath) + "_copy" }; if (dlg.ShowDialog() == true) { SavePdf(dlg.FileName); _currentFilePath = dlg.FileName; } }
        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e) { if (!string.IsNullOrEmpty(_selectedTextBuffer)) { Clipboard.SetText(_selectedTextBuffer); TxtStatus.Text = "텍스트 복사됨"; } CloseSelection(); }
        private void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e) { if (_selectedPageIndex == -1) return; var p = Pages[_selectedPageIndex]; var bmp = p.ImageSource as BitmapSource; if (bmp == null) return; double sx = bmp.PixelWidth / p.Width; double sy = bmp.PixelHeight / p.Height; int x = (int)(p.SelectionX * sx); int y = (int)(p.SelectionY * sy); int w = (int)(p.SelectionWidth * sx); int h = (int)(p.SelectionHeight * sy); if (w > 0 && h > 0 && x >= 0 && y >= 0 && x + w <= bmp.PixelWidth && y + h <= bmp.PixelHeight) { try { Clipboard.SetImage(new CroppedBitmap(bmp, new Int32Rect(x, y, w, h))); TxtStatus.Text = "이미지 복사됨"; } catch { TxtStatus.Text = "복사 실패"; } } CloseSelection(); }
        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);
        private void AddAnnotation(Color color, AnnotationType type) { if (_selectedPageIndex == -1) return; var p = Pages[_selectedPageIndex]; var ann = new PdfAnnotation { X = p.SelectionX, Y = p.SelectionY, Width = p.SelectionWidth, Height = p.SelectionHeight, Type = type, AnnotationColor = color }; if (type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B)); else { ann.Background = new SolidColorBrush(color); ann.Height = 2; ann.Y = p.SelectionY + p.SelectionHeight - 2; } p.Annotations.Add(ann); CloseSelection(); }
        private void CloseSelection() { SelectionPopup.IsOpen = false; foreach (var p in Pages) p.IsSelecting = false; }
        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e) { if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a) foreach (var p in Pages) if (p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; } }
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(false);
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(true);
        private void NavigateSearchResult(bool n) { if (_searchResults.Count == 0) return; if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count) _searchResults[_currentSearchIndex].Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)); if (n) { _currentSearchIndex++; if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0; } else { _currentSearchIndex--; if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1; } var c = _searchResults[_currentSearchIndex]; c.Background = new SolidColorBrush(Color.FromArgb(120, 255, 0, 255)); var tp = Pages.FirstOrDefault(p => p.Annotations.Contains(c)); if (tp != null) PdfListView.ScrollIntoView(tp); TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}"; }
        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { if (e.Delta > 0) UpdateZoom(_currentZoom + 0.1); else UpdateZoom(_currentZoom - 0.1); e.Handled = true; } }
        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom + 0.1);
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom - 0.1);
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) => FitWidth();
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) => FitHeight();
        private void UpdateZoom(double z) { _currentZoom = Math.Clamp(z, 0.2, 5.0); ViewScaleTransform.ScaleX = _currentZoom; ViewScaleTransform.ScaleY = _currentZoom; TxtZoom.Text = $"{Math.Round(_currentZoom * 100)}%"; }
        private void FitWidth() { if (Pages.Count > 0 && Pages[0].Width > 0) UpdateZoom((PdfListView.ActualWidth - 60) / Pages[0].Width); }
        private void FitHeight() { if (Pages.Count > 0 && Pages[0].Height > 0) UpdateZoom((PdfListView.ActualHeight - 60) / Pages[0].Height); }
        private BitmapImage RawBytesToBitmapImage(byte[] b, int w, int h) { var bm = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null); bm.WritePixels(new Int32Rect(0, 0, w, h), b, w * 4, 0); if (bm.CanFreeze) bm.Freeze(); return ConvertWriteableBitmapToBitmapImage(bm); }
        private BitmapImage ConvertWriteableBitmapToBitmapImage(WriteableBitmap wbm) { using (var ms = new MemoryStream()) { var enc = new PngBitmapEncoder(); enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(wbm)); enc.Save(ms); var img = new BitmapImage(); img.BeginInit(); img.CacheOption = BitmapCacheOption.OnLoad; img.StreamSource = ms; img.EndInit(); if (img.CanFreeze) img.Freeze(); return img; } }
    }
}