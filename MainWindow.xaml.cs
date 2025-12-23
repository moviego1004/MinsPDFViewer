using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Drawing;
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

// WinRT APIs for OCR
using Windows.Security.Cryptography;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace MinsPDFViewer
{
    public class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfDocument document) : base(document) { }
    }

    public class WindowsFontResolver : PdfSharp.Fonts.IFontResolver
    {
        public byte[]? GetFont(string faceName)
        {
            string fontPath = @"C:\Windows\Fonts\malgun.ttf";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\gulim.ttc";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            return null;
        }
        public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new PdfSharp.Fonts.FontResolverInfo("Malgun Gothic");
    }

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
        public string TextContent { get => _textContent; set { _textContent = value; OnPropertyChanged(nameof(TextContent)); } }
        
        public AnnotationType Type { get; set; } = AnnotationType.Other;
        public Color AnnotationColor { get; set; } = Colors.Transparent;
        
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
    public class RawAnnotationInfo { public AnnotationType Type; public XRect Rect; public string Content = ""; public double FontSize; public Color Color; }

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

        // [수정] 불필요한 보정 속성 제거 (Docnet 좌표 그대로 사용)
        
        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }
        private double _selX; public double SelectionX { get => _selX; set { _selX = value; OnPropertyChanged(nameof(SelectionX)); } }
        private double _selY; public double SelectionY { get => _selY; set { _selY = value; OnPropertyChanged(nameof(SelectionY)); } }
        private double _selW; public double SelectionWidth { get => _selW; set { _selW = value; OnPropertyChanged(nameof(SelectionWidth)); } }
        private double _selH; public double SelectionHeight { get => _selH; set { _selH = value; OnPropertyChanged(nameof(SelectionHeight)); } }
        
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private IDocLib _docLib;
        public ObservableCollection<PdfDocumentModel> Documents { get; set; } = new ObservableCollection<PdfDocumentModel>();
        
        private PdfDocumentModel? _selectedDocument;
        public PdfDocumentModel? SelectedDocument
        {
            get => _selectedDocument;
            set { _selectedDocument = value; OnPropertyChanged(nameof(SelectedDocument)); CheckToolbarVisibility(); }
        }

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
        private int _selectedPageIndex = -1;
        
        private bool _isDraggingAnnotation = false;
        private Point _annotationDragStartOffset;
        private bool _isUpdatingUiFromSelection = false;

        // 검색 관련 필드
        private List<PdfAnnotation> _searchResults = new List<PdfAnnotation>();
        private int _currentSearchIndex = -1;
        private string _lastSearchQuery = "";

        public MainWindow()
        {
            InitializeComponent();
            try { if (PdfSharp.Fonts.GlobalFontSettings.FontResolver == null) PdfSharp.Fonts.GlobalFontSettings.FontResolver = new WindowsFontResolver(); } catch { }
            _docLib = DocLib.Instance;
            DataContext = this;
            try { _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR")) ?? OcrEngine.TryCreateFromUserProfileLanguages(); } catch { }

            CbFont.ItemsSource = new string[] { "Malgun Gothic", "Gulim", "Dotum", "Batang" };
            CbFont.SelectedIndex = 0;
            CbSize.ItemsSource = new double[] { 10, 12, 14, 16, 18, 24, 32, 48 };
            CbSize.SelectedIndex = 2;
        }

        private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PdfDocumentModel doc)
            {
                doc.DocReader?.Dispose();
                Documents.Remove(doc);
                if (Documents.Count == 0) 
                {
                    TxtStatus.Text = "파일을 열어주세요.";
                    SelectedDocument = null; 
                }
            }
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e) { var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" }; if (dlg.ShowDialog() == true) LoadPdf(dlg.FileName); }
        
        private void LoadPdf(string path)
        {
            try
            {
                var newDoc = new PdfDocumentModel { FilePath = path, FileName = Path.GetFileName(path) };
                var fileBytes = File.ReadAllBytes(path);

                var extractedRawData = new Dictionary<int, List<RawAnnotationInfo>>();
                var pdfPageSizes = new Dictionary<int, XSize>();
                
                using (var msInput = new MemoryStream(fileBytes))
                using (var doc = PdfReader.Open(msInput, PdfDocumentOpenMode.Modify))
                {
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        pdfPageSizes[i] = new XSize(page.Width.Point, page.Height.Point);
                        
                        // [수정] 복잡한 CropBox/MediaBox 오프셋 계산 제거 (Docnet 호환성 위해)

                        extractedRawData[i] = new List<RawAnnotationInfo>();
                        
                        if (page.Annotations != null)
                        {
                            var annotsToRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for (int k = 0; k < page.Annotations.Count; k++)
                            {
                                var ann = page.Annotations[k];
                                var subtype = ann.Elements.GetString("/Subtype");

                                if (subtype == "/FreeText") {
                                    var rect = ann.Rectangle.ToXRect();
                                    extractedRawData[i].Add(new RawAnnotationInfo { Type = AnnotationType.FreeText, Rect = rect, Content = ann.Contents, FontSize = 14, Color = Colors.Red });
                                    annotsToRemove.Add(ann); 
                                }
                                else if (subtype == "/Highlight" || subtype == "/Underline") {
                                    var rect = ann.Rectangle.ToXRect();
                                    Color cColor = Colors.Yellow; 
                                    var cArray = ann.Elements.GetArray("/C");
                                    if (cArray != null && cArray.Elements.Count >= 3) {
                                        double r = (cArray.Elements[0] as PdfReal)?.Value ?? 1.0;
                                        double g = (cArray.Elements[1] as PdfReal)?.Value ?? 1.0;
                                        double b = (cArray.Elements[2] as PdfReal)?.Value ?? 0.0;
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
                    newDoc.DocReader = _docLib.GetDocReader(cleanStream.ToArray(), new PageDimensions(2.0)); // 2.0 배율 렌더링
                }

                if (newDoc.DocReader != null)
                {
                    int pc = newDoc.DocReader.GetPageCount();
                    for (int i = 0; i < pc; i++)
                    {
                        using (var r = newDoc.DocReader.GetPageReader(i))
                        {
                            double viewW = r.GetPageWidth(); double viewH = r.GetPageHeight();
                            
                            var pvm = new PdfPageViewModel { 
                                PageIndex = i, Width = viewW, Height = viewH,
                                PdfPageWidthPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Width : viewW,
                                PdfPageHeightPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Height : viewH
                            };

                            double scaleX = viewW / pvm.PdfPageWidthPoint; double scaleY = viewH / pvm.PdfPageHeightPoint;

                            if (extractedRawData.ContainsKey(i)) {
                                foreach (var raw in extractedRawData[i]) {
                                    // 기존 주석 처리 (PDFSharp 좌표계는 Bottom-Up이므로 변환 유지)
                                    // 주의: Docnet 검색 좌표와 PDFSharp 주석 좌표계는 다름
                                    double pdfTopY = raw.Rect.Y + raw.Rect.Height; 
                                    double viewY = (pvm.PdfPageHeightPoint - pdfTopY) * scaleY;
                                    var ann = new PdfAnnotation {
                                        Type = raw.Type, X = raw.Rect.X * scaleX, Y = viewY, Width = raw.Rect.Width * scaleX, Height = raw.Rect.Height * scaleY,
                                        TextContent = raw.Content, FontSize = raw.FontSize, FontFamily = "Malgun Gothic", Foreground = new SolidColorBrush(raw.Color), AnnotationColor = raw.Color
                                    };
                                    if (raw.Type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, raw.Color.R, raw.Color.G, raw.Color.B));
                                    else if (raw.Type == AnnotationType.Underline) { ann.Background = new SolidColorBrush(raw.Color); ann.Y = viewY + (raw.Rect.Height * scaleY) - 2; ann.Height = 2; }
                                    pvm.Annotations.Add(ann);
                                }
                            }
                            newDoc.Pages.Add(pvm);
                        }
                    }
                }

                Documents.Add(newDoc);
                SelectedDocument = newDoc;
                Task.Run(() => RenderPagesAsync(newDoc));
            }
            catch (Exception ex) { MessageBox.Show($"열기 실패: {ex.Message}"); }
        }

        private void RenderPagesAsync(PdfDocumentModel doc)
        {
            Parallel.For(0, doc.Pages.Count, new ParallelOptions { MaxDegreeOfParallelism = 4 }, i =>
            {
                if (doc.DocReader == null) return;
                using (var r = doc.DocReader.GetPageReader(i)) {
                    var bytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight();
                    Application.Current.Dispatcher.Invoke(() => { if (i < doc.Pages.Count) doc.Pages[i].ImageSource = RawBytesToBitmapImage(bytes, w, h); });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        private void PdfListView_Loaded(object sender, RoutedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView != null && listView.DataContext is PdfDocumentModel doc) {
                var scrollViewer = GetVisualChild<ScrollViewer>(listView);
                if (scrollViewer != null) {
                    scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
                    scrollViewer.ScrollToVerticalOffset(doc.SavedVerticalOffset);
                    scrollViewer.ScrollToHorizontalOffset(doc.SavedHorizontalOffset);
                    scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
                }
            }
        }

        private void PdfListView_Unloaded(object sender, RoutedEventArgs e)
        {
            var listView = sender as ListView;
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer != null)
            {
                scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
            }
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer sv && sv.DataContext is PdfDocumentModel doc)
            {
                doc.SavedVerticalOffset = sv.VerticalOffset;
                doc.SavedHorizontalOffset = sv.HorizontalOffset;
            }
        }
        
        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual
        {
            if (parent == null) return null;
            T? child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++) {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null) child = GetVisualChild<T>(v);
                if (child != null) break;
            }
            return child;
        }

        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) {
                if (SelectedDocument != null) {
                    if (e.Delta > 0) SelectedDocument.Zoom += 0.1; else SelectedDocument.Zoom -= 0.1;
                }
                e.Handled = true;
            }
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom += 0.1; }
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom -= 0.1; }

        private void BtnFitWidth_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0) return;
            double viewWidth = MainTabControl.ActualWidth - 60; 
            if (viewWidth > 0) {
                double pageWidth = SelectedDocument.Pages[0].Width; 
                if (pageWidth > 0) SelectedDocument.Zoom = viewWidth / pageWidth;
            }
        }

        private void BtnFitHeight_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0) return;
            double viewHeight = MainTabControl.ActualHeight - 60;
            if (viewHeight > 0) {
                double pageHeight = SelectedDocument.Pages[0].Height;
                if (pageHeight > 0) SelectedDocument.Zoom = viewHeight / pageHeight;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SavePdf(SelectedDocument.FilePath); }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) { if (SelectedDocument == null) return; var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_copy" }; if (dlg.ShowDialog() == true) SavePdf(dlg.FileName); }

        private void SavePdf(string savePath)
        {
            if (SelectedDocument == null) return;
            try
            {
                var originalBytes = File.ReadAllBytes(SelectedDocument.FilePath);
                using (var ms = new MemoryStream(originalBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    if (doc.Version < 14) doc.Version = 14;
                    for (int i = 0; i < doc.PageCount && i < SelectedDocument.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i]; var pageVM = SelectedDocument.Pages[i];
                        double scaleX = pdfPage.Width.Point / pageVM.Width; double scaleY = pdfPage.Height.Point / pageVM.Height;
                        
                        if (pdfPage.Annotations != null) {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            for(int k=0;k<pdfPage.Annotations.Count;k++) {
                                string st = pdfPage.Annotations[k].Elements.GetString("/Subtype");
                                if (st == "/FreeText" || st == "/Highlight" || st == "/Underline") toRemove.Add(pdfPage.Annotations[k]);
                            }
                            foreach(var a in toRemove) pdfPage.Annotations.Remove(a);
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

                            if (ann.Type == AnnotationType.FreeText) {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                var color = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else {
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
            if (SelectedDocument == null) return;

            var canvas = sender as Canvas; if (canvas == null) return;
            _activePageIndex = (int)canvas.Tag; _dragStartPoint = e.GetPosition(canvas);
            var pageVM = SelectedDocument.Pages[_activePageIndex];

            if (_currentTool == "TEXT") {
                var newAnnot = new PdfAnnotation {
                    Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50,
                    FontSize = _defaultFontSize, FontFamily = _defaultFontFamily,
                    Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold,
                    TextContent = "", IsSelected = true
                };
                pageVM.Annotations.Add(newAnnot);
                _selectedAnnotation = newAnnot;
                _currentTool = "CURSOR"; RbCursor.IsChecked = true;
                UpdateToolbarFromAnnotation(_selectedAnnotation);
                CheckToolbarVisibility(); e.Handled = true;
                return;
            }

            if (_currentTool == "CURSOR") {
                foreach (var p in SelectedDocument.Pages) { p.IsSelecting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                SelectionPopup.IsOpen = false;
                pageVM.IsSelecting = true;
                pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y;
                pageVM.SelectionWidth = 0; pageVM.SelectionHeight = 0;
                canvas.CaptureMouse(); e.Handled = true;
            }
            CheckToolbarVisibility();
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1 || SelectedDocument == null) return; 
            var canvas = sender as Canvas; if (canvas == null) return;
            var pageVM = SelectedDocument.Pages[_activePageIndex];

            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null && _selectedAnnotation.Type == AnnotationType.FreeText) {
                var currentPoint = e.GetPosition(canvas);
                _selectedAnnotation.X = currentPoint.X - _annotationDragStartOffset.X;
                _selectedAnnotation.Y = currentPoint.Y - _annotationDragStartOffset.Y;
                e.Handled = true;
                return;
            }

            if (_currentTool == "CURSOR" && pageVM.IsSelecting) {
                var pt = e.GetPosition(canvas);
                double x = Math.Min(_dragStartPoint.X, pt.X); double y = Math.Min(_dragStartPoint.Y, pt.Y);
                double w = Math.Abs(pt.X - _dragStartPoint.X); double h = Math.Abs(pt.Y - _dragStartPoint.Y);
                pageVM.SelectionX = x; pageVM.SelectionY = y; pageVM.SelectionWidth = w; pageVM.SelectionHeight = h;
            }
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var canvas = sender as Canvas; if (canvas == null || _activePageIndex == -1 || SelectedDocument == null) return;
            canvas.ReleaseMouseCapture(); 
            var p = SelectedDocument.Pages[_activePageIndex];
            
            if (p.IsSelecting && _currentTool == "CURSOR") {
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5) {
                    var rect = new Rect(p.SelectionX, p.SelectionY, p.SelectionWidth, p.SelectionHeight);
                    CheckTextInSelection(_activePageIndex, rect);
                    SelectionPopup.PlacementTarget = canvas; SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(canvas).X, e.GetPosition(canvas).Y + 10, 0, 0); SelectionPopup.IsOpen = true;
                    _selectedPageIndex = _activePageIndex;
                    TxtStatus.Text = string.IsNullOrEmpty(_selectedTextBuffer) ? "영역 선택됨" : "텍스트 선택됨";
                } else { 
                    p.IsSelecting = false; TxtStatus.Text = "준비"; 
                    if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; }
                }
            }
            _activePageIndex = -1; _isDraggingAnnotation = false; e.Handled = true;
            CheckToolbarVisibility(); 
        }

        private void Annotation_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool != "CURSOR") return;
            var element = sender as FrameworkElement;
            if (element?.DataContext is PdfAnnotation ann) {
                if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true;
                if (ann.Type == AnnotationType.FreeText) { _isDraggingAnnotation = true; _annotationDragStartOffset = e.GetPosition(element); }
                else { _isDraggingAnnotation = false; }
                UpdateToolbarFromAnnotation(ann); CheckToolbarVisibility(); e.Handled = true; 
            }
        }
        
        private void CheckToolbarVisibility() { bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null); if (TextStyleToolbar != null) TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed; }
        private void AnnotationTextBox_Loaded(object sender, RoutedEventArgs e) { if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann && ann.IsSelected) tb.Focus(); }
        private void AnnotationTextBox_GotFocus(object sender, RoutedEventArgs e) { if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann) { _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true; UpdateToolbarFromAnnotation(ann); CheckToolbarVisibility(); } }
        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e) { var thumb = sender as Thumb; if (thumb?.DataContext is PdfAnnotation ann) { ann.Width = Math.Max(50, ann.Width + e.HorizontalChange); ann.Height = Math.Max(30, ann.Height + e.VerticalChange); } }
        
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
            if (CbFont.SelectedItem != null) _defaultFontFamily = CbFont.SelectedItem.ToString() ?? "Malgun Gothic";
            if (CbSize.SelectedItem != null) _defaultFontSize = (double)CbSize.SelectedItem;
            _defaultIsBold = BtnBold.IsChecked == true;
            if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null) {
                string cn = item.Tag.ToString() ?? "Black";
                if (cn == "Black") _defaultFontColor = Colors.Black; else if (cn == "Red") _defaultFontColor = Colors.Red; else if (cn == "Blue") _defaultFontColor = Colors.Blue; else if (cn == "Green") _defaultFontColor = Colors.Green; else if (cn == "Orange") _defaultFontColor = Colors.Orange;
            }
            if (_selectedAnnotation != null) {
                _selectedAnnotation.FontFamily = _defaultFontFamily; _selectedAnnotation.FontSize = _defaultFontSize;
                _selectedAnnotation.IsBold = _defaultIsBold; _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor);
            }
        }

        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e) {
            if (_selectedAnnotation != null && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) { if (p.Annotations.Contains(_selectedAnnotation)) { p.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; CheckToolbarVisibility(); break; } } }
            else if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) if (p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; } }
        }

        private void Tool_Click(object sender, RoutedEventArgs e) { if (RbCursor.IsChecked == true) _currentTool = "CURSOR"; else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT"; else if (RbText.IsChecked == true) _currentTool = "TEXT"; CheckToolbarVisibility(); }

        private void Window_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; } else if (e.Key == Key.Delete) BtnDeleteAnnotation_Click(this, new RoutedEventArgs()); else if (e.Key == Key.Escape && _selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); } }

        private void BtnOCR_Click(object sender, RoutedEventArgs e) { 
            if (_ocrEngine == null || SelectedDocument == null) return; 
            BtnOCR.IsEnabled = false; PbStatus.Visibility = Visibility.Visible; PbStatus.Maximum = SelectedDocument.Pages.Count; PbStatus.Value = 0; TxtStatus.Text = "OCR 분석 중..."; 
            var targetDoc = SelectedDocument; 
            Task.Run(async () => { 
                try { 
                    for (int i = 0; i < targetDoc.Pages.Count; i++) { 
                        var pageVM = targetDoc.Pages[i]; if (targetDoc.DocReader == null) continue; 
                        using (var r = targetDoc.DocReader.GetPageReader(i)) { 
                            var rawBytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight(); 
                            using (var stream = new MemoryStream(rawBytes)) { 
                                var sb = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied);
                                var ibuffer = CryptographicBuffer.CreateFromByteArray(rawBytes);
                                sb.CopyFromBuffer(ibuffer); 
                                var res = await _ocrEngine.RecognizeAsync(sb); 
                                var list = new List<OcrWordInfo>(); foreach (var l in res.Lines) foreach (var wd in l.Words) list.Add(new OcrWordInfo { Text = wd.Text, BoundingBox = new Rect(wd.BoundingRect.X, wd.BoundingRect.Y, wd.BoundingRect.Width, wd.BoundingRect.Height) }); 
                                pageVM.OcrWords = list; 
                            } 
                        } 
                        Application.Current.Dispatcher.Invoke(() => PbStatus.Value = i + 1); 
                    } 
                    Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "OCR 완료"); 
                } catch (Exception ex) { Application.Current.Dispatcher.Invoke(() => MessageBox.Show($"OCR 오류: {ex.Message}")); } 
                finally { Application.Current.Dispatcher.Invoke(() => { BtnOCR.IsEnabled = true; PbStatus.Visibility = Visibility.Collapsed; }); } 
            }); 
        }

        private void CheckTextInSelection(int pageIndex, Rect uiRect) { 
            _selectedTextBuffer = ""; if (SelectedDocument?.DocReader == null) return; 
            var sb = new StringBuilder(); 
            using (var reader = SelectedDocument.DocReader.GetPageReader(pageIndex)) { 
                var chars = reader.GetCharacters().ToList(); 
                foreach (var c in chars) { var r = new Rect(Math.Min(c.Box.Left, c.Box.Right), Math.Min(c.Box.Top, c.Box.Bottom), Math.Abs(c.Box.Right - c.Box.Left), Math.Abs(c.Box.Bottom - c.Box.Top)); if (uiRect.IntersectsWith(r)) sb.Append(c.Char); } 
            } 
            var pageVM = SelectedDocument.Pages[pageIndex];
            if (pageVM.OcrWords != null) { foreach (var word in pageVM.OcrWords) if (uiRect.IntersectsWith(word.BoundingBox)) sb.Append(word.Text + " "); }
            _selectedTextBuffer = sb.ToString(); 
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            PerformSearch(TxtSearch.Text);
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
                    NavigateSearchResult(false); 
                else
                {
                    string query = TxtSearch.Text;
                    if (query == _lastSearchQuery && _searchResults.Count > 0)
                        NavigateSearchResult(true); 
                    else
                        PerformSearch(query); 
                }
                e.Handled = true;
            }
        }

        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e)
        {
            NavigateSearchResult(false);
        }

        private void BtnNextSearch_Click(object sender, RoutedEventArgs e)
        {
            NavigateSearchResult(true);
        }

        // [핵심 수정] 좌표잘찾는 로직 적용 (단순화)
        private void PerformSearch(string query)
        {
            if (string.IsNullOrWhiteSpace(query) || SelectedDocument == null) return;

            _lastSearchQuery = query;
            _searchResults.Clear();
            _currentSearchIndex = -1;

            foreach (var p in SelectedDocument.Pages)
            {
                var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                foreach (var r in toRemove) p.Annotations.Remove(r);
            }

            if (SelectedDocument.DocReader != null)
            {
                for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                {
                    var pageVM = SelectedDocument.Pages[i];
                    double scaleX = pageVM.Width / pageVM.PdfPageWidthPoint;
                    double scaleY = pageVM.Height / pageVM.PdfPageHeightPoint;
                    
                    // 1. OCR 검색
                    if (pageVM.OcrWords != null)
                    {
                        foreach (var word in pageVM.OcrWords)
                        {
                            if (word.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                            {
                                // OCR 좌표는 이미지 상의 좌표이므로 그대로 사용 (Scale만 적용 필요 시 적용하나 보통 이미지 기준)
                                // OcrEngine의 BoundingBox는 보통 이미지 픽셀 단위
                                var ann = new PdfAnnotation
                                {
                                    X = word.BoundingBox.X, // OCR 엔진이 이미지 좌표를 주므로 Scale 불필요 (이미 PageVM 크기에 맞음)
                                    Y = word.BoundingBox.Y,
                                    Width = word.BoundingBox.Width,
                                    Height = word.BoundingBox.Height,
                                    Background = new SolidColorBrush(Color.FromArgb(120, 255, 255, 0)),
                                    Type = AnnotationType.SearchHighlight
                                };
                                pageVM.Annotations.Add(ann);
                                _searchResults.Add(ann);
                            }
                        }
                    }

                    // 2. Docnet 텍스트 검색
                    using (var pageReader = SelectedDocument.DocReader.GetPageReader(i))
                    {
                        string pageText = pageReader.GetText();
                        var chars = pageReader.GetCharacters().ToList();

                        int index = 0;
                        while ((index = pageText.IndexOf(query, index, StringComparison.OrdinalIgnoreCase)) != -1)
                        {
                            double minX = double.MaxValue, minY = double.MaxValue;
                            double maxX = double.MinValue, maxY = double.MinValue;
                            bool found = false;

                            for (int c = 0; c < query.Length; c++)
                            {
                                if (index + c < chars.Count)
                                {
                                    var box = chars[index + c].Box;
                                    
                                    // [좌표잘찾는 파일의 핵심 로직] 
                                    // Docnet은 Top-Left 기준 좌표를 반환하므로 Y반전 없이 그대로 사용
                                    // Left/Right, Top/Bottom 중 작은/큰 값을 안전하게 추출
                                    double cLeft = Math.Min(box.Left, box.Right);
                                    double cRight = Math.Max(box.Left, box.Right);
                                    double cTop = Math.Min(box.Top, box.Bottom);
                                    double cBottom = Math.Max(box.Top, box.Bottom);

                                    minX = Math.Min(minX, cLeft);
                                    minY = Math.Min(minY, cTop);
                                    maxX = Math.Max(maxX, cRight);
                                    maxY = Math.Max(maxY, cBottom);
                                    
                                    found = true;
                                }
                            }

                            if (found)
                            {
                                // Scale만 적용 (MinsPDFViewer는 2.0배 렌더링 등 확대/축소 지원하므로)
                                double finalX = minX * scaleX;
                                double finalY = minY * scaleY;
                                double finalW = (maxX - minX) * scaleX;
                                double finalH = (maxY - minY) * scaleY;

                                var ann = new PdfAnnotation
                                {
                                    X = finalX,
                                    Y = finalY,
                                    Width = finalW,
                                    Height = finalH,
                                    Background = new SolidColorBrush(Color.FromArgb(120, 0, 255, 255)), // 진한 하늘색
                                    Type = AnnotationType.SearchHighlight
                                };
                                pageVM.Annotations.Add(ann);
                                _searchResults.Add(ann);
                            }
                            index += query.Length;
                        }
                    }
                }
            }

            TxtStatus.Text = $"검색 결과: {_searchResults.Count}건";
            if (_searchResults.Count > 0) NavigateSearchResult(true);
            else MessageBox.Show("검색 결과가 없습니다.");
        }

        private void NavigateSearchResult(bool next)
        {
            if (_searchResults.Count == 0) return;

            if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
            {
                _searchResults[_currentSearchIndex].Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255));
            }

            if (next)
            {
                _currentSearchIndex++;
                if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0;
            }
            else
            {
                _currentSearchIndex--;
                if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1;
            }

            var currentAnnot = _searchResults[_currentSearchIndex];
            currentAnnot.Background = new SolidColorBrush(Color.FromArgb(120, 255, 0, 255)); 

            if (SelectedDocument != null)
            {
                var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.Annotations.Contains(currentAnnot));
                if (targetPage != null)
                {
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null)
                    {
                        listView.ScrollIntoView(targetPage);
                    }
                    TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}";
                }
            }
        }
        
        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e) { Clipboard.SetText(_selectedTextBuffer); SelectionPopup.IsOpen = false; }
        private void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e) { SelectionPopup.IsOpen = false; }
        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) { AddAnnotation(Colors.Lime, AnnotationType.Highlight); }
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) { AddAnnotation(Colors.Orange, AnnotationType.Highlight); }
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) { AddAnnotation(Colors.Black, AnnotationType.Underline); }
        private void AddAnnotation(Color color, AnnotationType type) { 
            if (_selectedPageIndex == -1 || SelectedDocument == null) return; 
            var p = SelectedDocument.Pages[_selectedPageIndex]; 
            var ann = new PdfAnnotation { X = p.SelectionX, Y = p.SelectionY, Width = p.SelectionWidth, Height = p.SelectionHeight, Type = type, AnnotationColor = color }; 
            if (type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B)); else { ann.Background = new SolidColorBrush(color); ann.Height = 2; ann.Y = p.SelectionY + p.SelectionHeight - 2; } 
            p.Annotations.Add(ann); SelectionPopup.IsOpen = false; p.IsSelecting = false;
        }

        private BitmapImage RawBytesToBitmapImage(byte[] b, int w, int h) { var bm = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null); bm.WritePixels(new Int32Rect(0, 0, w, h), b, w * 4, 0); if (bm.CanFreeze) bm.Freeze(); return ConvertWriteableBitmapToBitmapImage(bm); }
        private BitmapImage ConvertWriteableBitmapToBitmapImage(WriteableBitmap wbm) { using (var ms = new MemoryStream()) { var enc = new PngBitmapEncoder(); enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(wbm)); enc.Save(ms); var img = new BitmapImage(); img.BeginInit(); img.CacheOption = BitmapCacheOption.OnLoad; img.StreamSource = ms; img.EndInit(); if (img.CanFreeze) img.Freeze(); return img; } }
        
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}