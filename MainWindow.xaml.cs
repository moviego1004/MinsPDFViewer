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
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Text;

namespace MinsPDFViewer
{
    // [ViewModel] 주석 타입 열거형
    public enum AnnotationType
    {
        Highlight,
        Underline,
        SearchHighlight, 
        Other
    }

    // [ViewModel] 주석
    public class PdfAnnotation : INotifyPropertyChanged
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        private Brush _background = Brushes.Transparent;
        public Brush Background { get => _background; set { _background = value; OnPropertyChanged(nameof(Background)); } }
        
        private Brush _borderBrush = Brushes.Transparent;
        public Brush BorderBrush { get => _borderBrush; set { _borderBrush = value; OnPropertyChanged(nameof(BorderBrush)); } }
        
        private Thickness _borderThickness = new Thickness(0);
        public Thickness BorderThickness { get => _borderThickness; set { _borderThickness = value; OnPropertyChanged(nameof(BorderThickness)); } }

        public string TextContent { get; set; } = "";
        public bool HasText => !string.IsNullOrEmpty(TextContent);
        
        public AnnotationType Type { get; set; } = AnnotationType.Other;
        public Color AnnotationColor { get; set; } = Colors.Transparent;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // [ViewModel] 페이지
    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        private ImageSource? _imageSource;
        public ImageSource? ImageSource { get => _imageSource; set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); } }

        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();

        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }

        private double _selX;
        public double SelectionX { get => _selX; set { _selX = value; OnPropertyChanged(nameof(SelectionX)); } }

        private double _selY;
        public double SelectionY { get => _selY; set { _selY = value; OnPropertyChanged(nameof(SelectionY)); } }

        private double _selW;
        public double SelectionWidth { get => _selW; set { _selW = value; OnPropertyChanged(nameof(SelectionWidth)); } }

        private double _selH;
        public double SelectionHeight { get => _selH; set { _selH = value; OnPropertyChanged(nameof(SelectionHeight)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // PDFsharp 6.x 호환용 커스텀 어노테이션
    public class CustomPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public CustomPdfAnnotation(PdfDocument document) : base(document) { }
    }

    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        private double _renderScale = 1.0; 
        private double _currentZoom = 1.0;
        private string _currentFilePath = "";

        private Point _dragStartPoint;
        private int _activePageIndex = -1; 
        private string _selectedTextBuffer = "";
        private List<Docnet.Core.Models.Character> _selectedChars = new List<Docnet.Core.Models.Character>();
        private int _selectedPageIndex = -1; 

        private string _currentTool = "CURSOR"; 

        private List<PdfAnnotation> _searchResults = new List<PdfAnnotation>();
        private int _currentSearchIndex = -1;
        private string _lastSearchQuery = "";

        private byte[] _originalFileBytes = Array.Empty<byte>();

        public MainWindow()
        {
            InitializeComponent();
            _docLib = DocLib.Instance;
            PdfListView.ItemsSource = Pages;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" };
            if (dlg.ShowDialog() == true) LoadPdf(dlg.FileName);
        }

        private void LoadPdf(string path)
        {
            try
            {
                _currentFilePath = path;
                _originalFileBytes = File.ReadAllBytes(path);

                Pages.Clear();

                using (var msInput = new MemoryStream(_originalFileBytes))
                using (var pdfDoc = PdfReader.Open(msInput, PdfDocumentOpenMode.Modify))
                {
                    int pageCount = pdfDoc.PageCount;
                    
                    for (int i = 0; i < pageCount; i++)
                    {
                        var pdfPage = pdfDoc.Pages[i];
                        double w = pdfPage.Width.Point;
                        double h = pdfPage.Height.Point;

                        var pageVM = new PdfPageViewModel { PageIndex = i, Width = w, Height = h };
                        Pages.Add(pageVM);

                        if (pdfPage.Annotations != null && pdfPage.Annotations.Count > 0)
                        {
                            var annotationsToRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();

                            foreach (var item in pdfPage.Annotations)
                            {
                                var annot = item as PdfSharp.Pdf.Annotations.PdfAnnotation;
                                if (annot == null) continue;

                                var subType = annot.Elements.GetString("/Subtype");

                                if (subType == "/Highlight" || subType == "/Underline")
                                {
                                    var rect = annot.Rectangle;
                                    double pW = rect.Width;
                                    double pH = rect.Height;
                                    double pX = rect.X1; 
                                    double pY = h - rect.Y2; 

                                    var newAnnot = new PdfAnnotation
                                    {
                                        X = pX, Y = pY, Width = pW, Height = pH,
                                        Type = (subType == "/Highlight") ? AnnotationType.Highlight : AnnotationType.Underline,
                                    };

                                    Color c = Colors.Yellow; 
                                    var colorItem = annot.Elements["/C"]; 
                                    if (colorItem is PdfSharp.Pdf.PdfArray arr && arr.Elements.Count >= 3)
                                    {
                                        byte r = (byte)(arr.Elements.GetReal(0) * 255);
                                        byte g = (byte)(arr.Elements.GetReal(1) * 255);
                                        byte b = (byte)(arr.Elements.GetReal(2) * 255);
                                        c = Color.FromRgb(r, g, b);
                                    }
                                    newAnnot.AnnotationColor = c;

                                    if (newAnnot.Type == AnnotationType.Highlight)
                                    {
                                        newAnnot.Background = new SolidColorBrush(Color.FromArgb(80, c.R, c.G, c.B));
                                    }
                                    else 
                                    {
                                        newAnnot.Background = new SolidColorBrush(c);
                                        newAnnot.Height = 2; 
                                        newAnnot.Y = pY + pH - 2; 
                                    }

                                    pageVM.Annotations.Add(newAnnot);
                                    annotationsToRemove.Add(annot);
                                }
                            }

                            foreach(var annot in annotationsToRemove)
                            {
                                pdfPage.Annotations.Remove(annot);
                            }
                        }
                    }

                    using (var msClean = new MemoryStream())
                    {
                        pdfDoc.Save(msClean);
                        msClean.Position = 0;
                        byte[] cleanBytes = msClean.ToArray(); 
                        _docReader?.Dispose();
                        _docReader = _docLib.GetDocReader(cleanBytes, new PageDimensions(_renderScale));
                    }
                }

                Task.Run(() => RenderAllPagesAsync(Pages.Count));
                FitWidth();
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
                    var bytes = r.GetImage();
                    var w = r.GetPageWidth();
                    var h = r.GetPageHeight();
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (i < Pages.Count)
                        {
                            var img = RawBytesToBitmapImage(bytes, w, h);
                            Pages[i].ImageSource = img;
                        }
                    });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        // =========================================================
        // [드래그 선택 로직]
        // =========================================================
        private void Page_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentTool != "CURSOR") return;

            var canvas = sender as Canvas;
            if (canvas == null) return;

            foreach (var p in Pages) 
            { 
                p.IsSelecting = false; 
                p.SelectionWidth = 0; 
                p.SelectionHeight = 0;
            }
            SelectionPopup.IsOpen = false;

            _activePageIndex = (int)canvas.Tag; 
            _dragStartPoint = e.GetPosition(canvas);
            
            var pageVM = Pages[_activePageIndex];
            pageVM.IsSelecting = true;
            pageVM.SelectionX = _dragStartPoint.X;
            pageVM.SelectionY = _dragStartPoint.Y;
            pageVM.SelectionWidth = 0;
            pageVM.SelectionHeight = 0;
            
            canvas.CaptureMouse();
            e.Handled = true; // 리스트뷰 간섭 방지
            
            TxtStatus.Text = $"드래그 시작: {_activePageIndex + 1} 페이지";
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1) return;
            var canvas = sender as Canvas;
            if (canvas == null) return;

            var currentPoint = e.GetPosition(canvas);
            var pageVM = Pages[_activePageIndex];

            double x = Math.Min(_dragStartPoint.X, currentPoint.X);
            double y = Math.Min(_dragStartPoint.Y, currentPoint.Y);
            double w = Math.Abs(currentPoint.X - _dragStartPoint.X);
            double h = Math.Abs(currentPoint.Y - _dragStartPoint.Y);

            pageVM.SelectionX = x;
            pageVM.SelectionY = y;
            pageVM.SelectionWidth = w;
            pageVM.SelectionHeight = h;
            
            TxtStatus.Text = $"선택 중.. W:{Math.Round(w)}, H:{Math.Round(h)}";
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var canvas = sender as Canvas;
            if (canvas == null || _activePageIndex == -1) return;

            canvas.ReleaseMouseCapture();
            var pageVM = Pages[_activePageIndex];
            
            if (pageVM.IsSelecting)
            {
                if (pageVM.SelectionWidth > 5 && pageVM.SelectionHeight > 5)
                {
                    var uiRect = new Rect(pageVM.SelectionX, pageVM.SelectionY, pageVM.SelectionWidth, pageVM.SelectionHeight);
                    CheckTextInSelection(_activePageIndex, uiRect);
                    
                    if (_selectedChars.Count > 0)
                    {
                        // 팝업 위치 설정
                        SelectionPopup.PlacementTarget = canvas;
                        SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(canvas).X, e.GetPosition(canvas).Y + 10, 0, 0);
                        SelectionPopup.IsOpen = true;
                        
                        _selectedPageIndex = _activePageIndex;
                        TxtStatus.Text = $"텍스트 선택됨: {_selectedChars.Count}자";
                    }
                    else
                    {
                        pageVM.IsSelecting = false; 
                        TxtStatus.Text = "선택된 텍스트 없음";
                    }
                }
                else
                {
                    pageVM.IsSelecting = false;
                    TxtStatus.Text = "준비";
                }
            }
            
            _activePageIndex = -1;
            e.Handled = true;
        }

        private void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            _selectedTextBuffer = "";
            _selectedChars.Clear();
            if (_docReader == null) return;

            using (var reader = _docReader.GetPageReader(pageIndex))
            {
                var chars = reader.GetCharacters().ToList();
                var sb = new StringBuilder();

                foreach (var c in chars)
                {
                    double cLeft = Math.Min(c.Box.Left, c.Box.Right);
                    double cRight = Math.Max(c.Box.Left, c.Box.Right);
                    double cTop = Math.Min(c.Box.Top, c.Box.Bottom);
                    double cBottom = Math.Max(c.Box.Top, c.Box.Bottom);
                    
                    var charRect = new Rect(cLeft, cTop, cRight - cLeft, cBottom - cTop);
                    
                    if (uiRect.IntersectsWith(charRect))
                    {
                        _selectedChars.Add(c);
                        sb.Append(c.Char);
                    }
                }
                _selectedTextBuffer = sb.ToString();
            }
        }

        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedTextBuffer))
            {
                Clipboard.SetText(_selectedTextBuffer);
                TxtStatus.Text = "클립보드에 복사됨";
            }
            CloseSelection();
        }

        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);

        private void AddAnnotation(Color color, AnnotationType type)
        {
            if (_selectedPageIndex == -1 || _selectedChars.Count == 0) return;
            var pageVM = Pages[_selectedPageIndex];

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (var c in _selectedChars)
            {
                double cLeft = Math.Min(c.Box.Left, c.Box.Right);
                double cRight = Math.Max(c.Box.Left, c.Box.Right);
                double cTop = Math.Min(c.Box.Top, c.Box.Bottom);
                double cBottom = Math.Max(c.Box.Top, c.Box.Bottom);

                minX = Math.Min(minX, cLeft);
                minY = Math.Min(minY, cTop);
                maxX = Math.Max(maxX, cRight);
                maxY = Math.Max(maxY, cBottom);
            }

            double w = maxX - minX;
            double h = maxY - minY;

            var annot = new PdfAnnotation
            {
                X = minX, Y = minY, Width = w, Height = h,
                Type = type,
                AnnotationColor = color
            };

            if (type == AnnotationType.Highlight)
            {
                annot.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B));
            }
            else if (type == AnnotationType.Underline)
            {
                annot.X = minX; 
                annot.Y = minY + h - 2; 
                annot.Width = w; 
                annot.Height = 2;
                annot.Background = new SolidColorBrush(color);
            }

            pageVM.Annotations.Add(annot);
            CloseSelection();
        }

        private void CloseSelection()
        {
            SelectionPopup.IsOpen = false;
            foreach (var p in Pages) p.IsSelecting = false;
        }

        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null && menuItem.CommandParameter is PdfAnnotation annot)
            {
                foreach(var page in Pages)
                {
                    if (page.Annotations.Contains(annot))
                    {
                        page.Annotations.Remove(annot);
                        break;
                    }
                }
            }
        }

        // =========================================================
        // [저장 로직 개선]
        // =========================================================
        
        // 1. 단순 저장 (덮어쓰기)
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;
            SavePdf(_currentFilePath);
        }

        // 2. 다른 이름으로 저장
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;

            var dlg = new SaveFileDialog 
            { 
                Filter = "PDF Files|*.pdf", 
                FileName = Path.GetFileNameWithoutExtension(_currentFilePath) + "_copy" 
            };
            
            if (dlg.ShowDialog() == true)
            {
                SavePdf(dlg.FileName);
                
                // 저장 후 현재 작업 파일 변경 (선택 사항, 여기서는 변경함)
                _currentFilePath = dlg.FileName;
            }
        }

        // 3. 공통 저장 메서드
        private void SavePdf(string savePath)
        {
            try
            {
                // 원본 바이트 기반으로 문서 열기
                using (var ms = new MemoryStream(_originalFileBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    for (int i = 0; i < doc.PageCount && i < Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i];
                        var pageVM = Pages[i];
                        double pageH = pdfPage.Height.Point;

                        // 기존 주석 정리 (중복 방지)
                        if (pdfPage.Annotations != null && pdfPage.Annotations.Count > 0)
                        {
                            var toRemove = new List<PdfSharp.Pdf.Annotations.PdfAnnotation>();
                            foreach (var item in pdfPage.Annotations)
                            {
                                var annot = item as PdfSharp.Pdf.Annotations.PdfAnnotation;
                                if (annot == null) continue;

                                var subType = annot.Elements.GetString("/Subtype");
                                if (subType == "/Highlight" || subType == "/Underline")
                                {
                                    toRemove.Add(annot);
                                }
                            }
                            foreach (var annot in toRemove) pdfPage.Annotations.Remove(annot);
                        }

                        // ViewModel 주석 추가
                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;

                            double pdfX = ann.X;
                            double pdfY = 0;
                            double pdfW = ann.Width;
                            double pdfH = ann.Height;

                            if (ann.Type == AnnotationType.Underline)
                            {
                                pdfH = 10; 
                                pdfY = pageH - (ann.Y - pdfH + 2); 
                            }
                            else
                            {
                                pdfY = pageH - (ann.Y + ann.Height);
                            }

                            var pdfAnnot = new CustomPdfAnnotation(doc);
                            
                            pdfAnnot.Rectangle = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));

                            string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                            pdfAnnot.Elements.SetName("/Subtype", subtype);

                            double r = ann.AnnotationColor.R / 255.0;
                            double g = ann.AnnotationColor.G / 255.0;
                            double b = ann.AnnotationColor.B / 255.0;
                            pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                            
                            pdfPage.Annotations.Add(pdfAnnot);
                        }
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }

        // =========================================================
        // [검색 및 기타 기능]
        // =========================================================
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text, true);
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(false);
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(true);

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                bool isShift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                if (TxtSearch.Text != _lastSearchQuery || _searchResults.Count == 0) PerformSearch(TxtSearch.Text, true);
                else NavigateSearchResult(!isShift);
            }
        }

        private void PerformSearch(string query, bool resetIndex)
        {
            if (_docReader == null || string.IsNullOrWhiteSpace(query)) return;
            
            _lastSearchQuery = query;
            _searchResults.Clear();

            foreach (var p in Pages) 
            {
                var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                foreach(var item in toRemove) p.Annotations.Remove(item);
            }

            int pageCount = _docReader.GetPageCount();

            for (int i = 0; i < pageCount; i++)
            {
                using (var r = _docReader.GetPageReader(i))
                {
                    string text = r.GetText();
                    var chars = r.GetCharacters().ToList();
                    int idx = 0;
                    while ((idx = text.IndexOf(query, idx, StringComparison.OrdinalIgnoreCase)) != -1)
                    {
                        double minX = double.MaxValue, minY = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue;
                        
                        for(int c=0; c<query.Length; c++)
                        {
                            if(idx+c < chars.Count)
                            {
                                var b = chars[idx+c].Box;
                                minX = Math.Min(minX, b.Left); minY = Math.Min(minY, b.Top);
                                maxX = Math.Max(maxX, b.Right); maxY = Math.Max(maxY, b.Bottom);
                            }
                        }
                        
                        var annotation = new PdfAnnotation 
                        { 
                            X=minX, Y=minY, Width=maxX-minX, Height=maxY-minY,
                            Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)),
                            Type = AnnotationType.SearchHighlight
                        };
                        
                        Pages[i].Annotations.Add(annotation);
                        _searchResults.Add(annotation);

                        idx += query.Length;
                    }
                }
            }
            
            TxtStatus.Text = $"검색: {_searchResults.Count}건";
            if (_searchResults.Count > 0)
            {
                _currentSearchIndex = -1;
                NavigateSearchResult(true);
            }
        }

        private void NavigateSearchResult(bool isNext)
        {
            if (_searchResults.Count == 0) return;

            if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
                _searchResults[_currentSearchIndex].Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255));

            if (isNext)
            {
                _currentSearchIndex++;
                if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0;
            }
            else
            {
                _currentSearchIndex--;
                if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1;
            }

            var currentAnnotation = _searchResults[_currentSearchIndex];
            currentAnnotation.Background = new SolidColorBrush(Color.FromArgb(120, 255, 0, 255));

            var targetPage = Pages.FirstOrDefault(p => p.Annotations.Contains(currentAnnotation));
            if (targetPage != null) PdfListView.ScrollIntoView(targetPage);

            TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}";
        }

        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (Keyboard.Modifiers == ModifierKeys.Control)
            {
                if (e.Delta > 0) UpdateZoom(_currentZoom + 0.1); else UpdateZoom(_currentZoom - 0.1);
                e.Handled = true;
            }
        }
        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom + 0.1);
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom - 0.1);
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) => FitWidth();
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) => FitHeight();

        private void UpdateZoom(double z)
        {
            _currentZoom = Math.Clamp(z, 0.2, 5.0);
            ViewScaleTransform.ScaleX = _currentZoom;
            ViewScaleTransform.ScaleY = _currentZoom;
            TxtZoom.Text = $"{Math.Round(_currentZoom * 100)}%";
        }
        private void FitWidth()
        {
            if(Pages.Count == 0) return;
            double w = PdfListView.ActualWidth - 60;
            if(Pages[0].Width > 0) UpdateZoom(w / Pages[0].Width);
        }
        private void FitHeight()
        {
            if(Pages.Count == 0) return;
            double h = PdfListView.ActualHeight - 60;
            if(Pages[0].Height > 0) UpdateZoom(h / Pages[0].Height);
        }

        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (RbCursor.IsChecked == true) _currentTool = "CURSOR";
            else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT";
            else if (RbText.IsChecked == true) _currentTool = "TEXT";
        }

        private void Window_KeyDown(object sender, KeyEventArgs e) { }

        private BitmapImage RawBytesToBitmapImage(byte[] b, int w, int h)
        {
            var bm = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
            bm.WritePixels(new Int32Rect(0, 0, w, h), b, w * 4, 0);
            if (bm.CanFreeze) bm.Freeze();
            return ConvertWriteableBitmapToBitmapImage(bm);
        }
        private BitmapImage ConvertWriteableBitmapToBitmapImage(WriteableBitmap wbm)
        {
            using (var ms = new MemoryStream())
            {
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(wbm));
                enc.Save(ms);
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.StreamSource = ms;
                img.EndInit();
                if (img.CanFreeze) img.Freeze();
                return img;
            }
        }
    }
}