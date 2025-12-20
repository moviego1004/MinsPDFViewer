using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
// PdfSharp 제거 (Docnet만 사용하여 좌표계 통일)
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
    public class PdfAnnotation : INotifyPropertyChanged
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        private Brush _background = Brushes.Transparent;
        public Brush Background 
        {
            get => _background;
            set { _background = value; OnPropertyChanged(nameof(Background)); }
        }

        private Brush _borderBrush = Brushes.Transparent;
        public Brush BorderBrush
        {
            get => _borderBrush;
            set { _borderBrush = value; OnPropertyChanged(nameof(BorderBrush)); }
        }

        private Thickness _borderThickness = new Thickness(0);
        public Thickness BorderThickness
        {
            get => _borderThickness;
            set { _borderThickness = value; OnPropertyChanged(nameof(BorderThickness)); }
        }

        public int SearchResultId { get; set; } = -1;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class PdfPageViewModel : INotifyPropertyChanged
    {
        public int PageIndex { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        // Docnet 1.0 배율 기준 원본 크기
        public double BaseW { get; set; }
        public double BaseH { get; set; }
        
        public double ScaleX { get; set; } = 1.0;
        public double ScaleY { get; set; } = 1.0;

        private ImageSource? _imageSource;
        public ImageSource? ImageSource 
        {
            get => _imageSource;
            set { _imageSource = value; OnPropertyChanged(nameof(ImageSource)); }
        }

        public ObservableCollection<PdfAnnotation> Annotations { get; set; } = new ObservableCollection<PdfAnnotation>();

        private bool _isSelecting;
        public bool IsSelecting
        {
            get => _isSelecting;
            set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); }
        }

        private Rect _selectionRect;
        public Rect SelectionRect
        {
            get => _selectionRect;
            set { _selectionRect = value; OnPropertyChanged(nameof(SelectionRect)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class SearchResultItem
    {
        public int PageIndex { get; set; }
        public int AnnotationIndex { get; set; }
    }

    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        // private double _renderScale = 1.5; 
        
        private double _renderScale = 1; 

        private Point _dragStartPoint;
        private int _activePageIndex = -1;
        private string _selectedTextBuffer = "";
        private List<Docnet.Core.Models.Character> _selectedCharacters = new List<Docnet.Core.Models.Character>();
        private int _selectedPageIndex = -1;

        private List<SearchResultItem> _searchResults = new List<SearchResultItem>();
        private int _currentSearchIndex = -1;
        private double _currentZoom = 1.0;

        public MainWindow()
        {
            InitializeComponent();
            _docLib = DocLib.Instance;
            PdfListView.ItemsSource = Pages;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                TxtSearch.Focus();
                TxtSearch.SelectAll();
                e.Handled = true;
            }
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
                _docReader?.Dispose();
                Pages.Clear();
                _searchResults.Clear();
                _currentSearchIndex = -1;

                // [Step 1] Docnet 1.0 배율로 '기준 크기(Base Size)' 측정
                // PdfSharp 대신 Docnet을 사용하여 좌표계 불일치 문제 해결
                var baseSizes = new List<(double w, double h)>();
                
                using (var tempReader = _docLib.GetDocReader(path, new PageDimensions(1.0)))
                {
                    int count = tempReader.GetPageCount();
                    for(int i=0; i<count; i++)
                    {
                        using(var r = tempReader.GetPageReader(i))
                        {
                            baseSizes.Add((r.GetPageWidth(), r.GetPageHeight()));
                        }
                    }
                }

                // [Step 2] 실제 렌더링 (1.5배)
                _docReader = _docLib.GetDocReader(path, new PageDimensions(_renderScale));
                int pageCount = _docReader.GetPageCount();
                TxtStatus.Text = $"총 {pageCount} 페이지 로딩 중...";

                for (int i = 0; i < pageCount; i++)
                {
                    using (var r = _docReader.GetPageReader(i))
                    {
                        double renderedW = r.GetPageWidth();
                        double renderedH = r.GetPageHeight();
                        
                        var (baseW, baseH) = (i < baseSizes.Count) ? baseSizes[i] : (renderedW/_renderScale, renderedH/_renderScale);

                        // 비율 계산 (렌더링 크기 / 기준 크기)
                        double scaleX = (baseW > 0) ? (renderedW / baseW) : _renderScale;
                        double scaleY = (baseH > 0) ? (renderedH / baseH) : _renderScale;

                        Pages.Add(new PdfPageViewModel 
                        { 
                            PageIndex = i, 
                            Width = renderedW, 
                            Height = renderedH,
                            BaseW = baseW,
                            BaseH = baseH,
                            ScaleX = scaleX,
                            ScaleY = scaleY
                        });
                    }
                }

                Task.Run(() => RenderAllPagesAsync(pageCount));
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.ApplicationIdle, new Action(() => FitWidth()));
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
                        var img = RawBytesToBitmapImage(bytes, w, h);
                        if (i < Pages.Count) Pages[i].ImageSource = img;
                    });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        // =========================================================
        // [검색 기능] 좌표계 자동 감지 및 단순화
        // =========================================================
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) 
        { 
            if (e.Key == Key.Enter) 
            {
                if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift) NavigateSearch(-1);
                else 
                {
                    if (_searchResults.Count > 0) NavigateSearch(1);
                    else PerformSearch(TxtSearch.Text);
                }
            } 
        }

        private void PerformSearch(string query)
        {
            if (_docReader == null || string.IsNullOrWhiteSpace(query)) return;
            
            foreach (var p in Pages) 
            {
                var toRemove = p.Annotations.Where(a => a.SearchResultId != -1).ToList();
                foreach (var item in toRemove) p.Annotations.Remove(item);
            }
            _searchResults.Clear();
            _currentSearchIndex = -1;

            int pageCount = _docReader.GetPageCount();
            int globalSearchId = 0;

            for (int i = 0; i < pageCount; i++)
            {
                var pageVM = Pages[i];

                using (var r = _docReader.GetPageReader(i))
                {
                    string text = r.GetText();
                    var chars = r.GetCharacters().ToList();
                    int idx = 0;

                    while ((idx = text.IndexOf(query, idx, StringComparison.OrdinalIgnoreCase)) != -1)
                    {
                        double minX = double.MaxValue, maxX = double.MinValue;
                        double minY = double.MaxValue, maxY = double.MinValue;
                        bool found = false;

                        // [좌표계 감지]
                        // Docnet(PDFium)이 반환하는 좌표가 Top-Left(화면식)인지 Bottom-Left(PDF식)인지 확인
                        bool isScreenCoord = false; 
                        
                        // 첫 글자로 확인 (Top이 Bottom보다 작으면 화면 좌표계임)
                        for (int c = 0; c < query.Length; c++)
                        {
                            if (idx + c < chars.Count)
                            {
                                var b = chars[idx + c].Box;
                                if (b.Top != 0 || b.Bottom != 0) // 유효한 박스라면
                                {
                                    if (b.Top < b.Bottom) isScreenCoord = true;
                                    break; 
                                }
                            }
                        }

                        for (int c = 0; c < query.Length; c++)
                        {
                            if (idx + c < chars.Count)
                            {
                                var b = chars[idx + c].Box;
                                if (b.Left == 0 && b.Right == 0) continue;

                                minX = Math.Min(minX, b.Left);
                                maxX = Math.Max(maxX, b.Right);
                                
                                // Y축: 무조건 Min/Max로 범위 확보
                                double y1 = b.Top;
                                double y2 = b.Bottom;
                                minY = Math.Min(minY, Math.Min(y1, y2));
                                maxY = Math.Max(maxY, Math.Max(y1, y2));
                                
                                found = true;
                            }
                        }

                        if (found)
                        {
                            // [좌표 변환 로직 - 단순화]
                            // 1. 너비/높이 계산
                            double pdfW = Math.Abs(maxX - minX);
                            double pdfH = Math.Abs(maxY - minY);

                            // 2. 화면 좌표계로 변환
                            double finalX = minX * pageVM.ScaleX;
                            double finalY = 0;

                            if (isScreenCoord)
                            {
                                // 이미 화면 좌표계(Top-Left 0)라면 그대로 사용
                                // minY가 상단(Top)임
                                finalY = minY * pageVM.ScaleY;
                            }
                            else
                            {
                                // PDF 좌표계(Bottom-Left 0)라면 뒤집어야 함
                                // maxY가 상단(Top)임 (값이 큼)
                                // 화면 Y = (PageHeight - MaxY) * Scale
                                finalY = (pageVM.BaseH - maxY) * pageVM.ScaleY;
                            }

                            var annotation = new PdfAnnotation 
                            { 
                                X = finalX, 
                                Y = finalY, 
                                Width = pdfW * pageVM.ScaleX, 
                                Height = pdfH * pageVM.ScaleY,
                                Background = new SolidColorBrush(Color.FromArgb(60, 255, 255, 0)),
                                SearchResultId = globalSearchId
                            };
                            pageVM.Annotations.Add(annotation);
                            _searchResults.Add(new SearchResultItem { PageIndex = i, AnnotationIndex = globalSearchId });
                            globalSearchId++;
                        }
                        idx += query.Length;
                    }
                }
            }

            TxtStatus.Text = $"검색 결과: {_searchResults.Count}건";
            if (_searchResults.Count > 0)
            {
                _currentSearchIndex = -1;
                NavigateSearch(1);
            }
            else MessageBox.Show("검색 결과가 없습니다.");
        }

        private void NavigateSearch(int direction)
        {
            if (_searchResults.Count == 0) return;
            if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count) HighlightCurrentResult(_currentSearchIndex, false);

            _currentSearchIndex += direction;
            if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0;
            if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1;

            HighlightCurrentResult(_currentSearchIndex, true);
            var result = _searchResults[_currentSearchIndex];
            PdfListView.ScrollIntoView(Pages[result.PageIndex]);
            TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}";
        }

        private void HighlightCurrentResult(int index, bool isFocused)
        {
            var targetId = index;
            var resultPageIdx = _searchResults[index].PageIndex;
            var pageVM = Pages[resultPageIdx];

            foreach (var ann in pageVM.Annotations)
            {
                if (ann.SearchResultId == targetId)
                {
                    ann.Background = isFocused 
                        ? new SolidColorBrush(Color.FromArgb(120, 255, 165, 0)) 
                        : new SolidColorBrush(Color.FromArgb(60, 255, 255, 0));
                }
            }
        }

        // =========================================================
        // [드래그] - 안정화 후 구현
        // =========================================================
        private void Page_MouseDown(object sender, MouseButtonEventArgs e) { }
        private void Page_MouseMove(object sender, MouseEventArgs e) { }
        private void Page_MouseUp(object sender, MouseButtonEventArgs e) { }
        private void BtnCopy_Click(object sender, RoutedEventArgs e) { }
        private void BtnHighlightYellow_Click(object sender, RoutedEventArgs e) { }
        private void BtnHighlightOrange_Click(object sender, RoutedEventArgs e) { }
        private void BtnUnderline_Click(object sender, RoutedEventArgs e) { }
        private void Tool_Click(object sender, RoutedEventArgs e) { } 

        // --- 뷰어 제어 ---
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
            if (w <= 0) return; 
            if(Pages[0].Width > 0) UpdateZoom(w / Pages[0].Width);
        }

        private void FitHeight()
        {
            if(Pages.Count == 0) return;
            double h = PdfListView.ActualHeight - 60;
            if (h <= 0) return;
            if(Pages[0].Height > 0) UpdateZoom(h / Pages[0].Height);
        }

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