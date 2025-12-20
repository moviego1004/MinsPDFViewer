using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
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
    // [ViewModel] 주석
    public class PdfAnnotation : INotifyPropertyChanged
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        
        // 스타일
        private Brush _background = Brushes.Transparent;
        public Brush Background { get => _background; set { _background = value; OnPropertyChanged(nameof(Background)); } }
        
        private Brush _borderBrush = Brushes.Transparent;
        public Brush BorderBrush { get => _borderBrush; set { _borderBrush = value; OnPropertyChanged(nameof(BorderBrush)); } }
        
        private Thickness _borderThickness = new Thickness(0);
        public Thickness BorderThickness { get => _borderThickness; set { _borderThickness = value; OnPropertyChanged(nameof(BorderThickness)); } }

        // 데이터
        public string TextContent { get; set; } = "";
        public bool HasText => !string.IsNullOrEmpty(TextContent);
        public int SearchResultId { get; set; } = -1;

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

        // 드래그 선택 박스
        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }

        private Rect _selectionRect;
        public Rect SelectionRect { get => _selectionRect; set { _selectionRect = value; OnPropertyChanged(nameof(SelectionRect)); } }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        // [중요] 1.0 배율 고정 (좌표 정확도)
        private double _renderScale = 1.0; 
        private double _currentZoom = 1.0;
        
        // [수정됨] 누락되었던 파일 경로 변수 추가
        private string _currentFilePath = "";

        // 드래그 상태
        private Point _dragStartPoint;
        private int _activePageIndex = -1; // 현재 드래그 중인 페이지
        private string _selectedTextBuffer = "";
        private List<Docnet.Core.Models.Character> _selectedChars = new List<Docnet.Core.Models.Character>();
        private int _selectedPageIndex = -1; // 팝업 액션을 위한 페이지 인덱스

        // 도구 상태
        private string _currentTool = "CURSOR"; 

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
                _docReader?.Dispose();
                _docReader = _docLib.GetDocReader(path, new PageDimensions(_renderScale));
                
                // [수정됨] 파일 경로 저장
                _currentFilePath = path; 

                Pages.Clear();
                int pageCount = _docReader.GetPageCount();
                TxtStatus.Text = $"총 {pageCount} 페이지 로딩 중...";

                for (int i = 0; i < pageCount; i++)
                {
                    using (var r = _docReader.GetPageReader(i))
                    {
                        Pages.Add(new PdfPageViewModel 
                        { 
                            PageIndex = i, 
                            Width = r.GetPageWidth(), 
                            Height = r.GetPageHeight() 
                        });
                    }
                }
                Task.Run(() => RenderAllPagesAsync(pageCount));
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
                        var img = RawBytesToBitmapImage(bytes, w, h);
                        if (i < Pages.Count) Pages[i].ImageSource = img;
                    });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        // =========================================================
        // [드래그 및 선택 로직] (연속 페이지 지원)
        // =========================================================
        private void Page_MouseDown(object sender, MouseButtonEventArgs e)
        {
            var canvas = sender as Canvas;
            if (canvas == null) return;

            // 다른 페이지의 선택 초기화
            foreach (var p in Pages) { p.IsSelecting = false; p.SelectionRect = new Rect(0,0,0,0); }
            SelectionPopup.IsOpen = false;

            _activePageIndex = (int)canvas.Tag; // 이벤트가 발생한 페이지 식별
            _dragStartPoint = e.GetPosition(canvas);
            
            if (_currentTool == "CURSOR")
            {
                var pageVM = Pages[_activePageIndex];
                pageVM.IsSelecting = true;
                pageVM.SelectionRect = new Rect(_dragStartPoint, new Size(0, 0));
                canvas.CaptureMouse();
            }
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1) return;
            var canvas = sender as Canvas;
            if (canvas == null || !canvas.IsMouseCaptured) return;

            var currentPoint = e.GetPosition(canvas);
            var pageVM = Pages[_activePageIndex];

            double x = Math.Min(_dragStartPoint.X, currentPoint.X);
            double y = Math.Min(_dragStartPoint.Y, currentPoint.Y);
            double w = Math.Abs(currentPoint.X - _dragStartPoint.X);
            double h = Math.Abs(currentPoint.Y - _dragStartPoint.Y);

            pageVM.SelectionRect = new Rect(x, y, w, h);
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e)
        {
            var canvas = sender as Canvas;
            if (canvas == null || _activePageIndex == -1) return;

            canvas.ReleaseMouseCapture();
            var pageVM = Pages[_activePageIndex];
            
            // 드래그 종료
            if (pageVM.IsSelecting)
            {
                if (pageVM.SelectionRect.Width > 5 && pageVM.SelectionRect.Height > 5)
                {
                    // 텍스트 선택 판정
                    CheckTextInSelection(_activePageIndex, pageVM.SelectionRect);
                    if (_selectedChars.Count > 0)
                    {
                        SelectionPopup.IsOpen = true;
                        _selectedPageIndex = _activePageIndex;
                    }
                    else
                    {
                        pageVM.IsSelecting = false; // 텍스트 없으면 박스 끄기
                    }
                }
                else
                {
                    pageVM.IsSelecting = false;
                }
            }
            _activePageIndex = -1;
        }

        private void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            _selectedTextBuffer = "";
            _selectedChars.Clear();
            if (_docReader == null) return;

            using (var reader = _docReader.GetPageReader(pageIndex))
            {
                // _renderScale = 1.0 이므로 좌표 변환 불필요
                var chars = reader.GetCharacters().ToList();
                var sb = new StringBuilder();

                foreach (var c in chars)
                {
                    var charRect = new Rect(c.Box.Left, c.Box.Top, c.Box.Right - c.Box.Left, c.Box.Bottom - c.Box.Top);
                    if (uiRect.IntersectsWith(charRect))
                    {
                        _selectedChars.Add(c);
                        sb.Append(c.Char);
                    }
                }
                _selectedTextBuffer = sb.ToString();
            }
        }

        // =========================================================
        // [팝업 액션]
        // =========================================================
        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedTextBuffer))
            {
                Clipboard.SetText(_selectedTextBuffer);
                TxtStatus.Text = "클립보드에 복사됨";
            }
            CloseSelection();
        }

        private void BtnPopupHighlightYellow_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Yellow, false);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, false);
        private void BtnPopupHighlightPurple_Click(object sender, RoutedEventArgs e) => AddAnnotation(Color.FromRgb(200, 130, 255), false);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, true);

        private void AddAnnotation(Color color, bool isUnderline)
        {
            if (_selectedPageIndex == -1 || _selectedChars.Count == 0) return;
            var pageVM = Pages[_selectedPageIndex];

            // 선택된 글자들의 Union Box 계산
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (var c in _selectedChars)
            {
                minX = Math.Min(minX, c.Box.Left);
                minY = Math.Min(minY, c.Box.Top);
                maxX = Math.Max(maxX, c.Box.Right);
                maxY = Math.Max(maxY, c.Box.Bottom);
            }

            double w = maxX - minX;
            double h = maxY - minY;

            if (isUnderline)
            {
                pageVM.Annotations.Add(new PdfAnnotation
                {
                    X = minX, Y = minY + h - 2, Width = w, Height = 2,
                    Background = new SolidColorBrush(color)
                });
            }
            else
            {
                pageVM.Annotations.Add(new PdfAnnotation
                {
                    X = minX, Y = minY, Width = w, Height = h,
                    Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B))
                });
            }
            CloseSelection();
        }

        private void CloseSelection()
        {
            SelectionPopup.IsOpen = false;
            foreach (var p in Pages) p.IsSelecting = false;
        }

        // =========================================================
        // [검색 및 줌 기능]
        // =========================================================
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) PerformSearch(TxtSearch.Text); }

        private void PerformSearch(string query)
        {
            if (_docReader == null || string.IsNullOrWhiteSpace(query)) return;
            
            // 기존 검색 결과 제거 (선택 사항)
            foreach (var p in Pages) 
            {
                var toRemove = p.Annotations.Where(a => a.SearchResultId != -1).ToList();
                foreach(var item in toRemove) p.Annotations.Remove(item);
            }

            int matchCount = 0;
            int firstPage = -1;
            int pageCount = _docReader.GetPageCount();
            int globalSearchId = 0;

            for (int i = 0; i < pageCount; i++)
            {
                using (var r = _docReader.GetPageReader(i))
                {
                    string text = r.GetText();
                    var chars = r.GetCharacters().ToList();
                    int idx = 0;
                    while ((idx = text.IndexOf(query, idx, StringComparison.OrdinalIgnoreCase)) != -1)
                    {
                        if(firstPage == -1) firstPage = i;
                        
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
                        
                        Pages[i].Annotations.Add(new PdfAnnotation 
                        { 
                            X=minX, Y=minY, Width=maxX-minX, Height=maxY-minY,
                            Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)), // Cyan 검색 하이라이트
                            SearchResultId = globalSearchId
                        });
                        
                        matchCount++;
                        globalSearchId++;
                        idx += query.Length;
                    }
                }
            }
            if(firstPage != -1) PdfListView.ScrollIntoView(Pages[firstPage]);
            TxtStatus.Text = $"검색: {matchCount}건";
        }

        // --- 뷰어 줌/맞춤 ---
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

        // 도구 선택 (단순 UI 토글)
        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (RbCursor.IsChecked == true) _currentTool = "CURSOR";
            else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT";
            else if (RbText.IsChecked == true) _currentTool = "TEXT";
        }

        // 저장 (PdfSharp 연동)
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;
            string savePath = _currentFilePath.Replace(".pdf", "_edited.pdf");
            try
            {
                using (var doc = PdfReader.Open(_currentFilePath, PdfDocumentOpenMode.Modify))
                {
                    foreach (var pageVM in Pages)
                    {
                        if (pageVM.Annotations.Count == 0) continue;
                        if (pageVM.PageIndex >= doc.PageCount) continue;
                        
                        var page = doc.Pages[pageVM.PageIndex];
                        using (var gfx = XGraphics.FromPdfPage(page))
                        {
                            foreach(var ann in pageVM.Annotations)
                            {
                                // 검색 결과는 저장하지 않음 (SearchResultId == -1인 것만 저장)
                                if (ann.SearchResultId != -1) continue;

                                var rect = new XRect(ann.X, ann.Y, ann.Width, ann.Height);
                                
                                if (ann.Background is SolidColorBrush solid)
                                {
                                    var c = solid.Color;
                                    var brush = new XSolidBrush(XColor.FromArgb(c.A, c.R, c.G, c.B));
                                    gfx.DrawRectangle(brush, rect);
                                }
                                if (ann.BorderThickness.Bottom > 0)
                                {
                                    gfx.DrawLine(XPens.Black, rect.BottomLeft, rect.BottomRight);
                                }
                            }
                        }
                    }
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
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