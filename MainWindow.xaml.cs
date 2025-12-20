using Microsoft.Win32;
using PdfiumViewer;
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing; // System.Drawing.Common 패키지 필요
using System.Drawing.Imaging;
using Point = System.Windows.Point; // WPF Point와 충돌 방지

namespace MinsPDFViewer
{
    public partial class MainWindow : Window
    {
        private PdfDocument? _pdfDoc;
        private int _currentPage = 0;
        private double _dpi = 96.0; // WPF 표준 DPI

        // 드래그 관련 변수
        private bool _isDragging = false;
        private Point _startPoint;
        private System.Windows.Shapes.Rectangle? _dragRect;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" };
            if (dlg.ShowDialog() == true)
            {
                LoadPdf(dlg.FileName);
            }
        }

        private void LoadPdf(string path)
        {
            try
            {
                // PDFium으로 로드 (좌표의 기준이 됨)
                _pdfDoc = PdfDocument.Load(path);
                _currentPage = 0;
                RenderPage(_currentPage);
                TxtStatus.Text = $"총 {_pdfDoc.PageCount} 페이지 중 1페이지";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 열기 실패: {ex.Message}");
            }
        }

        private void RenderPage(int pageIndex)
        {
            if (_pdfDoc == null) return;

            // [핵심] 96 DPI로 렌더링 -> 1픽셀 = 1WPF단위 (좌표 변환 불필요!)
            using (var bitmap = _pdfDoc.Render(pageIndex, (int)_dpi, (int)_dpi, true))
            {
                PdfImage.Source = ConvertBitmapToImageSource(bitmap);
                
                // 캔버스 크기를 이미지와 동일하게 맞춤
                PdfCanvas.Width = bitmap.Width;
                PdfCanvas.Height = bitmap.Height;
                PdfCanvas.Children.Clear();
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) PerformSearch(TxtSearch.Text);
        }

        // [핵심] 검색 및 하이라이트 (WYSIWYG 좌표)
        private void PerformSearch(string query)
        {
            if (_pdfDoc == null || string.IsNullOrWhiteSpace(query)) return;

            PdfCanvas.Children.Clear();
            int matchCount = 0;

            // PDFium 엔진이 찾아준 좌표를 그대로 사용
            var textMatches = _pdfDoc.Search(query, false, false); 

            foreach (var match in textMatches.Items)
            {
                if (match.Page != _currentPage) continue; 

                // PDFium 결과값은 Point 단위일 수 있으므로 DPI 비율만 맞춰줌 (96/72)
                double scaleFactor = _dpi / 72.0;

                var rect = new System.Windows.Shapes.Rectangle
                {
                    Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 255, 0)), // 반투명 노랑
                    Width = match.Bounds.Width * scaleFactor,
                    Height = match.Bounds.Height * scaleFactor
                };

                Canvas.SetLeft(rect, match.Bounds.Left * scaleFactor);
                Canvas.SetTop(rect, match.Bounds.Top * scaleFactor);

                PdfCanvas.Children.Add(rect);
                matchCount++;
            }
            TxtStatus.Text = $"{_currentPage + 1}페이지: {matchCount}개 발견";
        }

        // Bitmap -> BitmapImage 변환 헬퍼
        private BitmapImage ConvertBitmapToImageSource(Bitmap src)
        {
            var ms = new MemoryStream();
            src.Save(ms, ImageFormat.Bmp);
            var image = new BitmapImage();
            image.BeginInit();
            ms.Seek(0, SeekOrigin.Begin);
            image.StreamSource = ms;
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.EndInit();
            return image;
        }

        // --- 드래그 기능 (좌표 확인용) ---
        private void PdfCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            _isDragging = true;
            _startPoint = e.GetPosition(PdfCanvas);
            _dragRect = new System.Windows.Shapes.Rectangle
            {
                Stroke = System.Windows.Media.Brushes.Red,
                StrokeThickness = 2
            };
            Canvas.SetLeft(_dragRect, _startPoint.X);
            Canvas.SetTop(_dragRect, _startPoint.Y);
            PdfCanvas.Children.Add(_dragRect);
        }

        private void PdfCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDragging || _dragRect == null) return;
            var pos = e.GetPosition(PdfCanvas);
            var x = Math.Min(pos.X, _startPoint.X);
            var y = Math.Min(pos.Y, _startPoint.Y);
            var w = Math.Abs(pos.X - _startPoint.X);
            var h = Math.Abs(pos.Y - _startPoint.Y);
            _dragRect.Width = w; _dragRect.Height = h;
            Canvas.SetLeft(_dragRect, x); Canvas.SetTop(_dragRect, y);
        }

        private void PdfCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            _isDragging = false;
        }
    }
}