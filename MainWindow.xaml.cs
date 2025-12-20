using Microsoft.Win32;
using PdfiumViewer;
using System;
using System.IO;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing; // System.Drawing.Common 패키지
using System.Drawing.Imaging;
using Point = System.Windows.Point; // WPF Point 명시

namespace MinsPDFViewer
{
    public partial class MainWindow : Window
    {
        private PdfDocument? _pdfDoc;
        private int _currentPage = 0;
        private double _dpi = 96.0; // WPF 표준 DPI

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

            // 96 DPI로 렌더링 -> WPF 좌표계(96 DPI)와 1:1 매칭
            using (var image = _pdfDoc.Render(pageIndex, (int)_dpi, (int)_dpi, true))
            {
                // [중요] Image -> Bitmap 명시적 형변환
                var bitmap = (Bitmap)image;
                PdfImage.Source = ConvertBitmapToImageSource(bitmap);
                
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

        private void PerformSearch(string query)
        {
            if (_pdfDoc == null || string.IsNullOrWhiteSpace(query)) return;

            PdfCanvas.Children.Clear();
            int matchCount = 0;

            // 검색 수행 (대소문자 구분 X, 전체 단어 X)
            var textMatches = _pdfDoc.Search(query, false, false); 

            foreach (var match in textMatches.Items)
            {
                if (match.Page != _currentPage) continue; 

                // [중요] GetTextSegments로 정확한 좌표(Bounds) 가져오기
                var segments = _pdfDoc.GetTextSegments(match.Page, match.Index, match.Length);

                foreach (var segment in segments)
                {
                    // PDFium(Point단위 가능성) -> WPF(96DPI Pixel) 변환
                    // 96 DPI로 렌더링했으므로 (96/72) 스케일링 필요
                    double scaleFactor = _dpi / 72.0;

                    var rect = new System.Windows.Shapes.Rectangle
                    {
                        Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 255, 0)),
                        Width = segment.Bounds.Width * scaleFactor,
                        Height = segment.Bounds.Height * scaleFactor
                    };

                    Canvas.SetLeft(rect, segment.Bounds.Left * scaleFactor);
                    Canvas.SetTop(rect, segment.Bounds.Top * scaleFactor);

                    PdfCanvas.Children.Add(rect);
                }
                matchCount++;
            }
            TxtStatus.Text = $"{_currentPage + 1}페이지: {matchCount}개 발견";
        }

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