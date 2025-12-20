using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Editors;
using Docnet.Core.Readers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing; // Bitmap
using System.Drawing.Imaging;
using Point = System.Windows.Point;
using Brush = System.Windows.Media.Brush;
using Brushes = System.Windows.Media.Brushes;
using Pen = System.Windows.Media.Pen;
using Color = System.Windows.Media.Color;

namespace MinsPDFViewer
{
    // 데이터 모델: 주석 정보
    public class AnnotationData
    {
        public string Type { get; set; } = "HIGHLIGHT"; // HIGHLIGHT, TEXT
        public Rect Bounds { get; set; }
        public string TextContent { get; set; } = "";
        public int PageIndex { get; set; }
        public bool IsSelected { get; set; } = false;
    }

    public partial class MainWindow : Window
    {
        // 엔진
        private IDocLib _docLib;
        private IDocReader? _docReader;
        
        // 상태 변수
        private string _currentFilePath = "";
        private int _currentPage = 0;
        private double _renderScale = 1.5; // 1.5배 확대 렌더링 (선명도)
        
        // 주석 데이터 (메모리에 저장)
        private List<AnnotationData> _annotations = new List<AnnotationData>();
        
        // 그리기 도구 상태
        private string _currentTool = "CURSOR"; // CURSOR, HIGHLIGHT, TEXT
        private Point _startPoint;
        private bool _isDrawing = false;
        private bool _isDragging = false;
        private AnnotationData? _selectedAnnotation = null;

        public MainWindow()
        {
            InitializeComponent();
            _docLib = DocLib.Instance;
        }

        // --- 1. 파일 열기 & 렌더링 ---
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
                _currentFilePath = path;
                _currentPage = 0;
                _annotations.Clear(); // 새 파일 열면 주석 초기화
                RenderPage(_currentPage);
                TxtStatus.Text = $"총 {_docReader.GetPageCount()} 페이지 중 1페이지";
            }
            catch (Exception ex) { MessageBox.Show($"열기 실패: {ex.Message}"); }
        }

        private void RenderPage(int pageIndex)
        {
            if (_docReader == null) return;

            using (var pageReader = _docReader.GetPageReader(pageIndex))
            {
                // 이미지 생성
                var rawBytes = pageReader.GetImage();
                var width = pageReader.GetPageWidth();
                var height = pageReader.GetPageHeight();

                PdfImage.Source = RawBytesToBitmapImage(rawBytes, width, height);
                PdfCanvas.Width = width;
                PdfCanvas.Height = height;
                
                // 기존 주석 다시 그리기
                DrawAnnotations();
            }
        }

        // --- 2. 주석 그리기 (화면 표시용) ---
        private void DrawAnnotations()
        {
            PdfCanvas.Children.Clear();
            
            // 현재 페이지의 주석만 필터링
            var pageAnns = _annotations.Where(a => a.PageIndex == _currentPage).ToList();

            foreach (var ann in pageAnns)
            {
                UIElement element = null;

                if (ann.Type == "HIGHLIGHT")
                {
                    var rect = new System.Windows.Shapes.Rectangle
                    {
                        Width = ann.Bounds.Width,
                        Height = ann.Bounds.Height,
                        Fill = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)), // 반투명 노랑
                        Stroke = ann.IsSelected ? Brushes.Blue : Brushes.Transparent,
                        StrokeThickness = 2
                    };
                    element = rect;
                }
                else if (ann.Type == "TEXT")
                {
                    var border = new Border
                    {
                        Width = ann.Bounds.Width,
                        Height = ann.Bounds.Height,
                        BorderBrush = ann.IsSelected ? Brushes.Blue : Brushes.Transparent,
                        BorderThickness = new Thickness(ann.IsSelected ? 1 : 0),
                        Background = Brushes.Transparent
                    };
                    var tb = new TextBlock
                    {
                        Text = ann.TextContent,
                        FontSize = 14 * _renderScale * 0.7, // 스케일 보정
                        Foreground = Brushes.Black,
                        TextWrapping = TextWrapping.Wrap
                    };
                    border.Child = tb;
                    element = border;
                }

                if (element != null)
                {
                    Canvas.SetLeft(element, ann.Bounds.X);
                    Canvas.SetTop(element, ann.Bounds.Y);
                    // 클릭 이벤트 연결 (선택용)
                    element.MouseLeftButtonDown += (s, e) => 
                    { 
                        if (_currentTool == "CURSOR") 
                        {
                            e.Handled = true; 
                            SelectAnnotation(ann); 
                        }
                    };
                    PdfCanvas.Children.Add(element);
                }
            }
        }

        // --- 3. 마우스 조작 (그리기 및 선택) ---
        private void PdfCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_docReader == null) return;
            _startPoint = e.GetPosition(PdfCanvas);

            if (_currentTool == "CURSOR")
            {
                // 빈 곳을 찍으면 선택 해제
                if (e.OriginalSource == PdfCanvas) SelectAnnotation(null);
                else if (_selectedAnnotation != null) _isDragging = true; // 선택된거 드래그 시작
            }
            else
            {
                _isDrawing = true;
            }
        }

        private void PdfCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging && _selectedAnnotation != null)
            {
                // 이동 로직
                var pos = e.GetPosition(PdfCanvas);
                double dx = pos.X - _startPoint.X;
                double dy = pos.Y - _startPoint.Y;
                
                var newRect = _selectedAnnotation.Bounds;
                newRect.X += dx; newRect.Y += dy;
                _selectedAnnotation.Bounds = newRect;
                
                _startPoint = pos;
                DrawAnnotations(); // 다시 그리기
            }
        }

        private void PdfCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDrawing)
            {
                var endPoint = e.GetPosition(PdfCanvas);
                var rect = new Rect(_startPoint, endPoint);

                if (rect.Width > 5 && rect.Height > 5)
                {
                    var newAnn = new AnnotationData
                    {
                        PageIndex = _currentPage,
                        Bounds = rect,
                        Type = _currentTool == "HIGHLIGHT" ? "HIGHLIGHT" : "TEXT",
                        TextContent = _currentTool == "TEXT" ? "텍스트 입력" : "" // 기본 텍스트
                    };
                    
                    if (_currentTool == "TEXT")
                    {
                        // 텍스트 입력 받기 (간단히 InputBox 대신 기본값)
                        // 실제로는 여기서 팝업을 띄워 입력을 받아야 함
                        newAnn.TextContent = "여기에 텍스트"; 
                    }

                    _annotations.Add(newAnn);
                    DrawAnnotations();
                }
                _isDrawing = false;
            }
            _isDragging = false;
        }

        // --- 4. 저장 기능 (PdfSharp 연동) ---
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_currentFilePath)) return;
            
            string savePath = _currentFilePath.Replace(".pdf", "_edited.pdf");
            
            try
            {
                // 원본 파일 열기 (PdfSharp)
                using (var doc = PdfReader.Open(_currentFilePath, PdfDocumentOpenMode.Modify))
                {
                    // 각 페이지별로 주석 저장
                    foreach (var ann in _annotations)
                    {
                        if (ann.PageIndex >= doc.PageCount) continue;
                        var page = doc.Pages[ann.PageIndex];

                        // 좌표 변환: Docnet(Visual) -> PdfSharp(Point)
                        // Docnet은 _renderScale 배율로 렌더링되었으므로 나눠줘야 함
                        // 또한 PdfSharp은 72dpi 기준
                        
                        // 비율 계산: PdfSharp Page Width / Canvas Width
                        double ratioX = page.Width.Point / PdfCanvas.Width;
                        double ratioY = page.Height.Point / PdfCanvas.Height;

                        // PDF 좌표계 (Y축 반전 고려)
                        // PdfSharp DrawString 등은 Top-Left 기준이므로 Y축 반전 불필요 (XGraphics 사용시)
                        
                        using (var gfx = XGraphics.FromPdfPage(page))
                        {
                            var x = ann.Bounds.X * ratioX;
                            var y = ann.Bounds.Y * ratioY;
                            var w = ann.Bounds.Width * ratioX;
                            var h = ann.Bounds.Height * ratioY;

                            if (ann.Type == "HIGHLIGHT")
                            {
                                // 형광펜: 투명도 지원 브러시
                                var brush = new XSolidBrush(XColor.FromArgb(80, 255, 255, 0));
                                gfx.DrawRectangle(brush, x, y, w, h);
                            }
                            else if (ann.Type == "TEXT")
                            {
                                // 텍스트
                                var font = new XFont("Arial", 12, XFontStyleEx.Regular);
                                gfx.DrawString(ann.TextContent, font, XBrushes.Black, new XRect(x, y, w, h), XStringFormats.TopLeft);
                            }
                        }
                    }
                    
                    doc.Save(savePath);
                }
                MessageBox.Show($"저장 완료!\n{savePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"저장 실패: {ex.Message}");
            }
        }

        // --- 헬퍼 메서드 ---
        private void SelectAnnotation(AnnotationData? ann)
        {
            if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
            _selectedAnnotation = ann;
            if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = true;
            DrawAnnotations();
            
            // 삭제를 위한 포커스 설정
            this.Focus();
        }

        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (RbCursor.IsChecked == true) _currentTool = "CURSOR";
            else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT";
            else if (RbText.IsChecked == true) _currentTool = "TEXT";
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            // Delete 키로 삭제
            if (e.Key == Key.Delete && _selectedAnnotation != null)
            {
                _annotations.Remove(_selectedAnnotation);
                _selectedAnnotation = null;
                DrawAnnotations();
            }
            // 검색 (기존 기능 유지)
            if (e.Key == Key.Enter && TxtSearch.IsFocused) PerformSearch(TxtSearch.Text);
        }

        // 검색 로직 (기존 Docnet 로직 유지)
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter) PerformSearch(TxtSearch.Text); }

        private void PerformSearch(string query)
        {
            if (_docReader == null || string.IsNullOrWhiteSpace(query)) return;
            PdfCanvas.Children.Clear();
            _annotations.Clear(); // 검색 시 기존 수동 주석과 섞이지 않게 처리 (원하면 유지 가능)
            
            using (var pageReader = _docReader.GetPageReader(_currentPage))
            {
                string pageText = pageReader.GetText();
                var characters = pageReader.GetCharacters().ToList();
                int index = 0;
                while ((index = pageText.IndexOf(query, index, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    // 검색된 단어의 바운딩 박스 계산
                    double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
                    for (int i = 0; i < query.Length; i++)
                    {
                        if (index + i < characters.Count)
                        {
                            var b = characters[index + i].Box;
                            minX = Math.Min(minX, b.Left); minY = Math.Min(minY, b.Top);
                            maxX = Math.Max(maxX, b.Right); maxY = Math.Max(maxY, b.Bottom);
                        }
                    }
                    
                    // 주석으로 추가
                    _annotations.Add(new AnnotationData 
                    { 
                        Type = "HIGHLIGHT", 
                        Bounds = new Rect(minX, minY, maxX - minX, maxY - minY),
                        PageIndex = _currentPage
                    });

                    index += query.Length;
                }
            }
            DrawAnnotations();
            TxtStatus.Text = "검색 완료 (노란색 표시)";
        }

        private BitmapImage RawBytesToBitmapImage(byte[] rawBytes, int width, int height)
        {
            var bitmap = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgra32, null);
            bitmap.WritePixels(new Int32Rect(0, 0, width, height), rawBytes, width * 4, 0);
            return ConvertWriteableBitmapToBitmapImage(bitmap);
        }

        private BitmapImage ConvertWriteableBitmapToBitmapImage(WriteableBitmap wbm)
        {
            using (var stream = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(wbm));
                encoder.Save(stream);
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.StreamSource = stream;
                img.EndInit();
                return img;
            }
        }
    }
}