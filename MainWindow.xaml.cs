using Microsoft.Win32;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Editors; // IPageReader
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
using System.Runtime.InteropServices; // Marshal
using Point = System.Windows.Point;

namespace MinsPDFViewer
{
    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        private int _currentPage = 0;
        
        // 렌더링 스케일 (1.0 = 72 DPI 기준, 1.33 = 96 DPI, 2.0 = 고해상도)
        // 화면 선명도를 위해 2.0배로 렌더링하고 Canvas도 2배로 맞춥니다.
        private double _renderScale = 2.0; 

        public MainWindow()
        {
            InitializeComponent();
            _docLib = DocLib.Instance; // Docnet 라이브러리 초기화
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
                // 기존 리더 정리
                _docReader?.Dispose();

                // 1. Docnet으로 문서 로드
                _docReader = _docLib.GetDocReader(path, new PageDimensions(_renderScale));
                
                _currentPage = 0;
                RenderPage(_currentPage);
                TxtStatus.Text = $"총 {_docReader.GetPageCount()} 페이지 중 1페이지";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 열기 실패: {ex.Message}");
            }
        }

        private void RenderPage(int pageIndex)
        {
            if (_docReader == null) return;

            // 2. 페이지 읽기 (Docnet은 IDisposable이므로 using 사용)
            using (var pageReader = _docReader.GetPageReader(pageIndex))
            {
                // -- 이미지 렌더링 --
                var rawBytes = pageReader.GetImage(); // BGRA raw bytes
                var width = pageReader.GetPageWidth();
                var height = pageReader.GetPageHeight();

                // Raw Bytes -> BitmapImage 변환
                PdfImage.Source = RawBytesToBitmapImage(rawBytes, width, height);

                // Canvas 크기 동기화
                PdfCanvas.Width = width;
                PdfCanvas.Height = height;
                PdfCanvas.Children.Clear();
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) PerformSearch(TxtSearch.Text);
        }

        // [핵심] Docnet을 이용한 검색 및 좌표 추출 (WYSIWYG)
        private void PerformSearch(string query)
        {
            if (_docReader == null || string.IsNullOrWhiteSpace(query)) return;

            PdfCanvas.Children.Clear();
            int matchCount = 0;

            using (var pageReader = _docReader.GetPageReader(_currentPage))
            {
                // 1. 페이지 전체 텍스트 가져오기
                string pageText = pageReader.GetText();
                
                // 2. 모든 글자의 좌표 가져오기 (PDFium 엔진 원본 좌표)
                // Docnet은 렌더링 스케일(_renderScale)이 적용된 좌표를 반환합니다. (매우 편리!)
                var characters = pageReader.GetCharacters();

                // 3. 텍스트 내에서 검색어 위치 찾기 (단순 인덱스 검색)
                int index = 0;
                while ((index = pageText.IndexOf(query, index, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    // 4. 찾은 인덱스에 해당하는 글자들의 좌표(Box)를 가져와서 그리기
                    for (int i = 0; i < query.Length; i++)
                    {
                        int charIndex = index + i;
                        if (charIndex >= characters.Count) break;

                        var charInfo = characters[charIndex];
                        var box = charInfo.Box; // {Left, Top, Right, Bottom}

                        // 사각형 그리기 (좌표 변환 불필요! 이미 렌더링 스케일에 맞춰져 있음)
                        var rect = new System.Windows.Shapes.Rectangle
                        {
                            Fill = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 255, 0)),
                            Width = box.Right - box.Left,
                            Height = box.Bottom - box.Top
                        };

                        Canvas.SetLeft(rect, box.Left);
                        Canvas.SetTop(rect, box.Top);

                        PdfCanvas.Children.Add(rect);
                    }
                    
                    matchCount++;
                    index += query.Length;
                }
            }
            
            TxtStatus.Text = $"{_currentPage + 1}페이지: {matchCount}개 발견";
        }

        // Raw BGRA Bytes -> WPF BitmapImage 변환 헬퍼
        private BitmapImage RawBytesToBitmapImage(byte[] rawBytes, int width, int height)
        {
            var bitmap = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgra32, null);
            bitmap.WritePixels(new Int32Rect(0, 0, width, height), rawBytes, width * 4, 0);
            
            // WriteableBitmap은 바로 Source로 쓸 수 있지만, 호환성을 위해 BitmapImage 패턴 유지
            // (여기서는 성능을 위해 WriteableBitmap을 그대로 반환해도 되지만 Image.Source 타입 맞춤)
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

        // --- 드래그 기능 ---
        private void PdfCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            // 드래그 로직은 동일
        }
        private void PdfCanvas_MouseMove(object sender, MouseEventArgs e) { }
        private void PdfCanvas_MouseUp(object sender, MouseButtonEventArgs e) { }
    }
}