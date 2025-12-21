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
    // [시스템 폰트 리졸버]
    public class WindowsFontResolver : IFontResolver
    {
        public byte[]? GetFont(string faceName)
        {
            string fontPath = @"C:\Windows\Fonts\malgun.ttf";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            fontPath = @"C:\Windows\Fonts\gulim.ttc";
            if (File.Exists(fontPath)) return File.ReadAllBytes(fontPath);
            return null;
        }
        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new FontResolverInfo("Malgun Gothic");
    }

    public enum AnnotationType { Highlight, Underline, SearchHighlight, Other }

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

    public class OcrWordInfo
    {
        public string Text { get; set; } = "";
        public Rect BoundingBox { get; set; }
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
        private bool _isSelecting;
        public bool IsSelecting { get => _isSelecting; set { _isSelecting = value; OnPropertyChanged(nameof(IsSelecting)); } }
        private double _selX; public double SelectionX { get => _selX; set { _selX = value; OnPropertyChanged(nameof(SelectionX)); } }
        private double _selY; public double SelectionY { get => _selY; set { _selY = value; OnPropertyChanged(nameof(SelectionY)); } }
        private double _selW; public double SelectionWidth { get => _selW; set { _selW = value; OnPropertyChanged(nameof(SelectionWidth)); } }
        private double _selH; public double SelectionHeight { get => _selH; set { _selH = value; OnPropertyChanged(nameof(SelectionHeight)); } }
        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class CustomPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public CustomPdfAnnotation(PdfDocument document) : base(document) { }
    }

    public partial class MainWindow : Window
    {
        private IDocLib _docLib;
        private IDocReader? _docReader;
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        private double _renderScale = 2.0; 
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
        private OcrEngine? _ocrEngine;

        public MainWindow()
        {
            InitializeComponent();
            try { if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new WindowsFontResolver(); } catch { }
            _docLib = DocLib.Instance;
            PdfListView.ItemsSource = Pages;
            try { _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR")) ?? OcrEngine.TryCreateFromUserProfileLanguages(); } catch { }
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
                }
                
                _docReader?.Dispose();
                _docReader = _docLib.GetDocReader(_originalFileBytes, new PageDimensions(_renderScale));
                int pc = _docReader.GetPageCount();

                for(int i=0; i<pc; i++)
                {
                    using(var r = _docReader.GetPageReader(i))
                    {
                        Pages.Add(new PdfPageViewModel { PageIndex=i, Width=r.GetPageWidth(), Height=r.GetPageHeight() });
                    }
                }

                Task.Run(() => RenderAllPagesAsync(pc));
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
                        if (i < Pages.Count) Pages[i].ImageSource = RawBytesToBitmapImage(bytes, w, h);
                    });
                }
            });
            Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "로딩 완료");
        }

        private async void BtnOCR_Click(object sender, RoutedEventArgs e)
        {
            if (_ocrEngine == null) { MessageBox.Show("OCR 미지원"); return; }
            if (Pages.Count == 0) return;

            BtnOCR.IsEnabled = false; PbStatus.Visibility = Visibility.Visible;
            PbStatus.Maximum = Pages.Count; PbStatus.Value = 0; TxtStatus.Text = "OCR 분석 중...";

            try
            {
                await Task.Run(async () =>
                {
                    for (int i = 0; i < Pages.Count; i++)
                    {
                        var pageVM = Pages[i];
                        if (_docReader == null) continue;

                        using (var r = _docReader.GetPageReader(i))
                        {
                            var rawBytes = r.GetImage();
                            var w = r.GetPageWidth();
                            var h = r.GetPageHeight();

                            using (var stream = new MemoryStream(rawBytes))
                            {
                                var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied);
                                softwareBitmap.CopyFromBuffer(rawBytes.AsBuffer());

                                var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);

                                var wordList = new List<OcrWordInfo>();
                                foreach (var line in ocrResult.Lines)
                                {
                                    foreach (var word in line.Words)
                                    {
                                        wordList.Add(new OcrWordInfo 
                                        { 
                                            Text = word.Text, 
                                            BoundingBox = new Rect(word.BoundingRect.X, word.BoundingRect.Y, word.BoundingRect.Width, word.BoundingRect.Height)
                                        });
                                    }
                                }
                                pageVM.OcrWords = wordList;
                            }
                        }
                        Application.Current.Dispatcher.Invoke(() => { PbStatus.Value = i + 1; });
                    }
                });
                TxtStatus.Text = "OCR 완료. 저장하세요.";
            }
            catch (Exception ex) { MessageBox.Show($"OCR 오류: {ex.Message}"); }
            finally { BtnOCR.IsEnabled = true; PbStatus.Visibility = Visibility.Collapsed; }
        }

        private void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            _selectedTextBuffer = "";
            _selectedChars.Clear();
            if (_docReader == null) return;

            var sb = new StringBuilder();

            using (var reader = _docReader.GetPageReader(pageIndex))
            {
                var chars = reader.GetCharacters().ToList();
                foreach (var c in chars)
                {
                    var r = new Rect(Math.Min(c.Box.Left, c.Box.Right), Math.Min(c.Box.Top, c.Box.Bottom), 
                                     Math.Abs(c.Box.Right - c.Box.Left), Math.Abs(c.Box.Bottom - c.Box.Top));
                    
                    if (uiRect.IntersectsWith(r))
                    {
                        _selectedChars.Add(c);
                        sb.Append(c.Char);
                    }
                }
            }

            var pageVM = Pages[pageIndex];
            if (pageVM.OcrWords != null)
            {
                foreach (var word in pageVM.OcrWords)
                {
                    if (uiRect.IntersectsWith(word.BoundingBox))
                    {
                        sb.Append(word.Text + " "); 
                    }
                }
            }

            _selectedTextBuffer = sb.ToString();
        }

        // [수정됨] 저장 로직: 투명 텍스트 처리 개선
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
                        var pdfPage = doc.Pages[i];
                        var pageVM = Pages[i];
                        
                        double scaleX = pdfPage.Width.Point / pageVM.Width;
                        double scaleY = pdfPage.Height.Point / pageVM.Height;

                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0)
                        {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                foreach (var word in pageVM.OcrWords)
                                {
                                    double x = word.BoundingBox.X * scaleX;
                                    double y = word.BoundingBox.Y * scaleY;
                                    double w = word.BoundingBox.Width * scaleX;
                                    double h = word.BoundingBox.Height * scaleY;

                                    double fontSize = h * 0.75;
                                    if (fontSize < 1) fontSize = 1;

                                    var font = new XFont("Malgun Gothic", fontSize, XFontStyleEx.Regular);
                                    
                                    // [핵심 수정] Alpha=1 (아주 연한 흰색)으로 설정하여 "검은색 박스" 문제 해결
                                    // 0으로 설정 시 일부 뷰어에서 검은색으로 폴백되는 문제 방지
                                    var brush = new XSolidBrush(XColor.FromArgb(1, 255, 255, 255));
                                    
                                    gfx.DrawString(word.Text, font, brush, 
                                        new XRect(x, y, w, h), XStringFormats.Center);
                                }
                            }
                        }

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;

                            double ax = ann.X * scaleX;
                            double ay = ann.Y * scaleY;
                            double aw = ann.Width * scaleX;
                            double ah = ann.Height * scaleY;

                            var pdfAnnot = new CustomPdfAnnotation(doc);
                            double pdfY_BottomUp = pdfPage.Height.Point - (ay + ah);
                            
                            if (ann.Type == AnnotationType.Underline) pdfY_BottomUp = pdfPage.Height.Point - (ay + 2);

                            pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
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

        private void PerformSearch(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return;
            _lastSearchQuery = query; _searchResults.Clear();
            foreach (var p in Pages) {
                var tr = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                foreach (var i in tr) p.Annotations.Remove(i);
            }

            for (int i = 0; i < Pages.Count; i++)
            {
                var p = Pages[i];
                if (p.OcrWords != null)
                {
                    foreach (var word in p.OcrWords)
                    {
                        if (word.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                        {
                            var a = new PdfAnnotation
                            {
                                X = word.BoundingBox.X, Y = word.BoundingBox.Y,
                                Width = word.BoundingBox.Width, Height = word.BoundingBox.Height,
                                Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255)),
                                Type = AnnotationType.SearchHighlight
                            };
                            p.Annotations.Add(a); _searchResults.Add(a);
                        }
                    }
                }
            }
            
            if (_docReader != null) {
                 int pc=_docReader.GetPageCount();
                 for(int i=0;i<pc;i++){ using(var r=_docReader.GetPageReader(i)){
                     string t=r.GetText(); var cs=r.GetCharacters().ToList(); int idx=0;
                     while((idx=t.IndexOf(query,idx,StringComparison.OrdinalIgnoreCase))!=-1){
                         double mx=double.MaxValue, my=double.MaxValue, Mx=double.MinValue, My=double.MinValue;
                         for(int c=0;c<query.Length;c++){ if(idx+c<cs.Count){ var b=cs[idx+c].Box; mx=Math.Min(mx,b.Left); my=Math.Min(my,b.Top); Mx=Math.Max(Mx,b.Right); My=Math.Max(My,b.Bottom); } }
                         var a=new PdfAnnotation{X=mx,Y=my,Width=Mx-mx,Height=My-my,Background=new SolidColorBrush(Color.FromArgb(60,0,255,255)),Type=AnnotationType.SearchHighlight};
                         Pages[i].Annotations.Add(a); _searchResults.Add(a); idx+=query.Length;
                     }
                 }}
             }

            TxtStatus.Text = $"검색: {_searchResults.Count}건";
            if (_searchResults.Count > 0) { _currentSearchIndex = -1; NavigateSearchResult(true); }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (!string.IsNullOrEmpty(_currentFilePath)) SavePdf(_currentFilePath); }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(_currentFilePath)) return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(_currentFilePath) + "_copy" };
            if (dlg.ShowDialog() == true) { SavePdf(dlg.FileName); _currentFilePath = dlg.FileName; }
        }
        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e) { if (!string.IsNullOrEmpty(_selectedTextBuffer)) { Clipboard.SetText(_selectedTextBuffer); TxtStatus.Text = "텍스트 복사됨"; } CloseSelection(); }
        private void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e) { 
            if(_selectedPageIndex==-1)return; var p=Pages[_selectedPageIndex]; var bmp=p.ImageSource as BitmapSource; if(bmp==null)return;
            double sx=bmp.PixelWidth/p.Width; double sy=bmp.PixelHeight/p.Height;
            int x=(int)(p.SelectionX*sx); int y=(int)(p.SelectionY*sy); int w=(int)(p.SelectionWidth*sx); int h=(int)(p.SelectionHeight*sy);
            if(w>0&&h>0 && x>=0 && y>=0 && x+w<=bmp.PixelWidth && y+h<=bmp.PixelHeight){ 
                try{Clipboard.SetImage(new CroppedBitmap(bmp,new Int32Rect(x,y,w,h))); TxtStatus.Text="이미지 복사됨";}catch{TxtStatus.Text="복사 실패";} 
            } CloseSelection();
        }
        private void Page_MouseDown(object sender, MouseButtonEventArgs e) {
            if (_currentTool != "CURSOR") return; var canvas = sender as Canvas; if (canvas == null) return;
            foreach (var p in Pages) { p.IsSelecting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; } SelectionPopup.IsOpen = false;
            _activePageIndex = (int)canvas.Tag; _dragStartPoint = e.GetPosition(canvas);
            var pageVM = Pages[_activePageIndex]; pageVM.IsSelecting = true; pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y; pageVM.SelectionWidth = 0; pageVM.SelectionHeight = 0;
            canvas.CaptureMouse(); e.Handled = true; TxtStatus.Text = $"드래그 시작: {_activePageIndex + 1} 페이지";
        }
        private void Page_MouseMove(object sender, MouseEventArgs e) {
            if (_activePageIndex == -1) return; var canvas = sender as Canvas; if (canvas == null) return;
            var pt = e.GetPosition(canvas); var p = Pages[_activePageIndex];
            double x=Math.Min(_dragStartPoint.X,pt.X), y=Math.Min(_dragStartPoint.Y,pt.Y), w=Math.Abs(pt.X-_dragStartPoint.X), h=Math.Abs(pt.Y-_dragStartPoint.Y);
            p.SelectionX=x; p.SelectionY=y; p.SelectionWidth=w; p.SelectionHeight=h; TxtStatus.Text = $"선택 중.. {Math.Round(w)}x{Math.Round(h)}";
        }
        private void Page_MouseUp(object sender, MouseButtonEventArgs e) {
            var canvas = sender as Canvas; if (canvas == null || _activePageIndex == -1) return;
            canvas.ReleaseMouseCapture(); var p = Pages[_activePageIndex];
            if (p.IsSelecting) {
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5) {
                    var rect = new Rect(p.SelectionX, p.SelectionY, p.SelectionWidth, p.SelectionHeight);
                    CheckTextInSelection(_activePageIndex, rect);
                    SelectionPopup.PlacementTarget = canvas; SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(canvas).X, e.GetPosition(canvas).Y+10,0,0); SelectionPopup.IsOpen = true; _selectedPageIndex = _activePageIndex;
                    if(_selectedChars.Count>0 || !string.IsNullOrEmpty(_selectedTextBuffer)) TxtStatus.Text=$"텍스트 선택됨"; else TxtStatus.Text="영역 선택됨";
                } else { p.IsSelecting = false; TxtStatus.Text="준비"; }
            } _activePageIndex = -1; e.Handled = true;
        }
        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);
        private void AddAnnotation(Color color, AnnotationType type) {
            if (_selectedPageIndex == -1) return; var p = Pages[_selectedPageIndex];
            var ann = new PdfAnnotation { X=p.SelectionX, Y=p.SelectionY, Width=p.SelectionWidth, Height=p.SelectionHeight, Type=type, AnnotationColor=color };
            if(type==AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80,color.R,color.G,color.B));
            else { ann.Background = new SolidColorBrush(color); ann.Height=2; ann.Y=p.SelectionY+p.SelectionHeight-2; }
            p.Annotations.Add(ann); CloseSelection();
        }
        private void CloseSelection() { SelectionPopup.IsOpen = false; foreach(var p in Pages) p.IsSelecting = false; }
        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e) {
            var mi=sender as MenuItem; if(mi!=null && mi.CommandParameter is PdfAnnotation a) foreach(var p in Pages) if(p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; }
        }
        private void BtnSearch_Click(object sender, RoutedEventArgs e) => PerformSearch(TxtSearch.Text);
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(false);
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(true);
        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) { if(e.Key==Key.Enter){ PerformSearch(TxtSearch.Text); } }
        private void NavigateSearchResult(bool n) {
            if(_searchResults.Count==0)return;
            if(_currentSearchIndex>=0&&_currentSearchIndex<_searchResults.Count)_searchResults[_currentSearchIndex].Background=new SolidColorBrush(Color.FromArgb(60,0,255,255));
            if(n){_currentSearchIndex++;if(_currentSearchIndex>=_searchResults.Count)_currentSearchIndex=0;} else{_currentSearchIndex--;if(_currentSearchIndex<0)_currentSearchIndex=_searchResults.Count-1;}
            var c=_searchResults[_currentSearchIndex]; c.Background=new SolidColorBrush(Color.FromArgb(120,255,0,255));
            var tp=Pages.FirstOrDefault(p=>p.Annotations.Contains(c)); if(tp!=null)PdfListView.ScrollIntoView(tp);
            TxtStatus.Text=$"검색: {_currentSearchIndex+1} / {_searchResults.Count}";
        }
        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if((Keyboard.Modifiers&ModifierKeys.Control)==ModifierKeys.Control){ if(e.Delta>0)UpdateZoom(_currentZoom+0.1);else UpdateZoom(_currentZoom-0.1); e.Handled=true; } }
        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom+0.1);
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => UpdateZoom(_currentZoom-0.1);
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) => FitWidth();
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) => FitHeight();
        private void UpdateZoom(double z) { _currentZoom=Math.Clamp(z,0.2,5.0); ViewScaleTransform.ScaleX=_currentZoom; ViewScaleTransform.ScaleY=_currentZoom; TxtZoom.Text=$"{Math.Round(_currentZoom*100)}%"; }
        private void FitWidth() { if(Pages.Count>0 && Pages[0].Width>0) UpdateZoom((PdfListView.ActualWidth-60)/Pages[0].Width); }
        private void FitHeight() { if(Pages.Count>0 && Pages[0].Height>0) UpdateZoom((PdfListView.ActualHeight-60)/Pages[0].Height); }
        private void Tool_Click(object sender, RoutedEventArgs e) { if(RbCursor.IsChecked==true)_currentTool="CURSOR"; else if(RbHighlight.IsChecked==true)_currentTool="HIGHLIGHT"; else if(RbText.IsChecked==true)_currentTool="TEXT"; }
        private void Window_KeyDown(object sender, KeyEventArgs e){}
        private BitmapImage RawBytesToBitmapImage(byte[] b,int w,int h) { var bm=new WriteableBitmap(w,h,96,96,PixelFormats.Bgra32,null); bm.WritePixels(new Int32Rect(0,0,w,h),b,w*4,0); if(bm.CanFreeze)bm.Freeze(); return ConvertWriteableBitmapToBitmapImage(bm); }
        private BitmapImage ConvertWriteableBitmapToBitmapImage(WriteableBitmap wbm) { using(var ms=new MemoryStream()){ var enc=new PngBitmapEncoder(); enc.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(wbm)); enc.Save(ms); var img=new BitmapImage(); img.BeginInit(); img.CacheOption=BitmapCacheOption.OnLoad; img.StreamSource=ms; img.EndInit(); if(img.CanFreeze)img.Freeze(); return img; } }
    }
}