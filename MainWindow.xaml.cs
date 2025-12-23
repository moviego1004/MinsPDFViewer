using System;
using System.Collections.Generic;
using System.Linq;
using System.Text; 
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.IO;
using System.Threading.Tasks;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Security.Cryptography;
using Windows.Globalization;

namespace MinsPDFViewer
{
    public partial class MainWindow : Window, System.ComponentModel.INotifyPropertyChanged
    {
        private readonly PdfService _pdfService;
        private readonly SearchService _searchService;
        private OcrEngine? _ocrEngine;

        public System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel> Documents { get; set; } 
            = new System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel>();
        
        private PdfDocumentModel? _selectedDocument;
        public PdfDocumentModel? SelectedDocument
        {
            get => _selectedDocument;
            set { _selectedDocument = value; OnPropertyChanged(nameof(SelectedDocument)); CheckToolbarVisibility(); }
        }

        private List<PdfAnnotation> _searchResults = new List<PdfAnnotation>();
        private int _currentSearchIndex = -1;
        private string _lastSearchQuery = "";

        private string _currentTool = "CURSOR";
        private PdfAnnotation? _selectedAnnotation = null;
        private Point _dragStartPoint;
        private int _activePageIndex = -1;
        private string _selectedTextBuffer = "";
        private int _selectedPageIndex = -1;
        private bool _isDraggingAnnotation = false;
        private Point _annotationDragStartOffset;
        private bool _isUpdatingUiFromSelection = false;
        
        private string _defaultFontFamily = "Malgun Gothic"; 
        private double _defaultFontSize = 14;
        private Color _defaultFontColor = Colors.Red;
        private bool _defaultIsBold = false;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            _pdfService = new PdfService();
            _searchService = new SearchService();
            
            try { 
                if (PdfSharp.Fonts.GlobalFontSettings.FontResolver == null) 
                    PdfSharp.Fonts.GlobalFontSettings.FontResolver = new WindowsFontResolver(); 
            } catch { }

            try { _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR")) ?? OcrEngine.TryCreateFromUserProfileLanguages(); } catch { }

            CbFont.ItemsSource = new string[] { "Malgun Gothic", "Gulim", "Dotum", "Batang" };
            CbFont.SelectedIndex = 0;
            CbSize.ItemsSource = new double[] { 10, 12, 14, 16, 18, 24, 32, 48 };
            CbSize.SelectedIndex = 2;
        }

        // [복구] 필수 이벤트 핸들러
        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_selectedAnnotation != null) { 
                _selectedAnnotation.IsSelected = false; 
                _selectedAnnotation = null; 
                CheckToolbarVisibility(); 
            }
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
            if (scrollViewer != null) {
                scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
            }
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer sv && sv.DataContext is PdfDocumentModel doc) {
                doc.SavedVerticalOffset = sv.VerticalOffset;
                doc.SavedHorizontalOffset = sv.HorizontalOffset;
            }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "PDF Files|*.pdf" };
            if (dlg.ShowDialog() == true) {
                var docModel = _pdfService.LoadPdf(dlg.FileName);
                if (docModel != null) {
                    Documents.Add(docModel);
                    SelectedDocument = docModel;
                    _ = _pdfService.RenderPagesAsync(docModel);
                }
            }
        }

        private void BtnCloseTab_Click(object sender, RoutedEventArgs e) {
            if (sender is Button btn && btn.Tag is PdfDocumentModel doc) {
                doc.DocReader?.Dispose(); Documents.Remove(doc); if (Documents.Count == 0) SelectedDocument = null;
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e) => DoSearch(TxtSearch.Text);

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e) {
            if (e.Key == Key.Enter) {
                bool isShift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                if (TxtSearch.Text == _lastSearchQuery && _searchResults.Count > 0) NavigateSearchResult(!isShift);
                else DoSearch(TxtSearch.Text);
            }
        }

        private void DoSearch(string query) {
            if (SelectedDocument == null) return;
            _lastSearchQuery = query;
            _searchResults = _searchService.PerformSearch(SelectedDocument, query);
            TxtStatus.Text = $"검색 결과: {_searchResults.Count}건";
            if (_searchResults.Count > 0) { _currentSearchIndex = -1; NavigateSearchResult(true); }
            else MessageBox.Show("검색 결과가 없습니다.");
        }

        private void NavigateSearchResult(bool next) {
            if (_searchResults.Count == 0) return;
            if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
                _searchResults[_currentSearchIndex].Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255));

            if (next) { _currentSearchIndex++; if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0; }
            else { _currentSearchIndex--; if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1; }

            var currentAnnot = _searchResults[_currentSearchIndex];
            currentAnnot.Background = new SolidColorBrush(Color.FromArgb(120, 255, 0, 255));

            if (SelectedDocument != null) {
                var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.Annotations.Contains(currentAnnot));
                if (targetPage != null) {
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null) listView.ScrollIntoView(targetPage);
                    TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}";
                }
            }
        }

        private void Page_MouseDown(object sender, MouseButtonEventArgs e) {
            if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; }
            if (SelectedDocument == null) return;
            var canvas = sender as Canvas; if (canvas == null) return;
            _activePageIndex = (int)canvas.Tag; _dragStartPoint = e.GetPosition(canvas);
            var pageVM = SelectedDocument.Pages[_activePageIndex];
            if (_currentTool == "TEXT") {
                var newAnnot = new PdfAnnotation { Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50, FontSize = _defaultFontSize, FontFamily = _defaultFontFamily, Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold, TextContent = "", IsSelected = true };
                pageVM.Annotations.Add(newAnnot); _selectedAnnotation = newAnnot; _currentTool = "CURSOR"; RbCursor.IsChecked = true; UpdateToolbarFromAnnotation(_selectedAnnotation); CheckToolbarVisibility(); e.Handled = true; return;
            }
            if (_currentTool == "CURSOR") {
                foreach (var p in SelectedDocument.Pages) { p.IsSelecting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                SelectionPopup.IsOpen = false; pageVM.IsSelecting = true; pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y; pageVM.SelectionWidth = 0; pageVM.SelectionHeight = 0; canvas.CaptureMouse(); e.Handled = true;
            }
            CheckToolbarVisibility();
        }

        private void Page_MouseMove(object sender, MouseEventArgs e) {
            if (_activePageIndex == -1 || SelectedDocument == null) return; 
            var canvas = sender as Canvas; if (canvas == null) return;
            var pageVM = SelectedDocument.Pages[_activePageIndex];
            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null && _selectedAnnotation.Type == AnnotationType.FreeText) {
                var currentPoint = e.GetPosition(canvas); _selectedAnnotation.X = currentPoint.X - _annotationDragStartOffset.X; _selectedAnnotation.Y = currentPoint.Y - _annotationDragStartOffset.Y; e.Handled = true; return;
            }
            if (_currentTool == "CURSOR" && pageVM.IsSelecting) {
                var pt = e.GetPosition(canvas); double x = Math.Min(_dragStartPoint.X, pt.X); double y = Math.Min(_dragStartPoint.Y, pt.Y); double w = Math.Abs(pt.X - _dragStartPoint.X); double h = Math.Abs(pt.Y - _dragStartPoint.Y); pageVM.SelectionX = x; pageVM.SelectionY = y; pageVM.SelectionWidth = w; pageVM.SelectionHeight = h;
            }
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e) {
            var canvas = sender as Canvas; if (canvas == null || _activePageIndex == -1 || SelectedDocument == null) return;
            canvas.ReleaseMouseCapture(); 
            var p = SelectedDocument.Pages[_activePageIndex];
            if (p.IsSelecting && _currentTool == "CURSOR") {
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5) {
                    var rect = new Rect(p.SelectionX, p.SelectionY, p.SelectionWidth, p.SelectionHeight); CheckTextInSelection(_activePageIndex, rect); SelectionPopup.PlacementTarget = canvas; SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(canvas).X, e.GetPosition(canvas).Y + 10, 0, 0); SelectionPopup.IsOpen = true; _selectedPageIndex = _activePageIndex; TxtStatus.Text = string.IsNullOrEmpty(_selectedTextBuffer) ? "영역 선택됨" : "텍스트 선택됨";
                } else { p.IsSelecting = false; TxtStatus.Text = "준비"; if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; } }
            }
            _activePageIndex = -1; _isDraggingAnnotation = false; e.Handled = true; CheckToolbarVisibility(); 
        }

        private void Annotation_PreviewMouseDown(object sender, MouseButtonEventArgs e) {
            if (_currentTool != "CURSOR") return; var element = sender as FrameworkElement;
            if (element?.DataContext is PdfAnnotation ann) {
                if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false; _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true;
                if (ann.Type == AnnotationType.FreeText) { _isDraggingAnnotation = true; _annotationDragStartOffset = e.GetPosition(element); } else { _isDraggingAnnotation = false; }
                UpdateToolbarFromAnnotation(ann); CheckToolbarVisibility(); e.Handled = true; 
            }
        }

        private void BtnOCR_Click(object sender, RoutedEventArgs e) { if (_ocrEngine == null || SelectedDocument == null) return; BtnOCR.IsEnabled = false; PbStatus.Visibility = Visibility.Visible; PbStatus.Maximum = SelectedDocument.Pages.Count; PbStatus.Value = 0; TxtStatus.Text = "OCR 분석 중..."; var targetDoc = SelectedDocument; Task.Run(async () => { try { for (int i = 0; i < targetDoc.Pages.Count; i++) { var pageVM = targetDoc.Pages[i]; if (targetDoc.DocReader == null) continue; using (var r = targetDoc.DocReader.GetPageReader(i)) { var rawBytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight(); using (var stream = new MemoryStream(rawBytes)) { var sb = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied); var ibuffer = CryptographicBuffer.CreateFromByteArray(rawBytes); sb.CopyFromBuffer(ibuffer); var res = await _ocrEngine.RecognizeAsync(sb); var list = new List<OcrWordInfo>(); foreach (var l in res.Lines) foreach (var wd in l.Words) list.Add(new OcrWordInfo { Text = wd.Text, BoundingBox = new Rect(wd.BoundingRect.X, wd.BoundingRect.Y, wd.BoundingRect.Width, wd.BoundingRect.Height) }); pageVM.OcrWords = list; } } Application.Current.Dispatcher.Invoke(() => PbStatus.Value = i + 1); } Application.Current.Dispatcher.Invoke(() => TxtStatus.Text = "OCR 완료"); } catch (Exception ex) { Application.Current.Dispatcher.Invoke(() => MessageBox.Show($"OCR 오류: {ex.Message}")); } finally { Application.Current.Dispatcher.Invoke(() => { BtnOCR.IsEnabled = true; PbStatus.Visibility = Visibility.Collapsed; }); } }); }
        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) _pdfService.SavePdf(SelectedDocument, SelectedDocument.FilePath); }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) { if (SelectedDocument == null) return; var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_copy" }; if (dlg.ShowDialog() == true) _pdfService.SavePdf(SelectedDocument, dlg.FileName); }
        private void CheckTextInSelection(int pageIndex, Rect uiRect) { _selectedTextBuffer = ""; if (SelectedDocument?.DocReader == null) return; var sb = new StringBuilder(); using (var reader = SelectedDocument.DocReader.GetPageReader(pageIndex)) { var chars = reader.GetCharacters().ToList(); foreach (var c in chars) { var r = new Rect(Math.Min(c.Box.Left, c.Box.Right), Math.Min(c.Box.Top, c.Box.Bottom), Math.Abs(c.Box.Right - c.Box.Left), Math.Abs(c.Box.Bottom - c.Box.Top)); if (uiRect.IntersectsWith(r)) sb.Append(c.Char); } } var pageVM = SelectedDocument.Pages[pageIndex]; if (pageVM.OcrWords != null) { foreach (var word in pageVM.OcrWords) if (uiRect.IntersectsWith(word.BoundingBox)) sb.Append(word.Text + " "); } _selectedTextBuffer = sb.ToString(); }
        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e) { Clipboard.SetText(_selectedTextBuffer); SelectionPopup.IsOpen = false; }
        private void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e) { SelectionPopup.IsOpen = false; }
        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);
        private void AddAnnotation(Color color, AnnotationType type) { if (_selectedPageIndex == -1 || SelectedDocument == null) return; var p = SelectedDocument.Pages[_selectedPageIndex]; var ann = new PdfAnnotation { X = p.SelectionX, Y = p.SelectionY, Width = p.SelectionWidth, Height = p.SelectionHeight, Type = type, AnnotationColor = color }; if (type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B)); else { ann.Background = new SolidColorBrush(color); ann.Height = 2; ann.Y = p.SelectionY + p.SelectionHeight - 2; } p.Annotations.Add(ann); SelectionPopup.IsOpen = false; p.IsSelecting = false; }
        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom += 0.1; }
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom -= 0.1; }
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null && SelectedDocument.Pages.Count > 0) { double viewWidth = MainTabControl.ActualWidth - 60; if (viewWidth > 0 && SelectedDocument.Pages[0].Width > 0) SelectedDocument.Zoom = viewWidth / SelectedDocument.Pages[0].Width; } }
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null && SelectedDocument.Pages.Count > 0) { double viewHeight = MainTabControl.ActualHeight - 60; if (viewHeight > 0 && SelectedDocument.Pages[0].Height > 0) SelectedDocument.Zoom = viewHeight / SelectedDocument.Pages[0].Height; } }
        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { if (SelectedDocument != null) { if (e.Delta > 0) SelectedDocument.Zoom += 0.1; else SelectedDocument.Zoom -= 0.1; } e.Handled = true; } }
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(false);
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(true);
        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e) { if (_selectedAnnotation != null && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) { if (p.Annotations.Contains(_selectedAnnotation)) { p.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; CheckToolbarVisibility(); break; } } } else if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) if (p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; } } }
        private void CheckToolbarVisibility() { bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null); if (TextStyleToolbar != null) TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed; }
        private void AnnotationTextBox_Loaded(object sender, RoutedEventArgs e) { if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann && ann.IsSelected) tb.Focus(); }
        private void AnnotationTextBox_GotFocus(object sender, RoutedEventArgs e) { if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann) { _selectedAnnotation = ann; _selectedAnnotation.IsSelected = true; UpdateToolbarFromAnnotation(ann); CheckToolbarVisibility(); } }
        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e) { var thumb = sender as Thumb; if (thumb?.DataContext is PdfAnnotation ann) { ann.Width = Math.Max(50, ann.Width + e.HorizontalChange); ann.Height = Math.Max(30, ann.Height + e.VerticalChange); } }
        private void UpdateToolbarFromAnnotation(PdfAnnotation ann) { _isUpdatingUiFromSelection = true; try { CbFont.SelectedItem = ann.FontFamily; CbSize.SelectedItem = ann.FontSize; BtnBold.IsChecked = ann.IsBold; if (ann.Foreground is SolidColorBrush brush) { foreach (ComboBoxItem item in CbColor.Items) { string cn = item.Tag?.ToString() ?? ""; Color c = Colors.Black; if (cn == "Red") c = Colors.Red; else if (cn == "Blue") c = Colors.Blue; else if (cn == "Green") c = Colors.Green; else if (cn == "Orange") c = Colors.Orange; if (brush.Color == c) { CbColor.SelectedItem = item; break; } } } } finally { _isUpdatingUiFromSelection = false; } }
        private void StyleChanged(object sender, RoutedEventArgs e) { if (!IsLoaded || _isUpdatingUiFromSelection) return; if (CbFont.SelectedItem != null) _defaultFontFamily = CbFont.SelectedItem.ToString() ?? "Malgun Gothic"; if (CbSize.SelectedItem != null) _defaultFontSize = (double)CbSize.SelectedItem; _defaultIsBold = BtnBold.IsChecked == true; if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null) { string cn = item.Tag.ToString() ?? "Black"; if (cn == "Black") _defaultFontColor = Colors.Black; else if (cn == "Red") _defaultFontColor = Colors.Red; else if (cn == "Blue") _defaultFontColor = Colors.Blue; else if (cn == "Green") _defaultFontColor = Colors.Green; else if (cn == "Orange") _defaultFontColor = Colors.Orange; } if (_selectedAnnotation != null) { _selectedAnnotation.FontFamily = _defaultFontFamily; _selectedAnnotation.FontSize = _defaultFontSize; _selectedAnnotation.IsBold = _defaultIsBold; _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor); } }
        private void Tool_Click(object sender, RoutedEventArgs e) { if (RbCursor.IsChecked == true) _currentTool = "CURSOR"; else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT"; else if (RbText.IsChecked == true) _currentTool = "TEXT"; CheckToolbarVisibility(); }
        private void Window_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; } else if (e.Key == Key.Delete) BtnDeleteAnnotation_Click(this, new RoutedEventArgs()); else if (e.Key == Key.Escape && _selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); } }
        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual { if (parent == null) return null; T? child = default(T); int numVisuals = VisualTreeHelper.GetChildrenCount(parent); for (int i = 0; i < numVisuals; i++) { Visual v = (Visual)VisualTreeHelper.GetChild(parent, i); child = v as T; if (child == null) child = GetVisualChild<T>(v); if (child != null) break; } return child; }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
    }
}