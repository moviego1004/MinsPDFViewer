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
using System.Globalization; 
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Globalization;
using Microsoft.Win32;

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

        private string _currentTool = "CURSOR";
        private PdfAnnotation? _selectedAnnotation = null;
        private Point _dragStartPoint;
        private int _activePageIndex = -1;
        private string _selectedTextBuffer = "";
        private int _selectedPageIndex = -1;
        
        // 드래그 관련 변수
        private bool _isDraggingAnnotation = false;
        private Point _annotationDragStartOffset;
        private bool _isUpdatingUiFromSelection = false;
        
        private string _defaultFontFamily = "Malgun Gothic"; 
        private double _defaultFontSize = 14;
        private Color _defaultFontColor = Colors.Red;
        private bool _defaultIsBold = false;

        // [신규] 현재 로그인된 서명 설정 저장
        private SignatureConfig? _currentSignatureConfig;

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

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e) { if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); } }
        
        private void UpdateScrollViewerState(ListView listView, PdfDocumentModel? oldDoc, PdfDocumentModel? newDoc)
        {
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer == null) return;
            scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
            if (oldDoc != null) { oldDoc.SavedVerticalOffset = scrollViewer.VerticalOffset; oldDoc.SavedHorizontalOffset = scrollViewer.HorizontalOffset; }
            if (newDoc != null) { scrollViewer.ScrollToVerticalOffset(newDoc.SavedVerticalOffset); scrollViewer.ScrollToHorizontalOffset(newDoc.SavedHorizontalOffset); }
            if (newDoc != null) scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
        }

        private void PdfListView_Loaded(object sender, RoutedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null) return;
            
            listView.DataContextChanged -= PdfListView_DataContextChanged;
            listView.DataContextChanged += PdfListView_DataContextChanged;
            if (listView.DataContext is PdfDocumentModel doc) UpdateScrollViewerState(listView, null, doc);

            if (listView.ItemContainerStyle == null)
            {
                var style = new Style(typeof(ListViewItem));
                style.Setters.Add(new Setter(UIElement.FocusableProperty, false));
                listView.ItemContainerStyle = style;
            }
        }

        private void PdfListView_Unloaded(object sender, RoutedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null) return;
            listView.DataContextChanged -= PdfListView_DataContextChanged;
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer != null) scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
        }

        private void PdfListView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView != null) UpdateScrollViewerState(listView, e.OldValue as PdfDocumentModel, e.NewValue as PdfDocumentModel);
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer sv && sv.DataContext is PdfDocumentModel doc) { doc.SavedVerticalOffset = sv.VerticalOffset; doc.SavedHorizontalOffset = sv.HorizontalOffset; }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e) { var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" }; if (dlg.ShowDialog() == true) { var docModel = _pdfService.LoadPdf(dlg.FileName); if (docModel != null) { Documents.Add(docModel); SelectedDocument = docModel; _ = _pdfService.RenderPagesAsync(docModel); } } }
        private void BtnCloseTab_Click(object sender, RoutedEventArgs e) { if (sender is Button btn && btn.Tag is PdfDocumentModel doc) { doc.DocReader?.Dispose(); Documents.Remove(doc); if (Documents.Count == 0) SelectedDocument = null; } }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e) => await FindNextSearchResult();
        private async void BtnNextSearch_Click(object sender, RoutedEventArgs e) => await FindNextSearchResult();

        // [수정] PreviewKeyDown으로 변경하여 Ctrl+Enter 확실하게 처리
        private async void TxtSearch_PreviewKeyDown(object sender, KeyEventArgs e) 
        { 
            if (e.Key == Key.Enter) 
            {
                e.Handled = true; // 이벤트 버블링 방지
                
                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                {
                    // 뒤로 찾기
                    BtnPrevSearch_Click(sender, e);
                }
                else
                {
                    // 앞으로 찾기
                    await FindNextSearchResult();
                }
            } 
        }

        private async void BtnPrevSearch_Click(object sender, RoutedEventArgs e) 
        { 
            if (SelectedDocument == null) return;
            string query = TxtSearch.Text;
            if (string.IsNullOrWhiteSpace(query)) return;
            TxtStatus.Text = "이전 찾는 중...";
            
            var foundAnnot = await _searchService.FindPrevAsync(SelectedDocument, query);
            
            if (foundAnnot != null) {
                PdfPageViewModel? targetPage = null;
                foreach(var p in SelectedDocument.Pages) { if(p.Annotations.Contains(foundAnnot)) { targetPage = p; break; } }
                if (targetPage != null) { 
                    var listView = GetVisualChild<ListView>(MainTabControl); 
                    if (listView != null) listView.ScrollIntoView(targetPage); 
                    TxtStatus.Text = $"발견: {targetPage.PageIndex + 1}페이지"; 
                }
            } else {
                var result = MessageBox.Show("문서의 시작입니다. 끝에서부터 다시 찾으시겠습니까?", "검색 완료", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.OK) { 
                    _searchService.ResetSearch(); 
                    await _searchService.FindPrevAsync(SelectedDocument, query); 
                } else { 
                    TxtStatus.Text = "검색 종료"; 
                }
            }
        } 

        private async Task FindNextSearchResult()
        {
            if (SelectedDocument == null) return;
            string query = TxtSearch.Text;
            if (string.IsNullOrWhiteSpace(query)) return;
            TxtStatus.Text = "검색 중...";
            var foundAnnot = await _searchService.FindNextAsync(SelectedDocument, query);
            if (foundAnnot != null) {
                PdfPageViewModel? targetPage = null;
                foreach(var p in SelectedDocument.Pages) { if(p.Annotations.Contains(foundAnnot)) { targetPage = p; break; } }
                if (targetPage != null) { var listView = GetVisualChild<ListView>(MainTabControl); if (listView != null) listView.ScrollIntoView(targetPage); TxtStatus.Text = $"발견: {targetPage.PageIndex + 1}페이지"; }
            } else {
                var result = MessageBox.Show("문서의 끝입니다. 처음부터 다시 찾으시겠습니까?", "검색 완료", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.OK) { _searchService.ResetSearch(); await FindNextSearchResult(); } else { TxtStatus.Text = "검색 종료"; }
            }
        }

        private void Page_MouseDown(object sender, MouseButtonEventArgs e) 
        { 
            if (_selectedAnnotation != null) 
            { 
                _selectedAnnotation.IsSelected = false; 
                _selectedAnnotation = null; 
                CheckToolbarVisibility();
            } 
            
            if (SelectedDocument == null) return; 
            
            var grid = sender as Grid; 
            if (grid == null) return; 
            
            _activePageIndex = (int)grid.Tag; 
            _dragStartPoint = e.GetPosition(grid); 
            
            var pageVM = SelectedDocument.Pages[_activePageIndex]; 

            if (_currentTool == "TEXT") 
            { 
                var newAnnot = new PdfAnnotation { 
                    Type = AnnotationType.FreeText, 
                    X = _dragStartPoint.X, 
                    Y = _dragStartPoint.Y, 
                    Width = 150, Height = 50, 
                    FontSize = _defaultFontSize, 
                    FontFamily = _defaultFontFamily, 
                    Foreground = new SolidColorBrush(_defaultFontColor), 
                    IsBold = _defaultIsBold, 
                    TextContent = "", 
                    IsSelected = true
                }; 
                pageVM.Annotations.Add(newAnnot); 
                _selectedAnnotation = newAnnot; 
                _currentTool = "CURSOR"; 
                RbCursor.IsChecked = true; 
                UpdateToolbarFromAnnotation(_selectedAnnotation); 
                CheckToolbarVisibility(); 
                e.Handled = true; 
                return; 
            } 
            
            if (_currentTool == "CURSOR" || _currentTool == "HIGHLIGHT") 
            { 
                foreach (var p in SelectedDocument.Pages) { 
                    p.IsSelecting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; 
                } 
                SelectionPopup.IsOpen = false; 
                pageVM.IsSelecting = true; 
                pageVM.SelectionX = _dragStartPoint.X; 
                pageVM.SelectionY = _dragStartPoint.Y; 
                pageVM.SelectionWidth = 0; 
                pageVM.SelectionHeight = 0; 
                
                grid.CaptureMouse();
                e.Handled = true; 
            } 
            CheckToolbarVisibility(); 
        }
        
        private void Page_MouseMove(object sender, MouseEventArgs e) 
        { 
            if (_activePageIndex == -1 || SelectedDocument == null) return; 
            
            var grid = sender as Grid; 
            if (grid == null) return; 
            var pageVM = SelectedDocument.Pages[_activePageIndex]; 
            
            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null) 
            { 
                var currentPoint = e.GetPosition(grid); 
                double newX = currentPoint.X - _annotationDragStartOffset.X;
                double newY = currentPoint.Y - _annotationDragStartOffset.Y;

                if (newX < 0) newX = 0;
                if (newY < 0) newY = 0;
                
                _selectedAnnotation.X = newX;
                _selectedAnnotation.Y = newY;
                
                e.Handled = true; 
                return; 
            } 
            
            if ((_currentTool == "CURSOR" || _currentTool == "HIGHLIGHT") && pageVM.IsSelecting) 
            { 
                var pt = e.GetPosition(grid); 
                double x = Math.Min(_dragStartPoint.X, pt.X); 
                double y = Math.Min(_dragStartPoint.Y, pt.Y); 
                double w = Math.Abs(pt.X - _dragStartPoint.X); 
                double h = Math.Abs(pt.Y - _dragStartPoint.Y); 
                pageVM.SelectionX = x; 
                pageVM.SelectionY = y; 
                pageVM.SelectionWidth = w; 
                pageVM.SelectionHeight = h; 
            } 
        }

        private void Page_MouseUp(object sender, MouseButtonEventArgs e) 
        { 
            var grid = sender as Grid; 
            if (grid == null || _activePageIndex == -1 || SelectedDocument == null) return; 
            
            grid.ReleaseMouseCapture(); 
            
            var p = SelectedDocument.Pages[_activePageIndex]; 
            
            if (_currentTool == "HIGHLIGHT" && p.IsSelecting)
            {
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5)
                {
                    _selectedPageIndex = _activePageIndex;
                    AddAnnotation(Colors.Yellow, AnnotationType.Highlight); 
                }
                p.IsSelecting = false;
            }
            else if (p.IsSelecting && _currentTool == "CURSOR") 
            { 
                if (p.SelectionWidth > 5 && p.SelectionHeight > 5) 
                { 
                    var rect = new Rect(p.SelectionX, p.SelectionY, p.SelectionWidth, p.SelectionHeight); 
                    CheckTextInSelection(_activePageIndex, rect); 
                    SelectionPopup.PlacementTarget = grid; 
                    SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(grid).X, e.GetPosition(grid).Y + 10, 0, 0); 
                    SelectionPopup.IsOpen = true; 
                    _selectedPageIndex = _activePageIndex; 
                    TxtStatus.Text = string.IsNullOrEmpty(_selectedTextBuffer) ? "영역 선택됨" : "텍스트 선택됨"; 
                } 
                else 
                { 
                    p.IsSelecting = false; 
                    TxtStatus.Text = "준비"; 
                } 
            } 
            
            _activePageIndex = -1; 
            _isDraggingAnnotation = false; 
            e.Handled = true; 
            CheckToolbarVisibility(); 
        }
        
        private void AnnotationDragHandle_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as FrameworkElement;
            if (border?.DataContext is PdfAnnotation ann)
            {
                if (_selectedAnnotation != null && _selectedAnnotation != ann)
                    _selectedAnnotation.IsSelected = false;

                _selectedAnnotation = ann;
                _selectedAnnotation.IsSelected = true;
                UpdateToolbarFromAnnotation(ann);
                CheckToolbarVisibility();

                _activePageIndex = -1;
                DependencyObject parent = border;
                Grid? parentGrid = null;

                while (parent != null)
                {
                    if (parent is Grid g && g.Tag is int pageIndex)
                    {
                        _activePageIndex = pageIndex;
                        parentGrid = g;
                        break;
                    }
                    parent = VisualTreeHelper.GetParent(parent);
                }

                if (parentGrid != null)
                {
                    _currentTool = "CURSOR";
                    _isDraggingAnnotation = true;
                    
                    var container = GetParentContentPresenter(border);
                    if (container != null)
                        _annotationDragStartOffset = e.GetPosition(container);
                    else
                        _annotationDragStartOffset = e.GetPosition(border);

                    parentGrid.CaptureMouse(); 
                    e.Handled = true; 
                }
            }
        }
        
        private FrameworkElement? GetParentContentPresenter(DependencyObject child)
        {
            DependencyObject parent = child;
            while (parent != null)
            {
                if (parent is ContentPresenter cp) return cp;
                parent = VisualTreeHelper.GetParent(parent);
            }
            return null;
        }

        private void AnnotationTextBox_Loaded(object sender, RoutedEventArgs e) 
        { 
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann) 
            {
                tb.TextChanged -= AnnotationTextBox_TextChanged; 
                tb.TextChanged += AnnotationTextBox_TextChanged;
                if (ann.IsSelected) tb.Focus();
            }
        }

        private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                var formattedText = new FormattedText(
                    tb.Text,
                    CultureInfo.CurrentCulture,
                    FlowDirection.LeftToRight,
                    new Typeface(ann.FontFamily),
                    ann.FontSize,
                    Brushes.Black,
                    VisualTreeHelper.GetDpi(this).PixelsPerDip);

                double minWidth = 100;
                double minHeight = 50; 
                ann.Width = Math.Max(minWidth, formattedText.Width + 25);
                ann.Height = Math.Max(minHeight, formattedText.Height + 20 + 10);
            }
        }

        private void AnnotationTextBox_GotFocus(object sender, RoutedEventArgs e) 
        { 
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann) 
            { 
                if (!_isDraggingAnnotation)
                {
                    if (_selectedAnnotation != null && _selectedAnnotation != ann)
                        _selectedAnnotation.IsSelected = false;

                    _selectedAnnotation = ann; 
                    _selectedAnnotation.IsSelected = true;
                    UpdateToolbarFromAnnotation(ann); 
                    CheckToolbarVisibility(); 
                }
            } 
        }

        private void Annotation_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var element = sender as FrameworkElement;
            if (element?.DataContext is PdfAnnotation ann && ann.Type != AnnotationType.FreeText)
            {
                if (_currentTool != "CURSOR")
                {
                    _currentTool = "CURSOR";
                    RbCursor.IsChecked = true;
                    CheckToolbarVisibility();
                }

                if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = ann;
                _selectedAnnotation.IsSelected = true;
                UpdateToolbarFromAnnotation(ann);
                
                DependencyObject parent = element;
                Grid? parentGrid = null;
                while (parent != null)
                {
                    if (parent is Grid g && g.Tag is int idx) { 
                        _activePageIndex = idx; 
                        parentGrid = g;
                        break; 
                    }
                    parent = VisualTreeHelper.GetParent(parent);
                }
                
                if (parentGrid != null)
                {
                    _annotationDragStartOffset = e.GetPosition(element);
                    _isDraggingAnnotation = true;
                    parentGrid.CaptureMouse();
                    e.Handled = true;
                }
            }
        }
        
        private async void BtnOCR_Click(object sender, RoutedEventArgs e)
        {
            if (_ocrEngine == null) { MessageBox.Show("OCR 미지원 (Windows 10/11 기능 필요)"); return; }
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0) return;

            BtnOCR.IsEnabled = false;
            PbStatus.Visibility = Visibility.Visible;
            PbStatus.Maximum = SelectedDocument.Pages.Count;
            PbStatus.Value = 0;
            TxtStatus.Text = "OCR 분석 중...";

            var targetDoc = SelectedDocument; 

            try
            {
                await Task.Run(async () =>
                {
                    for (int i = 0; i < targetDoc.Pages.Count; i++)
                    {
                        var pageVM = targetDoc.Pages[i];
                        if (targetDoc.DocReader == null) continue;

                        using (var r = targetDoc.DocReader.GetPageReader(i))
                        {
                            var rawBytes = r.GetImage();
                            var w = r.GetPageWidth();
                            var h = r.GetPageHeight();

                            using (var stream = new MemoryStream(rawBytes))
                            {
                                var ibuffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(rawBytes);
                                using (var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied))
                                {
                                    softwareBitmap.CopyFromBuffer(ibuffer);
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
                        }
                        Application.Current.Dispatcher.Invoke(() => { PbStatus.Value = i + 1; });
                    }
                });
                TxtStatus.Text = "OCR 완료. 저장하세요.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"OCR 오류: {ex.Message}");
                TxtStatus.Text = "오류 발생";
            }
            finally
            {
                BtnOCR.IsEnabled = true;
                PbStatus.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) _pdfService.SavePdf(SelectedDocument, SelectedDocument.FilePath); }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e) { if (SelectedDocument == null) return; var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_ocr" }; if (dlg.ShowDialog() == true) _pdfService.SavePdf(SelectedDocument, dlg.FileName); }
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
        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e) { if (_selectedAnnotation != null && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) { if (p.Annotations.Contains(_selectedAnnotation)) { p.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; CheckToolbarVisibility(); break; } } } else if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a && SelectedDocument != null) { foreach(var p in SelectedDocument.Pages) if (p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; } } }
        private void CheckToolbarVisibility() { bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null); if (TextStyleToolbar != null) TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed; }
        
        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e) { var thumb = sender as Thumb; if (thumb?.DataContext is PdfAnnotation ann) { ann.Width = Math.Max(50, ann.Width + e.HorizontalChange); ann.Height = Math.Max(30, ann.Height + e.VerticalChange); } }
        private void UpdateToolbarFromAnnotation(PdfAnnotation ann) { _isUpdatingUiFromSelection = true; try { CbFont.SelectedItem = ann.FontFamily; CbSize.SelectedItem = ann.FontSize; BtnBold.IsChecked = ann.IsBold; if (ann.Foreground is SolidColorBrush brush) { foreach (ComboBoxItem item in CbColor.Items) { string cn = item.Tag?.ToString() ?? ""; Color c = Colors.Black; if (cn == "Red") c = Colors.Red; else if (cn == "Blue") c = Colors.Blue; else if (cn == "Green") c = Colors.Green; else if (cn == "Orange") c = Colors.Orange; if (brush.Color == c) { CbColor.SelectedItem = item; break; } } } } finally { _isUpdatingUiFromSelection = false; } }
        private void StyleChanged(object sender, RoutedEventArgs e) { if (!IsLoaded || _isUpdatingUiFromSelection) return; if (CbFont.SelectedItem != null) _defaultFontFamily = CbFont.SelectedItem.ToString() ?? "Malgun Gothic"; if (CbSize.SelectedItem != null) _defaultFontSize = (double)CbSize.SelectedItem; _defaultIsBold = BtnBold.IsChecked == true; if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null) { string cn = item.Tag.ToString() ?? "Black"; if (cn == "Black") _defaultFontColor = Colors.Black; else if (cn == "Red") _defaultFontColor = Colors.Red; else if (cn == "Blue") _defaultFontColor = Colors.Blue; else if (cn == "Green") _defaultFontColor = Colors.Green; else if (cn == "Orange") _defaultFontColor = Colors.Orange; } if (_selectedAnnotation != null) { _selectedAnnotation.FontFamily = _defaultFontFamily; _selectedAnnotation.FontSize = _defaultFontSize; _selectedAnnotation.IsBold = _defaultIsBold; _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor); } }
        private void Tool_Click(object sender, RoutedEventArgs e) { if (RbCursor.IsChecked == true) _currentTool = "CURSOR"; else if (RbHighlight.IsChecked == true) _currentTool = "HIGHLIGHT"; else if (RbText.IsChecked == true) _currentTool = "TEXT"; CheckToolbarVisibility(); }
        private void Window_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; } else if (e.Key == Key.Delete) BtnDeleteAnnotation_Click(this, new RoutedEventArgs()); else if (e.Key == Key.Escape && _selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); } }
        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual { if (parent == null) return null; T? child = default(T); int numVisuals = VisualTreeHelper.GetChildrenCount(parent); for (int i = 0; i < numVisuals; i++) { Visual v = (Visual)VisualTreeHelper.GetChild(parent, i); child = v as T; if (child == null) child = GetVisualChild<T>(v); if (child != null) break; } return child; }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));
   

        // [신규] 서명 설정 버튼 클릭 핸들러
        private void BtnSignSettings_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CertificateWindow();
            dlg.Owner = this;
            
            // 이미 로그인된 상태라면 안내? (일단은 그냥 다시 띄워서 변경 가능하게)
            
            if (dlg.ShowDialog() == true)
            {
                _currentSignatureConfig = dlg.ResultConfig;
                // 상태바나 버튼 텍스트 등을 업데이트해서 로그인 상태임을 알려주면 좋음
                TxtStatus.Text = $"서명 준비 완료: {_currentSignatureConfig?.Certificate.Subject}";
                BtnSignSettings.Content = "🔐 서명 설정 (완료)";
                BtnSignSettings.Foreground = Brushes.Green;
            }
        }
    }

}