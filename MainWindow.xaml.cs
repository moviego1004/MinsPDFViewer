using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using Microsoft.Win32;
using PdfiumViewer;

namespace MinsPDFViewer
{
    public partial class MainWindow : Window, System.ComponentModel.INotifyPropertyChanged
    {
        private readonly PdfService _pdfService;
        private readonly SearchService _searchService;
        private readonly PdfSignatureService _signatureService;
        private readonly SignatureVerificationService _signatureVerificationService;
        private readonly OcrService _ocrService;
        private readonly HistoryService _historyService;
        private const int ThumbnailRenderPagesBefore = 3;
        private const int ThumbnailRenderPagesAfter = 8;
        private const double MinThumbnailImageWidth = 80;
        private const double MaxThumbnailImageWidth = 320;
        private CancellationTokenSource? _thumbnailRenderCts;
        private bool _syncingThumbnailSelection;
        private Point _thumbnailDragStartPoint;
        private PdfPageViewModel? _draggedThumbnailPage;
        private List<PdfPageViewModel> _draggedThumbnailPages = new();
        private bool _isThumbnailDragging;
        private double _thumbnailImageWidth = 140;

        public System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel> Documents
        {
            get; set;
        }
            = new System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel>();

        public double ThumbnailImageWidth
        {
            get => _thumbnailImageWidth;
            private set
            {
                double normalized = Math.Max(MinThumbnailImageWidth, Math.Min(MaxThumbnailImageWidth, value));
                if (Math.Abs(_thumbnailImageWidth - normalized) < 0.1)
                    return;

                _thumbnailImageWidth = normalized;
                OnPropertyChanged(nameof(ThumbnailImageWidth));
                OnPropertyChanged(nameof(ThumbnailFrameWidth));
            }
        }

        public double ThumbnailFrameWidth => ThumbnailImageWidth + 10;

        private PdfDocumentModel? _selectedDocument;
        public PdfDocumentModel? SelectedDocument
        {
            get => _selectedDocument;
            set
            {
                _selectedDocument = value;
                OnPropertyChanged(nameof(SelectedDocument));
                CheckToolbarVisibility();
            }
        }

        private string _currentTool = "CURSOR";
        private bool _isEditMode = false;
        private PdfAnnotation? _selectedAnnotation = null;
        private Point _dragStartPoint;
        private int _activePageIndex = -1;
        private string _selectedTextBuffer = "";
        private bool _isSelectionTextPending = false;
        private int _selectionTextRequestId = 0;
        private int _selectedPageIndex = -1;

        private bool _isDraggingAnnotation = false;
        private bool _isPendingAnnotationDrag = false;
        private Point _annotationDragStartOffset;
        private Point _annotationMouseDownPoint;
        private Grid? _dragGrid = null;
        private FrameworkElement? _dragElement = null;
        private Point _lastAnnotationDragPoint;
        private bool _isUpdatingUiFromSelection = false;

        private string _defaultFontFamily = "Malgun Gothic";
        private double _defaultFontSize = 14;
        private Color _defaultFontColor = Colors.Red;
        private bool _defaultIsBold = false;
        private const int MaxImageStampPixelSide = 6000;

        private bool _isSpacePressed = false;
        private bool _isPanning = false;
        private Point _lastPanPoint;

        private ScrollViewer? _cachedScrollViewer;
        private GridLength _lastSidebarWidth = new GridLength(250);

        public bool IsEditMode
        {
            get => _isEditMode;
            private set
            {
                if (_isEditMode == value) return;
                _isEditMode = value;
                OnPropertyChanged(nameof(IsEditMode));
                ApplyInteractionModeToVisibleCanvases();
                if (!_isEditMode)
                    ClearSelectedAnnotation();
                CheckToolbarVisibility();
                TxtStatus.Text = _isEditMode ? "편집 모드" : "읽기 모드";
            }
        }

        public MainWindow()
        {
            // Register encoding provider for proper PDF text handling
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            InitializeComponent();
            DataContext = this;
            _pdfService = new PdfService();
            _searchService = new SearchService(_pdfService);
            _signatureService = new PdfSignatureService();
            _signatureVerificationService = new SignatureVerificationService();
            _ocrService = new OcrService();
            _historyService = new HistoryService();

            // Setup global logging and exception handling
            Log("=== Application Starting ===");
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Log($"[CRITICAL] UnhandledException: {e.ExceptionObject}");
            this.Dispatcher.UnhandledException += (s, e) => {
                Log($"[CRITICAL] DispatcherUnhandledException: {e.Exception}");
            };

            CbFont.ItemsSource = new string[] { "Malgun Gothic", "Gulim", "Dotum", "Batang" };
            CbFont.SelectedIndex = 0;
            CbSize.ItemsSource = new double[] { 10, 12, 14, 16, 18, 24, 32, 48 };
            CbSize.SelectedIndex = 2;
        }

        private void Log(string message)
        {
            try
            {
                DebugLog.Append(string.Empty, message);
            }
            catch { }
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.OriginalSource, MainTabControl))
                return;

            if (_selectedAnnotation != null)
            {
                _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = null;
                CheckToolbarVisibility();
            }

            if (SelectedDocument != null)
                StartThumbnailRendering(SelectedDocument);
        }

        private void UpdateScrollViewerState(ListView listView, PdfDocumentModel? oldDoc, PdfDocumentModel? newDoc)
        {
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer == null)
                return;

            scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;

            if (oldDoc != null)
            {
                oldDoc.SavedVerticalOffset = scrollViewer.VerticalOffset;
                oldDoc.SavedHorizontalOffset = scrollViewer.HorizontalOffset;
            }

            if (newDoc != null)
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    scrollViewer.ScrollToVerticalOffset(newDoc.SavedVerticalOffset);
                    scrollViewer.ScrollToHorizontalOffset(newDoc.SavedHorizontalOffset);
                    scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else
            {
                scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
            }
        }

        private void PdfListView_Loaded(object sender, RoutedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null)
                return;
            listView.DataContextChanged -= PdfListView_DataContextChanged;
            listView.DataContextChanged += PdfListView_DataContextChanged;
            if (listView.DataContext is PdfDocumentModel doc)
                UpdateScrollViewerState(listView, null, doc);
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
            if (listView == null)
                return;
            listView.DataContextChanged -= PdfListView_DataContextChanged;
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer != null)
                scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;
        }

        private void PdfListView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            var listView = sender as ListView;
            if (listView != null)
                UpdateScrollViewerState(listView, e.OldValue as PdfDocumentModel, e.NewValue as PdfDocumentModel);
        }

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (sender is ScrollViewer sv && sv.DataContext is PdfDocumentModel doc)
            {
                doc.SavedVerticalOffset = sv.VerticalOffset;
                doc.SavedHorizontalOffset = sv.HorizontalOffset;
            }
        }

        private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PdfDocumentModel doc)
            {
                if (doc == SelectedDocument)
                {
                    int currentPage = GetCurrentPageIndex();
                    _historyService.SetLastPage(doc.FilePath, currentPage);
                    _historyService.SaveHistory();
                }
                doc.Dispose();
                Documents.Remove(doc);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (SelectedDocument != null)
            {
                int currentPage = GetCurrentPageIndex();
                _historyService.SetLastPage(SelectedDocument.FilePath, currentPage);
                SelectedDocument.IsDisposed = true;
            }
            _historyService.SaveHistory();
            foreach (var doc in Documents)
            {
                doc.Dispose();
            }
        }

        private void AnnotationCanvas_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is Canvas canvas && canvas.DataContext is PdfPageViewModel pageVM)
            {
                if (canvas.Tag is Tuple<PdfPageViewModel, NotifyCollectionChangedEventHandler> previous)
                    previous.Item1.Annotations.CollectionChanged -= previous.Item2;

                NotifyCollectionChangedEventHandler handler = (s, args) => RefreshCanvas(canvas, pageVM);
                canvas.Tag = Tuple.Create(pageVM, handler);
                pageVM.Annotations.CollectionChanged += handler;
                canvas.IsHitTestVisible = IsEditMode;
                
                RefreshCanvas(canvas, pageVM);
            }
        }

        private void RefreshCanvas(Canvas canvas, PdfPageViewModel pageVM)
        {
            canvas.Children.Clear();
            foreach (var ann in pageVM.Annotations)
            {
                var ui = CreateAnnotationUI(ann);
                canvas.Children.Add(ui);
            }
        }

        private FrameworkElement CreateAnnotationUI(PdfAnnotation ann)
        {
            // [Fix] Must set BOTH DataContext (for outer bindings like Canvas.Left) AND Content (for inner DataTemplate)
            var cc = new ContentControl { DataContext = ann, Content = ann, Focusable = false };
            
            // Set Resource Key based on type
            string templateKey = ann.Type switch
            {
                AnnotationType.FreeText => "FreeTextTemplate",
                AnnotationType.Highlight => "HighlightTemplate",
                AnnotationType.Underline => "UnderlineTemplate",
                AnnotationType.SearchHighlight => "SearchHighlightTemplate",
                AnnotationType.SignaturePlaceholder => "SignaturePlaceholderTemplate",
                AnnotationType.SignatureField => "SignatureFieldTemplate",
                AnnotationType.ImageStamp => "ImageStampTemplate",
                _ => null
            };

            if (templateKey != null && Resources.Contains(templateKey))
            {
                cc.ContentTemplate = (DataTemplate)Resources[templateKey];
            }

            // Bind position and size
            cc.SetBinding(Canvas.LeftProperty, new Binding("X") { Mode = BindingMode.TwoWay });
            cc.SetBinding(Canvas.TopProperty, new Binding("Y") { Mode = BindingMode.TwoWay });
            cc.SetBinding(WidthProperty, new Binding("Width") { Mode = BindingMode.TwoWay });
            cc.SetBinding(HeightProperty, new Binding("Height") { Mode = BindingMode.TwoWay });
            
            Panel.SetZIndex(cc, 100);
            return cc;
        }

        private void PageGrid_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement elem && elem.DataContext is PdfPageViewModel pageVM)
            {
                if (SelectedDocument != null)
                {
                    // Cancel any previous render for this page
                    pageVM.CancelRender();
                    var cts = new System.Threading.CancellationTokenSource();
                    pageVM.RenderCts = cts;
                    var doc = SelectedDocument;
                    _pdfService.RenderPageAsync(doc, pageVM, cts.Token);
                }
            }
        }

        private void PageGrid_Unloaded(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement elem && elem.DataContext is PdfPageViewModel pageVM)
            {
                pageVM.Unload();
            }
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e) => await FindNextSearchResult();
        private async void BtnNextSearch_Click(object sender, RoutedEventArgs e) => await FindNextSearchResult();
        private async void TxtSearch_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
                    BtnPrevSearch_Click(sender, e);
                else
                    await FindNextSearchResult();
            }
        }

        private async void BtnPrevSearch_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            string query = TxtSearch.Text;
            if (string.IsNullOrWhiteSpace(query))
                return;
            TxtStatus.Text = "이전 찾는 중...";
            var foundAnnot = await _searchService.FindPrevAsync(SelectedDocument, query);
            HandleSearchResult(foundAnnot, query, true);
        }

        private async Task FindNextSearchResult()
        {
            if (SelectedDocument == null)
                return;
            string query = TxtSearch.Text;
            if (string.IsNullOrWhiteSpace(query))
                return;
            TxtStatus.Text = "검색 중...";
            var foundAnnot = await _searchService.FindNextAsync(SelectedDocument, query);
            HandleSearchResult(foundAnnot, query, false);
        }

        private async void HandleSearchResult(PdfAnnotation? foundAnnot, string query, bool isPrev)
        {
            if (foundAnnot != null && SelectedDocument != null)
            {
                PdfPageViewModel? targetPage = null;
                foreach (var p in SelectedDocument.Pages)
                {
                    if (p.Annotations.Contains(foundAnnot))
                    {
                        targetPage = p;
                        break;
                    }
                }
                if (targetPage != null)
                {
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null)
                        listView.ScrollIntoView(targetPage);
                    TxtStatus.Text = $"발견: {targetPage.PageIndex + 1}페이지";
                }
            }
            else
            {
                string msg = isPrev ? "문서의 시작입니다. 끝에서부터 다시 찾으시겠습니까?" : "문서의 끝입니다. 처음부터 다시 찾으시겠습니까?";
                var result = MessageBox.Show(msg, "검색 완료", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.OK)
                {
                    _searchService.ResetSearch();
                    if (isPrev)
                        await _searchService.FindPrevAsync(SelectedDocument!, query);
                    else
                        await _searchService.FindNextAsync(SelectedDocument!, query);
                }
                else
                {
                    TxtStatus.Text = "검색 종료";
                }
            }
        }

        public void Page_MouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                // 1. Check if we clicked on an annotation (including TextBox)
                if (IsAnnotationObject(e.OriginalSource))
                {
                    return;
                }

                // 2. Only if NOT clicking an annotation, clear focus from TextBox
                if (Keyboard.FocusedElement is TextBox)
                {
                    Keyboard.ClearFocus();
                }

                if (SelectedDocument == null) return;
                
                // Since we removed EventSetter from ListViewItem, sender should be the Grid from DataTemplate
                var grid = sender as Grid;
                if (grid == null) return;
                
                var pageVM = grid.DataContext as PdfPageViewModel;
                if (pageVM == null) return;
                
                _activePageIndex = pageVM.PageIndex;
                _dragStartPoint = e.GetPosition(grid);

                // [Fix] Check if clicked area is a Signature Field (Fallback for missing Widget annotation)
                if (IsEditMode && _currentTool == "CURSOR")
                {
                    if (CheckSignatureFieldClick(pageVM, _dragStartPoint))
                    {
                        e.Handled = true;
                        return;
                    }
                }

                // [Text Tool Logic]
                if (_currentTool == "TEXT")
                {
                    if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                    var newAnnot = new PdfAnnotation { Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50, FontSize = _defaultFontSize, FontFamily = _defaultFontFamily, Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold, TextContent = "", IsSelected = true };
                    pageVM.Annotations.Add(newAnnot);
                    var canvas = FindChild<Canvas>(grid, "AnnotationCanvas");
                    if (canvas != null) RefreshCanvas(canvas, pageVM);
                    _selectedAnnotation = newAnnot; _currentTool = "CURSOR"; RbCursor.IsChecked = true;
                    UpdateToolbarFromAnnotation(_selectedAnnotation); CheckToolbarVisibility(); 
                    e.Handled = true; 
                    return;
                }

                // [Highlight Tool Logic]
                if (_currentTool == "HIGHLIGHT")
                {
                    if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                    
                    // Clear selection on all pages
                    foreach (var p in SelectedDocument.Pages) 
                    { 
                        p.IsSelecting = false; 
                        p.IsHighlighting = false; 
                        p.SelectionWidth = 0; 
                        p.SelectionHeight = 0; 
                    }
                    
                    SelectionPopup.IsOpen = false; 
                    
                    // Setup new selection
                    pageVM.IsSelecting = true; 
                    pageVM.IsHighlighting = true; 
                    pageVM.SelectionX = _dragStartPoint.X; 
                    pageVM.SelectionY = _dragStartPoint.Y; 
                    pageVM.SelectionWidth = 0;
                    pageVM.SelectionHeight = 0;
                    
                    // Capture mouse to ensure we get MouseMove/MouseUp events
                    if (grid.CaptureMouse())
                    {
                        e.Handled = true;
                    }
                    return;
                }
                
                if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); }
                
                // [Cursor Tool Logic - Selection]
                if (_currentTool == "CURSOR")
                {
                    if (IsEditMode)
                    {
                        SelectionPopup.IsOpen = false;
                        e.Handled = true;
                        return;
                    }

                    foreach (var p in SelectedDocument.Pages) { p.IsSelecting = false; p.IsHighlighting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                    SelectionPopup.IsOpen = false; 
                    pageVM.IsSelecting = true; 
                    pageVM.IsHighlighting = false; 
                    pageVM.SelectionX = _dragStartPoint.X; 
                    pageVM.SelectionY = _dragStartPoint.Y; 
                    pageVM.SelectionWidth = 0;
                    pageVM.SelectionHeight = 0;

                    if (grid.CaptureMouse())
                    {
                        e.Handled = true;
                    }
                }
                CheckToolbarVisibility();
            }
            catch (Exception) { }
        }

        private bool IsAnnotationObject(object source)
        {
            if (!(source is DependencyObject current)) return false;

            if (source is FrameworkContentElement fce)
                current = fce.Parent;

            while (current != null)
            {
                if (current is FrameworkElement fe && fe.DataContext is PdfAnnotation)
                    return true;
                if (current is Grid g && g.DataContext is PdfPageViewModel)
                    break;

                DependencyObject? parent = null;
                try
                {
                    parent = VisualTreeHelper.GetParent(current);
                }
                catch { }

                if (parent == null && current is FrameworkElement fe2)
                    parent = fe2.Parent;

                current = parent;
            }
            return false;
        }

        private bool CheckSignatureFieldClick(PdfPageViewModel pageVM, Point clickPoint)
        {
            if (SelectedDocument == null || !File.Exists(SelectedDocument.FilePath)) return false;

            try
            {
                // UI Coordinate -> PDF Coordinate Logic
                // clickPoint is relative to Grid (Page View).
                // PDF Origin is Bottom-Left. UI is Top-Left.
                
                // Calculate Scale
                double scaleX = pageVM.Width / pageVM.PdfPageWidthPoint;
                double scaleY = pageVM.Height / pageVM.PdfPageHeightPoint;
                
                if (scaleX <= 0 || scaleY <= 0) return false;

                double pdfX = clickPoint.X / scaleX;
                double pdfY = pageVM.PdfPageHeightPoint - (clickPoint.Y / scaleY);

                var result = _signatureVerificationService.VerifySignatureAtPoint(
                    SelectedDocument.FilePath,
                    pageVM.OriginalPageIndex,
                    pdfX,
                    pdfY);

                if (result != null)
                {
                    new SignatureResultWindow(result) { Owner = this }.ShowDialog();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log($"[CheckSignatureFieldClick] Error: {ex}");
            }
            return false;
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1 || SelectedDocument == null)
                return;
            
            var pageVM = (sender as FrameworkElement)?.DataContext as PdfPageViewModel;
            if (pageVM == null) return;

            var relativeTo = sender as IInputElement;

            if (_currentTool == "CURSOR" &&
                _isPendingAnnotationDrag &&
                !_isDraggingAnnotation &&
                _selectedAnnotation != null &&
                _dragGrid != null &&
                e.LeftButton == MouseButtonState.Pressed)
            {
                var currentPoint = e.GetPosition(_dragGrid);
                if (Math.Abs(currentPoint.X - _annotationMouseDownPoint.X) >= SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(currentPoint.Y - _annotationMouseDownPoint.Y) >= SystemParameters.MinimumVerticalDragDistance)
                {
                    _isDraggingAnnotation = true;
                    _isPendingAnnotationDrag = false;
                    _dragGrid.CaptureMouse();
                }
                else
                {
                    return;
                }
            }

            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null)
            {
                var dragRelativeTo = (_dragGrid as IInputElement) ?? relativeTo;
                var currentPoint = e.GetPosition(dragRelativeTo);
                double newX = currentPoint.X - _annotationDragStartOffset.X;
                double newY = currentPoint.Y - _annotationDragStartOffset.Y;
                if (newX < 0) newX = 0;
                if (newY < 0) newY = 0;
                if (Math.Abs(newX - _lastAnnotationDragPoint.X) < 0.5 &&
                    Math.Abs(newY - _lastAnnotationDragPoint.Y) < 0.5)
                {
                    e.Handled = true;
                    return;
                }
                _lastAnnotationDragPoint = new Point(newX, newY);
                _selectedAnnotation.X = newX;
                _selectedAnnotation.Y = newY;
                e.Handled = true;
                return;
            }
            if ((_currentTool == "CURSOR" || _currentTool == "HIGHLIGHT") && pageVM.IsSelecting)
            {
                var pt = e.GetPosition(relativeTo);
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

        public void Page_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // Force release mouse capture from any element (Grid, ListViewItem, etc.)
            Mouse.Capture(null);
            
            if (_activePageIndex == -1 || SelectedDocument == null) return;
            
            var pageVM = SelectedDocument.Pages[_activePageIndex];
            
            if (_currentTool == "HIGHLIGHT" && pageVM.IsSelecting)
            {
                if (pageVM.SelectionWidth > 5 && pageVM.SelectionHeight > 5)
                {
                    _selectedPageIndex = _activePageIndex;
                    AddAnnotation(Colors.Yellow, AnnotationType.Highlight);
                    var grid = sender as Grid ?? FindAncestor<Grid>(sender as DependencyObject);
                    if (grid != null)
                    {
                        var canvas = FindChild<Canvas>(grid, "AnnotationCanvas");
                        if (canvas != null) RefreshCanvas(canvas, pageVM);
                    }
                }
                pageVM.IsSelecting = false;
                pageVM.IsHighlighting = false;
            }
            else if (pageVM.IsSelecting && _currentTool == "CURSOR")
            {
                if (pageVM.SelectionWidth > 5 && pageVM.SelectionHeight > 5)
                {
                    var rect = new Rect(pageVM.SelectionX, pageVM.SelectionY, pageVM.SelectionWidth, pageVM.SelectionHeight);
                    SelectionPopup.PlacementTarget = sender as UIElement;
                    SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(sender as IInputElement).X, e.GetPosition(sender as IInputElement).Y + 10, 0, 0);
                    SelectionPopup.IsOpen = true;
                    _selectedPageIndex = _activePageIndex;
                    CheckTextInSelection(_activePageIndex, rect);
                }
                else
                {
                    pageVM.IsSelecting = false;
                    TxtStatus.Text = "준비";
                }
            }
            _activePageIndex = -1;
            _isDraggingAnnotation = false;
            _isPendingAnnotationDrag = false;
            _dragGrid = null;
            _dragElement = null;
            e.Handled = true;
            CheckToolbarVisibility();
        }

        private async void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            int requestId = ++_selectionTextRequestId;
            _selectedTextBuffer = "";
            _isSelectionTextPending = true;
            BtnPopupCopy.IsEnabled = false;
            TxtStatus.Text = "텍스트 추출 중...";

            if (SelectedDocument?.PdfDocument == null)
            {
                _isSelectionTextPending = false;
                TxtStatus.Text = "영역 선택됨";
                return;
            }

            var doc = SelectedDocument;

            string extractedText = await Task.Run(() =>
            {
                try
                {
                    if (pageIndex < 0 || pageIndex >= doc.Pages.Count)
                        return string.Empty;

                    var page = doc.Pages[pageIndex];
                    return _pdfService.ExtractTextInRect(
                        doc.FilePath,
                        page.OriginalPageIndex,
                        uiRect,
                        page.Width,
                        page.Height,
                        page.PdfPageWidthPoint,
                        page.PdfPageHeightPoint);
                }
                catch { return ""; }
            });

            if (requestId != _selectionTextRequestId)
                return;

            _selectedTextBuffer = extractedText;
            _isSelectionTextPending = false;
            BtnPopupCopy.IsEnabled = !string.IsNullOrWhiteSpace(_selectedTextBuffer);
            if (!string.IsNullOrEmpty(_selectedTextBuffer))
            {
                TxtStatus.Text = "텍스트 선택됨";
            }
            else
            {
                TxtStatus.Text = "영역 선택됨";
            }
        }

        private async void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SelectionPopup.IsOpen = false;
                if (SelectedDocument == null || _selectedPageIndex == -1) return;

                var pageVM = SelectedDocument.Pages[_selectedPageIndex];
                if (pageVM.SelectionWidth <= 0 || pageVM.SelectionHeight <= 0) return;

                TxtStatus.Text = "이미지 처리 중...";
                BitmapSource? croppedBitmap = null;

                await Task.Run(() =>
                {
                    lock (PdfService.PdfiumLock)
                    {
                        if (SelectedDocument.PdfDocument == null) return;

                        int scale = 3;
                        int dpi = 96 * scale;
                        int w = (int)(pageVM.Width * scale);
                        int h = (int)(pageVM.Height * scale);

                        using (var fullBmp = (System.Drawing.Bitmap)SelectedDocument.PdfDocument.Render(_selectedPageIndex, w, h, dpi, dpi, PdfRenderFlags.Annotations))
                        {
                            int cx = (int)(pageVM.SelectionX * scale);
                            int cy = (int)(pageVM.SelectionY * scale);
                            int cw = (int)(pageVM.SelectionWidth * scale);
                            int ch = (int)(pageVM.SelectionHeight * scale);

                            if (cw > 0 && ch > 0)
                            {
                                if (cx < 0) cx = 0;
                                if (cy < 0) cy = 0;
                                if (cx + cw > fullBmp.Width) cw = fullBmp.Width - cx;
                                if (cy + ch > fullBmp.Height) ch = fullBmp.Height - cy;

                                using (var cropped = fullBmp.Clone(new System.Drawing.Rectangle(cx, cy, cw, ch), fullBmp.PixelFormat))
                                {
                                    var bs = PdfService.ToBitmapSource(cropped);
                                    bs.Freeze();
                                    croppedBitmap = bs;
                                }
                            }
                        }
                    }
                });

                if (croppedBitmap != null)
                {
                    Clipboard.SetImage(croppedBitmap);
                    TxtStatus.Text = "이미지가 클립보드에 복사되었습니다.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이미지 복사 실패: {ex.Message}");
            }
        }

        private void AnnotationDragHandle_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var border = sender as FrameworkElement;
            if (border?.DataContext is PdfAnnotation ann)
            {
                if (e.ClickCount == 2 && ann.Type == AnnotationType.SignaturePlaceholder)
                {
                    PerformSignatureProcess(ann);
                    e.Handled = true;
                    return;
                }

                SelectAnnotation(ann);

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
                    _isPendingAnnotationDrag = true;
                    _isDraggingAnnotation = false;
                    _dragGrid = parentGrid;
                    _dragElement = border;
                    var container = GetParentContentPresenter(border);
                    _annotationDragStartOffset = container != null ? e.GetPosition(container) : e.GetPosition(border);
                    _annotationMouseDownPoint = e.GetPosition(parentGrid);
                    _lastAnnotationDragPoint = new Point(ann.X, ann.Y);
                    e.Handled = true;
                }
            }
        }

        private async void PerformSignatureProcess(PdfAnnotation ann)
        {
            if (SelectedDocument == null) return;

            int pageIndex = -1;
            PdfPageViewModel? targetPage = null;
            for (int i = 0; i < SelectedDocument.Pages.Count; i++)
            {
                if (SelectedDocument.Pages[i].Annotations.Contains(ann))
                {
                    pageIndex = i;
                    targetPage = SelectedDocument.Pages[i];
                    break;
                }
            }

            if (pageIndex == -1 || targetPage == null) return;

            var dlg = new CertificateWindow();
            dlg.Owner = this;

            if (dlg.ShowDialog() == true && dlg.ResultConfig != null)
            {
                var config = dlg.ResultConfig;
                config.VisualStampPath = ann.VisualStampPath;
                config.UseVisualStamp = true;

                var saveDlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_signed.pdf" };

                if (saveDlg.ShowDialog() == true)
                {
                    try
                    {
                        double effectivePdfWidth = (targetPage.CropWidthPoint > 0) ? targetPage.CropWidthPoint : targetPage.PdfPageWidthPoint;
                        double effectivePdfHeight = (targetPage.CropHeightPoint > 0) ? targetPage.CropHeightPoint : targetPage.PdfPageHeightPoint;
                        double scaleX = effectivePdfWidth / targetPage.Width;
                        double scaleY = effectivePdfHeight / targetPage.Height;
                        double pdfX = targetPage.CropX + (ann.X * scaleX);
                        double pdfY = (targetPage.CropY + effectivePdfHeight) - ((ann.Y + ann.Height) * scaleY);

                        var pdfRect = new SignaturePdfRect(pdfX, pdfY, ann.Width * scaleX, ann.Height * scaleY);
                        string tempPath = Path.GetTempFileName();
                        string signedPath = saveDlg.FileName;

                        try
                        {
                            await _pdfService.SavePdf(SelectedDocument, tempPath);
                            _signatureService.SignPdf(tempPath, signedPath, config, pageIndex, pdfRect);
                        }
                        finally
                        {
                            if (File.Exists(tempPath)) File.Delete(tempPath);
                        }

                        targetPage.Annotations.Remove(ann);
                        _historyService.SetLastPage(signedPath, pageIndex);
                        _historyService.SaveHistory();

                        await ReloadPdfFromPathAsync(signedPath, pageIndex);

                        TxtStatus.Text = "전자서명 완료";
                        MessageBox.Show($"전자서명 완료!\n저장됨: {signedPath}");
                    }
                    catch (Exception ex) { MessageBox.Show($"서명 실패: {ex.Message}"); }
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
            var tb = sender as TextBox;
            if (tb == null) return;

            object dc = tb.DataContext;
            // Keep this quiet; virtualized text annotations can load repeatedly while scrolling.

            if (dc is PdfAnnotation ann)
            {
                // [Robustness] Force subscribe to ensure event fires
                tb.TextChanged -= AnnotationTextBox_TextChanged;
                tb.TextChanged += AnnotationTextBox_TextChanged;
                
                // [Sync] Ensure initial text is correct
                if (ann.TextContent != tb.Text) tb.Text = ann.TextContent;

                if (ann.IsSelected) tb.Focus();
                // Avoid logging full annotation text; pasted/OCR text can be very large.
            }
            else
            {
                Log($"[CRITICAL] AnnotationTextBox_Loaded: DataContext is NOT PdfAnnotation! It is {dc}");
            }
        }

        private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {            
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                // [Direct Sync] Always push text to model, bypassing binding if necessary
                ann.TextContent = tb.Text;
         
                if (!tb.IsKeyboardFocusWithin)
                    return;

                var formattedText = new FormattedText(tb.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface(new System.Windows.Media.FontFamily(ann.FontFamily), FontStyles.Normal, ann.IsBold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal),
                    ann.FontSize, Brushes.Black, VisualTreeHelper.GetDpi(this).PixelsPerDip);

                double newWidth = Math.Min(900, Math.Max(100, formattedText.Width + 25));
                double newHeight = Math.Min(600, Math.Max(50, formattedText.Height + 30));

                if (Math.Abs(ann.Width - newWidth) > 2) ann.Width = newWidth;
                if (Math.Abs(ann.Height - newHeight) > 2) ann.Height = newHeight;
            }
        }

        private void AnnotationTextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                if (!_isDraggingAnnotation)
                {
                    SelectAnnotation(ann);
                }
            }
        }

        private void AnnotationTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                ann.TextContent = tb.Text;
            }
        }

        private void AnnotationTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Keyboard.ClearFocus();
                if (_selectedAnnotation != null)
                {
                    _selectedAnnotation.IsSelected = false;
                    _selectedAnnotation = null;
                    CheckToolbarVisibility();
                }
                e.Handled = true;
            }
        }

        private T? FindAncestorDataContext<T>(DependencyObject child) where T : class
        {
            DependencyObject? current = child;
            while (current != null)
            {
                if (current is FrameworkElement fe && fe.DataContext is T t) return t;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private void SelectAnnotation(PdfAnnotation ann)
        {
            if (!IsEditMode)
                return;

            bool changed = _selectedAnnotation != ann;
            if (changed && _selectedAnnotation != null)
                _selectedAnnotation.IsSelected = false;

            _selectedAnnotation = ann;
            if (!ann.IsSelected)
            {
                ann.IsSelected = true;
                changed = true;
            }

            if (changed)
            {
                UpdateToolbarFromAnnotation(ann);
                CheckToolbarVisibility();
            }
        }

        private void ClearSelectedAnnotation()
        {
            if (_selectedAnnotation != null)
                _selectedAnnotation.IsSelected = false;
            _selectedAnnotation = null;
            _isDraggingAnnotation = false;
            _isPendingAnnotationDrag = false;
            _dragGrid = null;
            _dragElement = null;
        }

        private void Annotation_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var element = sender as FrameworkElement;
            if (element?.DataContext is PdfAnnotation ann && ann.Type != AnnotationType.FreeText && ann.Type != AnnotationType.SignaturePlaceholder)
            {
                if (_currentTool != "CURSOR")
                {
                    _currentTool = "CURSOR";
                    RbCursor.IsChecked = true;
                    CheckToolbarVisibility();
                }
                SelectAnnotation(ann);
                
                DependencyObject parent = element;
                Grid? parentGrid = null;
                while (parent != null)
                {
                    if (parent is Grid g && g.Tag is int idx) { _activePageIndex = idx; parentGrid = g; break; }
                    parent = VisualTreeHelper.GetParent(parent);
                }
                if (parentGrid != null)
                {
                    _annotationDragStartOffset = e.GetPosition(element);
                    _annotationMouseDownPoint = e.GetPosition(parentGrid);
                    _lastAnnotationDragPoint = new Point(ann.X, ann.Y);
                    _isPendingAnnotationDrag = true;
                    _isDraggingAnnotation = false;
                    _dragGrid = parentGrid;
                    _dragElement = element;
                    e.Handled = true;
                }
            }
        }

        private async void BtnOCR_Click(object sender, RoutedEventArgs e)
        {         
            if (!_ocrService.IsAvailable)
            {
                Log("[DEBUG] OCR engine is not available.");
                MessageBox.Show("OCR 엔진을 초기화하지 못했습니다. Windows 언어 설정에서 한국어 OCR 기능이 설치되어 있는지 확인해 주세요.");
                return;
            }

            var document = SelectedDocument;
            if (document == null)
            {
                Log("[DEBUG] OCR skipped: no selected document.");
                MessageBox.Show("OCR을 실행할 PDF를 먼저 열어 주세요.");
                return;
            }

            BtnOCR.IsEnabled = false;
            PbStatus.Visibility = Visibility.Visible;
            PbStatus.Maximum = document.Pages.Count;
            PbStatus.Value = 0;
            TxtStatus.Text = $"OCR 분석 중... ({_ocrService.CurrentLanguage})";
            var previousCursor = Mouse.OverrideCursor;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Render);

                var progress = new Progress<int>(value =>
                {
                    PbStatus.Value = value;
                    TxtStatus.Text = $"OCR 분석 중... {value}/{document.Pages.Count}";
                });
                await _ocrService.RunOcrAsync(document, progress);

                int wordCount = document.Pages.Sum(p => p.OcrWords?.Count ?? 0);
                TxtStatus.Text = $"OCR 완료. {wordCount:N0}개 단어 - 저장하면 검색 가능한 PDF로 반영됩니다.";
                Log($"[DEBUG] OCR completed. Pages={document.Pages.Count}, Words={wordCount}, Language={_ocrService.CurrentLanguage}");
            }
            catch (Exception ex)
            {
                Log($"[CRITICAL] OCR Error: {ex}");
                TxtStatus.Text = "OCR 오류 발생";
                MessageBox.Show($"OCR 처리 중 오류가 발생했습니다.\n\n{ex.Message}");
            }
            finally
            {
                Mouse.OverrideCursor = previousCursor;
                BtnOCR.IsEnabled = true;
                PbStatus.Visibility = Visibility.Collapsed;
            }
        }

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // [Fix] Force update bindings by clearing focus (handling IME composition/active TextBox)
            Keyboard.ClearFocus();
            FocusManager.SetFocusedElement(this, this);

            if (SelectedDocument == null) return;

            // Capture the document to save in a local variable to prevent NRE if selection changes
            var docToSave = SelectedDocument;

            string originalPath = docToSave.FilePath;
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");
            int currentPageIndex = GetCurrentPageIndex();
            var sv = GetCurrentScrollViewer();
            double vOffset = sv?.VerticalOffset ?? 0;
            double hOffset = sv?.HorizontalOffset ?? 0;
            var previousCursor = Mouse.OverrideCursor;
            var saveWatch = Stopwatch.StartNew();
            void LogSaveStep(string step)
            {
                Log($"[SAVE_PROFILE] {step} elapsed={saveWatch.ElapsedMilliseconds}ms path='{originalPath}'");
                saveWatch.Restart();
            }

            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                TxtStatus.Text = "저장 중...";
                LogSaveStep("start");
                await _pdfService.SavePdf(docToSave, tempPath);
                LogSaveStep("SavePdf temp");
                
                // [Fix] Close every tab for this path so reload cannot select stale document state.
                await CloseOpenDocumentsForPathAsync(originalPath);
                LogSaveStep("CloseOpenDocuments");
                
                await Task.Run(() => File.Copy(tempPath, originalPath, true));
                LogSaveStep("CopyTempToOriginal");
                PdfService.ClearPageCacheForFile(originalPath);
                
                _historyService.SetLastPage(originalPath, currentPageIndex);
                _historyService.SaveHistory();
                LogSaveStep("SaveHistory");
                
                await ReloadPdfFromPathAsync(originalPath, currentPageIndex);
                LogSaveStep("ReloadPdfFromPath");

                var newSv = GetCurrentScrollViewer();
                if (newSv != null)
                {
                    newSv.ScrollToVerticalOffset(vOffset);
                    newSv.ScrollToHorizontalOffset(hOffset);
                    
                    // [Fix] Force re-render visible pages after reload and scroll
                    await Dispatcher.InvokeAsync(() => 
                    {
                        if (SelectedDocument != null)
                        {
                            int currentIdx = GetCurrentPageIndex();
                            int start = Math.Max(0, currentIdx - 2);
                            int end = Math.Min(SelectedDocument.Pages.Count - 1, currentIdx + 5);

                            for (int i = start; i <= end; i++)
                            {
                                var p = SelectedDocument.Pages[i];
                                var cts = new System.Threading.CancellationTokenSource();
                                p.RenderCts?.Cancel();
                                p.RenderCts = cts;
                                _pdfService.RenderPageAsync(SelectedDocument, p, cts.Token);
                            }
                        }
                    }, System.Windows.Threading.DispatcherPriority.ContextIdle);
                }
                LogSaveStep("RestoreScrollAndVisibleRender");
                
                TxtStatus.Text = "저장 완료";
                MessageBox.Show("저장되었습니다.");
            }
            catch (Exception ex) 
            { 
                Log($"[CRITICAL] Save Error: {ex}"); 
                MessageBox.Show($"저장 오류: {ex.Message}"); 
            }
            finally 
            { 
                Mouse.OverrideCursor = previousCursor;
                if (File.Exists(tempPath)) try { File.Delete(tempPath); } catch { } 
            } 
        }

        private async void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            Keyboard.ClearFocus();
            FocusManager.SetFocusedElement(this, this);

            if (SelectedDocument == null) return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_annotated" };
            if (dlg.ShowDialog() == true)
            {
                var docToSave = SelectedDocument;
                string savePath = dlg.FileName;
                int currentPageIndex = GetCurrentPageIndex();
                var sv = GetCurrentScrollViewer();
                double vOffset = sv?.VerticalOffset ?? 0;
                double hOffset = sv?.HorizontalOffset ?? 0;
                var previousCursor = Mouse.OverrideCursor;
                var saveWatch = Stopwatch.StartNew();
                void LogSaveStep(string step)
                {
                    Log($"[SAVE_PROFILE] {step} elapsed={saveWatch.ElapsedMilliseconds}ms path='{savePath}'");
                    saveWatch.Restart();
                }

                try {
                    Mouse.OverrideCursor = Cursors.Wait;
                    TxtStatus.Text = "다른 이름으로 저장 중...";
                    LogSaveStep("start");
                    await _pdfService.SavePdf(docToSave, savePath);
                    LogSaveStep("SavePdf target");
                    PdfService.ClearPageCacheForFile(savePath);

                    _historyService.SetLastPage(savePath, currentPageIndex);
                    _historyService.SaveHistory();
                    LogSaveStep("SaveHistory");

                    await ReloadPdfFromPathAsync(savePath, currentPageIndex);
                    LogSaveStep("ReloadPdfFromPath");

                    var newSv = GetCurrentScrollViewer();
                    if (newSv != null)
                    {
                        newSv.ScrollToVerticalOffset(vOffset);
                        newSv.ScrollToHorizontalOffset(hOffset);
                    }
                    LogSaveStep("RestoreScroll");

                    TxtStatus.Text = "저장 완료";
                    MessageBox.Show("저장되었습니다.");
                } catch (Exception ex) { Log($"[CRITICAL] SaveAs Error: {ex}"); MessageBox.Show($"저장 실패: {ex.Message}"); }
                finally { Mouse.OverrideCursor = previousCursor; }
            }
        }

        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e)
        {
            if (_isSelectionTextPending || string.IsNullOrEmpty(_selectedTextBuffer))
                return;

            Clipboard.SetText(_selectedTextBuffer);
            SelectionPopup.IsOpen = false;
        }
        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);
        
        private void AddAnnotation(Color color, AnnotationType type)
        {
            if (_selectedPageIndex == -1 || SelectedDocument == null) return;
            var p = SelectedDocument.Pages[_selectedPageIndex];
            var ann = new PdfAnnotation { X = p.SelectionX, Y = p.SelectionY, Width = p.SelectionWidth, Height = p.SelectionHeight, Type = type, AnnotationColor = color };
            if (type == AnnotationType.Highlight) ann.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B));
            else { ann.Background = new SolidColorBrush(color); ann.Height = 2; ann.Y = p.SelectionY + p.SelectionHeight - 2; }
            p.Annotations.Add(ann);
            SelectionPopup.IsOpen = false; p.IsSelecting = false;
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) => AdjustDocumentZoom(1);
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) => AdjustDocumentZoom(-1);
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e) {
            if (SelectedDocument != null && SelectedDocument.Pages.Count > 0) {
                double viewWidth = MainTabControl.ActualWidth - 60;
                if (viewWidth > 0 && SelectedDocument.Pages[0].Width > 0) SelectedDocument.Zoom = viewWidth / SelectedDocument.Pages[0].Width;
            }
        }
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e) {
            if (SelectedDocument != null && SelectedDocument.Pages.Count > 0) {
                double viewHeight = MainTabControl.ActualHeight - 60;
                if (viewHeight > 0 && SelectedDocument.Pages[0].Height > 0) SelectedDocument.Zoom = viewHeight / SelectedDocument.Pages[0].Height;
            }
        }
        private void DocumentZoom_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control)
                return;

            if (IsThumbnailWheelEvent(sender, e))
                AdjustThumbnailZoom(e.Delta > 0 ? 1 : -1);
            else
                AdjustDocumentZoom(e.Delta > 0 ? 1 : -1);

            e.Handled = true;
        }

        private void AdjustDocumentZoom(int direction)
        {
            if (SelectedDocument == null || direction == 0)
                return;

            double nextZoom = SelectedDocument.Zoom + (direction > 0 ? 0.1 : -0.1);
            SelectedDocument.Zoom = Math.Max(0.1, Math.Min(5.0, nextZoom));
        }

        private void AdjustThumbnailZoom(int direction)
        {
            if (direction == 0)
                return;

            ThumbnailImageWidth += direction > 0 ? 14 : -14;
            if (SelectedDocument != null)
                StartThumbnailRendering(SelectedDocument);
        }

        private bool IsThumbnailWheelEvent(object sender, MouseWheelEventArgs e)
        {
            if (ReferenceEquals(sender, PageThumbnailList))
                return true;

            return e.OriginalSource is DependencyObject source && IsDescendantOf(source, PageThumbnailList);
        }

        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedAnnotation != null && SelectedDocument != null) { 
                foreach (var p in SelectedDocument.Pages) if (p.Annotations.Contains(_selectedAnnotation)) { p.Annotations.Remove(_selectedAnnotation); _selectedAnnotation = null; CheckToolbarVisibility(); break; }
            } else if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a && SelectedDocument != null) {
                foreach (var p in SelectedDocument.Pages) if (p.Annotations.Contains(a)) { p.Annotations.Remove(a); break; }
            }
        }
        private void CheckToolbarVisibility()
        {
            if (ViewToolsToolbar != null)
                ViewToolsToolbar.Visibility = IsEditMode ? Visibility.Collapsed : Visibility.Visible;

            if (EditToolsToolbar != null)
                EditToolsToolbar.Visibility = IsEditMode ? Visibility.Visible : Visibility.Collapsed;

            bool shouldShowTextTools = IsEditMode &&
                (_currentTool == "TEXT" ||
                 (_selectedAnnotation != null && _selectedAnnotation.Type == AnnotationType.FreeText));
            if (TextStyleToolbar != null)
                TextStyleToolbar.Visibility = shouldShowTextTools ? Visibility.Visible : Visibility.Collapsed;

            if (BtnDeleteAnnotation != null)
                BtnDeleteAnnotation.IsEnabled = IsEditMode && _selectedAnnotation != null;
        }

        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if ((sender as Thumb)?.DataContext is PdfAnnotation ann) {
                double minWidth = ann.Type == AnnotationType.ImageStamp ? 20 : 50;
                double minHeight = ann.Type == AnnotationType.ImageStamp ? 20 : 30;
                double width = Math.Max(minWidth, ann.Width + e.HorizontalChange);
                double height = Math.Max(minHeight, ann.Height + e.VerticalChange);
                if (Math.Abs(width - ann.Width) >= 0.5) ann.Width = width;
                if (Math.Abs(height - ann.Height) >= 0.5) ann.Height = height;
            }
        }
        private void UpdateToolbarFromAnnotation(PdfAnnotation ann)
        {
            _isUpdatingUiFromSelection = true;
            try {
                CbFont.SelectedItem = ann.FontFamily; CbSize.SelectedItem = ann.FontSize; BtnBold.IsChecked = ann.IsBold;
                if (ann.Foreground is SolidColorBrush brush) {
                    foreach (ComboBoxItem item in CbColor.Items) {
                        string cn = item.Tag?.ToString() ?? "";
                        Color c = cn=="Red" ? Colors.Red : cn=="Blue" ? Colors.Blue : cn=="Green" ? Colors.Green : cn=="Orange" ? Colors.Orange : Colors.Black;
                        if (brush.Color == c) { CbColor.SelectedItem = item; break; }
                    }
                }
            } finally { _isUpdatingUiFromSelection = false; }
        }
        private void StyleChanged(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _isUpdatingUiFromSelection) return;
            _defaultFontFamily = CbFont.SelectedItem?.ToString() ?? "Malgun Gothic";
            if (CbSize.SelectedItem != null) _defaultFontSize = (double)CbSize.SelectedItem;
            _defaultIsBold = BtnBold.IsChecked == true;
            if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null) {
                string cn = item.Tag.ToString() ?? "Black";
                _defaultFontColor = cn=="Red" ? Colors.Red : cn=="Blue" ? Colors.Blue : cn=="Green" ? Colors.Green : cn=="Orange" ? Colors.Orange : Colors.Black;
            }
            if (_selectedAnnotation != null) {
                _selectedAnnotation.FontFamily = _defaultFontFamily; _selectedAnnotation.FontSize = _defaultFontSize;
                _selectedAnnotation.IsBold = _defaultIsBold; _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor);
            }
        }

        private void Tool_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is RadioButton rb && rb.IsChecked == true)
                {
                    string content = rb.Content.ToString();
                    if (content.Contains("선택")) _currentTool = "CURSOR";
                    else if (content.Contains("형광펜")) _currentTool = "HIGHLIGHT";
                    else if (content.Contains("텍스트"))
                    {
                        _currentTool = "TEXT";
                        IsEditMode = true;
                        if (RbEditMode != null) RbEditMode.IsChecked = true;
                    }
                    CheckToolbarVisibility();
                }
            }
            catch (Exception ex)
            {
                Log($"[CRITICAL] Tool_Checked Error: {ex}");
            }
        }

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || sender is not RadioButton rb || rb.IsChecked != true)
                return;

            IsEditMode = rb == RbEditMode;
            if (!IsEditMode && _currentTool == "TEXT")
            {
                _currentTool = "CURSOR";
                if (RbCursor != null) RbCursor.IsChecked = true;
            }
            else if (IsEditMode && _currentTool == "HIGHLIGHT")
            {
                _currentTool = "CURSOR";
                if (RbCursor != null) RbCursor.IsChecked = true;
            }

            CheckToolbarVisibility();
        }

        private void ListView_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Log($"[DEBUG] ListView_PreviewMouseDown: OriginalSource={e.OriginalSource}, Type={e.OriginalSource?.GetType().Name}");
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.V && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && !(Keyboard.FocusedElement is TextBox)) { PasteClipboardImageToCurrentPage(); e.Handled = true; }
            else if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; }
            else if (e.Key == Key.Delete) BtnDeleteAnnotation_Click(this, new RoutedEventArgs());
            else if (e.Key == Key.Escape && _selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); }
            if (e.Key == Key.Space && !_isSpacePressed && _currentTool != "TEXT" && !(Keyboard.FocusedElement is TextBox)) { _isSpacePressed = true; Mouse.OverrideCursor = Cursors.Hand; e.Handled = true; }
            else if (e.Key == Key.PageUp || e.Key == Key.PageDown) { var sv = GetScrollViewer(); if (sv != null) { if (e.Key == Key.PageDown) sv.PageDown(); else sv.PageUp(); e.Handled = true; } }
        }

        private void PasteClipboardImageToCurrentPage()
        {
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0)
                return;

            IsEditMode = true;
            if (RbEditMode != null) RbEditMode.IsChecked = true;

            if (!Clipboard.ContainsImage())
            {
                TxtStatus.Text = "클립보드에 이미지가 없습니다.";
                return;
            }

            BitmapSource? image = Clipboard.GetImage();
            if (image == null)
            {
                TxtStatus.Text = "클립보드 이미지를 읽지 못했습니다.";
                return;
            }

            image.Freeze();
            byte[] imageBytes = EncodeClipboardImageAsJpeg(image);
            if (imageBytes.Length == 0)
            {
                TxtStatus.Text = "클립보드 이미지를 변환하지 못했습니다.";
                return;
            }

            int pageIndex = Math.Max(0, Math.Min(GetCurrentPageIndex(), SelectedDocument.Pages.Count - 1));
            var page = SelectedDocument.Pages[pageIndex];

            double dpiX = image.DpiX > 0 ? image.DpiX : 96;
            double dpiY = image.DpiY > 0 ? image.DpiY : 96;
            double naturalWidth = image.PixelWidth * 72.0 / dpiX * SelectedDocument.Zoom;
            double naturalHeight = image.PixelHeight * 72.0 / dpiY * SelectedDocument.Zoom;
            if (naturalWidth <= 0 || naturalHeight <= 0)
                return;

            double maxWidth = Math.Max(40, page.Width * 0.8);
            double maxHeight = Math.Max(40, page.Height * 0.8);
            double scale = Math.Min(1.0, Math.Min(maxWidth / naturalWidth, maxHeight / naturalHeight));
            double width = Math.Max(20, naturalWidth * scale);
            double height = Math.Max(20, naturalHeight * scale);

            var annotation = new PdfAnnotation
            {
                Type = AnnotationType.ImageStamp,
                X = Math.Max(0, (page.Width - width) / 2),
                Y = Math.Max(0, (page.Height - height) / 2),
                Width = width,
                Height = height,
                ImageSource = CreateDisplayImageSource(imageBytes, width, height) ?? image,
                ImageBytes = imageBytes,
                IsSelected = true
            };

            if (_selectedAnnotation != null)
                _selectedAnnotation.IsSelected = false;

            page.Annotations.Add(annotation);
            _selectedAnnotation = annotation;
            _activePageIndex = page.PageIndex;
            RefreshCanvasIfVisible(page);
            CheckToolbarVisibility();
            TxtStatus.Text = $"{page.PageIndex + 1}페이지에 클립보드 이미지를 붙였습니다.";
        }

        private static byte[] EncodeClipboardImageAsJpeg(BitmapSource source)
        {
            BitmapSource sourceToEncode = DownsampleForImageStamp(source);
            var encoder = new JpegBitmapEncoder { QualityLevel = 96 };
            encoder.Frames.Add(BitmapFrame.Create(sourceToEncode));
            using var stream = new MemoryStream();
            encoder.Save(stream);
            return stream.ToArray();
        }

        private static BitmapSource DownsampleForImageStamp(BitmapSource source)
        {
            int maxSide = Math.Max(source.PixelWidth, source.PixelHeight);
            if (maxSide <= MaxImageStampPixelSide)
                return source;

            double scale = MaxImageStampPixelSide / (double)maxSide;
            var scaled = new TransformedBitmap(source, new ScaleTransform(scale, scale));
            scaled.Freeze();
            return scaled;
        }

        private static BitmapSource? CreateDisplayImageSource(byte[] imageBytes, double displayWidth, double displayHeight)
        {
            if (imageBytes.Length == 0)
                return null;

            try
            {
                var bitmap = new BitmapImage();
                using var stream = new MemoryStream(imageBytes);
                bitmap.BeginInit();
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                int decodeWidth = GetDecodePixelWidth(displayWidth, displayHeight);
                if (decodeWidth > 0)
                    bitmap.DecodePixelWidth = decodeWidth;
                bitmap.StreamSource = stream;
                bitmap.EndInit();
                bitmap.Freeze();
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        private static int GetDecodePixelWidth(double displayWidth, double displayHeight)
        {
            double maxDisplaySide = Math.Max(displayWidth, displayHeight);
            if (maxDisplaySide <= 0)
                return 0;

            return (int)Math.Max(64, Math.Min(2048, Math.Ceiling(maxDisplaySide * 2)));
        }

        private void RefreshCanvasIfVisible(PdfPageViewModel page)
        {
            var canvas = FindPageCanvas(page);
            if (canvas != null)
                RefreshCanvas(canvas, page);
        }

        private void ApplyInteractionModeToVisibleCanvases()
        {
            if (SelectedDocument == null)
                return;

            foreach (var page in SelectedDocument.Pages)
            {
                var canvas = FindPageCanvas(page);
                if (canvas != null)
                    canvas.IsHitTestVisible = IsEditMode;
            }
        }

        protected override void OnKeyUp(KeyEventArgs e) { base.OnKeyUp(e); if (e.Key == Key.Space) { _isSpacePressed = false; _isPanning = false; Mouse.OverrideCursor = null; GetCurrentListView()?.ReleaseMouseCapture(); } }

        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++) {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                if (v is T child) return child;
                T? result = GetVisualChild<T>(v);
                if (result != null) return result;
            }
            return default;
        }

        private ListView? GetCurrentListView() => GetVisualChild<ListView>(MainTabControl);
        private T? FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null) return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++) {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild) return typedChild;
                var result = FindChild<T>(child);
                if (result != null) return result;
            }
            return null;
        }

        private ScrollViewer? GetScrollViewer() { if (_cachedScrollViewer == null) { var lv = GetCurrentListView(); if (lv != null) _cachedScrollViewer = FindChild<ScrollViewer>(lv); } return _cachedScrollViewer; }
        private ScrollViewer? GetCurrentScrollViewer() { var lv = GetCurrentListView(); return lv != null ? FindChild<ScrollViewer>(lv) : null; }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));

        private void BtnSign_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null) return;
            var res = MessageBox.Show("이미지 도장을 사용하시겠습니까?", "서명 방식", MessageBoxButton.YesNoCancel);
            if (res == MessageBoxResult.Cancel) return;
            string? path = null;
            if (res == MessageBoxResult.Yes) {
                var fd = new OpenFileDialog { Filter = "Images|*.png;*.jpg;*.jpeg;*.bmp" };
                if (fd.ShowDialog() == true) path = fd.FileName; else return;
            }
            if (SelectedDocument.Pages.Count == 0) return;
            var pVM = SelectedDocument.Pages[0];
            var ph = new PdfAnnotation { Type = AnnotationType.SignaturePlaceholder, Width = 120, Height = 60, X = (pVM.Width - 120) / 2, Y = (pVM.Height - 60) / 2, IsSelected = true, VisualStampPath = path };
            pVM.Annotations.Add(ph);
        }

        private void BtnVerifySignature_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null || !(sender is FrameworkElement el) || !(el.DataContext is PdfAnnotation ann)) return;
            if (ann.Type == AnnotationType.SignatureField) {
                try {
                    int pageIndex = -1;
                    PdfPageViewModel? pageVM = null;
                    for (int i = 0; i < SelectedDocument.Pages.Count; i++) {
                        if (SelectedDocument.Pages[i].Annotations.Contains(ann)) {
                            pageIndex = i;
                            pageVM = SelectedDocument.Pages[i];
                            break;
                        }
                    }

                    if (pageIndex == -1 || pageVM == null) {
                        MessageBox.Show("서명 데이터를 찾을 수 없습니다.");
                        return;
                    }

                    double scaleX = pageVM.PdfPageWidthPoint / pageVM.Width;
                    double scaleY = pageVM.PdfPageHeightPoint / pageVM.Height;
                    var pdfRect = new SignaturePdfRect(
                        ann.X * scaleX,
                        pageVM.PdfPageHeightPoint - ((ann.Y + ann.Height) * scaleY),
                        ann.Width * scaleX,
                        ann.Height * scaleY);

                    var result = _signatureVerificationService.VerifySignatureForAnnotation(
                        SelectedDocument.FilePath,
                        pageIndex,
                        ann.FieldName,
                        pdfRect);

                    if (result != null) {
                        new SignatureResultWindow(result) { Owner = this }.ShowDialog();
                    } else {
                        MessageBox.Show("서명 데이터를 찾을 수 없습니다.");
                    }
                } catch (Exception ex) { MessageBox.Show($"검증 오류: {ex.Message}"); }
            }
        }

        private void BtnToggleSidebar_Click(object sender, RoutedEventArgs e)
        {
            if (SidebarBorder.Visibility == Visibility.Visible)
            {
                if (SidebarColumn.ActualWidth > 0)
                    _lastSidebarWidth = new GridLength(SidebarColumn.ActualWidth);

                SidebarBorder.Visibility = Visibility.Collapsed;
                SidebarSplitter.Visibility = Visibility.Collapsed;
                SidebarColumn.MinWidth = 0;
                SidebarColumn.Width = new GridLength(0);
                return;
            }

            SidebarColumn.MinWidth = 180;
            SidebarColumn.Width = _lastSidebarWidth.Value > 0
                ? _lastSidebarWidth
                : new GridLength(250);
            SidebarBorder.Visibility = Visibility.Visible;
            SidebarSplitter.Visibility = Visibility.Visible;

            if (SelectedDocument != null)
                StartThumbnailRendering(SelectedDocument);
        }
        private void ScrollToPage(int pageIndex)
        {
            if (SelectedDocument == null || pageIndex < 0 || pageIndex >= SelectedDocument.Pages.Count) return;

            RenderPagesAround(pageIndex);

            var sv = GetCurrentScrollViewer();
            if (sv == null) return;

            double targetOffset = 0;
            for (int i = 0; i < pageIndex; i++)
            {
                var p = SelectedDocument.Pages[i];
                // Item Height = PageHeight + Margin(20) + Border(2)
                targetOffset += p.Height + 22;
            }

            sv.ScrollToVerticalOffset(targetOffset);

            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (SelectedDocument == null || pageIndex < 0 || pageIndex >= SelectedDocument.Pages.Count) return;
                GetCurrentListView()?.ScrollIntoView(SelectedDocument.Pages[pageIndex]);
                RenderPagesAround(pageIndex, 2, 4);
            }), System.Windows.Threading.DispatcherPriority.ContextIdle);
        }

        private void RenderPagesAround(int pageIndex, int pagesBefore = 1, int pagesAfter = 2)
        {
            if (SelectedDocument == null) return;

            int start = Math.Max(0, pageIndex - pagesBefore);
            int end = Math.Min(SelectedDocument.Pages.Count - 1, pageIndex + pagesAfter);
            for (int i = start; i <= end; i++)
            {
                var page = SelectedDocument.Pages[i];
                page.RenderCts?.Cancel();
                var cts = new System.Threading.CancellationTokenSource();
                page.RenderCts = cts;
                _pdfService.RenderPageAsync(SelectedDocument, page, cts.Token);
            }
        }

        private void BookmarkTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e) {
            if (e.NewValue is PdfBookmarkViewModel bm && SelectedDocument != null) {
                ScrollToPage(bm.PageIndex);
            }
        }

        private Point _startPoint;
        private bool _isDragging = false;
        private PdfBookmarkViewModel? _draggedBookmark;
        private void BookmarkTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);
            _draggedBookmark = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource)?.DataContext as PdfBookmarkViewModel;
            if (_draggedBookmark != null)
                _draggedBookmark.IsSelected = true;
        }

        private void BookmarkTree_PreviewMouseMove(object sender, MouseEventArgs e) {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging) {
                Point pos = e.GetPosition(null); Vector diff = _startPoint - pos;
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) {
                    var bm = _draggedBookmark ?? BookmarkTree.SelectedItem as PdfBookmarkViewModel;
                    if (bm != null && !bm.IsEditing) {
                        _isDragging = true;
                        DragDrop.DoDragDrop(BookmarkTree, new DataObject("MinsBookmark", bm), DragDropEffects.Move);
                        _isDragging = false;
                        _draggedBookmark = null;
                    }
                }
            }
        }
        private void BookmarkTree_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else if (e.Data.GetDataPresent("MinsBookmark") &&
                     CanDropBookmark(e.Data.GetData("MinsBookmark") as PdfBookmarkViewModel, GetBookmarkDropTarget(e)))
            {
                e.Effects = DragDropEffects.Move;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }

            e.Handled = true;
        }

        private void BookmarkTree_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) { Window_Drop(sender, e); return; }
            if (!e.Data.GetDataPresent("MinsBookmark") || SelectedDocument == null) return;
            var src = e.Data.GetData("MinsBookmark") as PdfBookmarkViewModel;
            var dropTarget = GetBookmarkDropTarget(e);
            if (!CanDropBookmark(src, dropTarget)) return;

            MoveBookmark(src!, dropTarget);
            MarkBookmarksChanged();
            e.Handled = true;
        }

        private BookmarkDropTarget GetBookmarkDropTarget(DragEventArgs e)
        {
            var item = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource);
            if (item?.DataContext is not PdfBookmarkViewModel target)
                return new BookmarkDropTarget(null, BookmarkDropPlacement.After);

            var position = e.GetPosition(item);
            double third = Math.Max(1.0, item.ActualHeight / 3.0);
            BookmarkDropPlacement placement =
                position.Y < third ? BookmarkDropPlacement.Before :
                position.Y > item.ActualHeight - third ? BookmarkDropPlacement.After :
                BookmarkDropPlacement.AsChild;

            return new BookmarkDropTarget(target, placement);
        }

        private bool CanDropBookmark(PdfBookmarkViewModel? source, BookmarkDropTarget target)
        {
            if (source == null || SelectedDocument == null)
                return false;

            if (target.Bookmark == null)
                return true;

            return target.Bookmark != source && !IsChildOf(target.Bookmark, source);
        }

        private void MoveBookmark(PdfBookmarkViewModel source, BookmarkDropTarget target)
        {
            if (SelectedDocument == null)
                return;

            GetCurrentCollection(source)?.Remove(source);

            if (target.Bookmark == null)
            {
                SelectedDocument.Bookmarks.Add(source);
                source.Parent = null;
                source.IsSelected = true;
                return;
            }

            if (target.Placement == BookmarkDropPlacement.AsChild)
            {
                target.Bookmark.Children.Add(source);
                source.Parent = target.Bookmark;
                target.Bookmark.IsExpanded = true;
                source.IsSelected = true;
                return;
            }

            var targetCollection = GetCurrentCollection(target.Bookmark);
            if (targetCollection == null)
                return;

            int targetIndex = targetCollection.IndexOf(target.Bookmark);
            if (targetIndex < 0)
                targetIndex = targetCollection.Count - 1;

            int insertIndex = target.Placement == BookmarkDropPlacement.Before
                ? targetIndex
                : targetIndex + 1;
            insertIndex = Math.Max(0, Math.Min(insertIndex, targetCollection.Count));

            targetCollection.Insert(insertIndex, source);
            source.Parent = target.Bookmark.Parent;
            source.IsSelected = true;
        }

        private bool IsChildOf(PdfBookmarkViewModel c, PdfBookmarkViewModel p) { var cur = c.Parent; while (cur != null) { if (cur == p) return true; cur = cur.Parent; } return false; }
        private static T? FindAncestor<T>(DependencyObject cur) where T : DependencyObject
        {
            while (cur != null)
            {
                if (cur is T t) return t;
                cur = cur is Visual or Visual3D
                    ? VisualTreeHelper.GetParent(cur)
                    : LogicalTreeHelper.GetParent(cur);
            }
            return null;
        }

        private enum BookmarkDropPlacement
        {
            Before,
            AsChild,
            After
        }

        private readonly record struct BookmarkDropTarget(PdfBookmarkViewModel? Bookmark, BookmarkDropPlacement Placement);
        private void MarkBookmarksChanged()
        {
            if (SelectedDocument != null) SelectedDocument.HasBookmarkChanges = true;
        }

        private void BookmarkTree_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.F2 && BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) bm.IsEditing = true; }
        private void BtnRenameBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) bm.IsEditing = true; }
        private void BookmarkRename_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter && (sender as TextBox)?.DataContext is PdfBookmarkViewModel bm) { bm.IsEditing = false; MarkBookmarksChanged(); e.Handled = true; } }
        private void BookmarkRename_LostFocus(object sender, RoutedEventArgs e) { if ((sender as TextBox)?.DataContext is PdfBookmarkViewModel bm) { bm.IsEditing = false; MarkBookmarksChanged(); } }
        private void BookmarkRename_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (e.NewValue is true && sender is TextBox textBox)
            {
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    textBox.Focus();
                    textBox.SelectAll();
                }), System.Windows.Threading.DispatcherPriority.Input);
            }
        }
        private void PdfListView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (SelectedDocument != null)
            {
                int pageIndex = GetCurrentPageIndex();
                TxtPageInfo.Text = $"{pageIndex + 1} / {SelectedDocument.Pages.Count}";
                SyncSelectedThumbnail(pageIndex);
            }
        }

        private void PageThumbnailList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_syncingThumbnailSelection || SelectedDocument == null)
                return;

            if (PageThumbnailList.SelectedItems.Count == 1 && PageThumbnailList.SelectedItem is PdfPageViewModel page)
                ScrollToPage(page.PageIndex);
        }

        private void PageThumbnailList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _thumbnailDragStartPoint = e.GetPosition(PageThumbnailList);
            _draggedThumbnailPage = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource)?.DataContext as PdfPageViewModel;
            _draggedThumbnailPages = GetSelectedThumbnailPages(_draggedThumbnailPage);
        }

        private void PageThumbnailItem_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not ListBoxItem item)
                return;

            _thumbnailDragStartPoint = e.GetPosition(PageThumbnailList);
            _draggedThumbnailPage = item.DataContext as PdfPageViewModel;
            _draggedThumbnailPages = GetSelectedThumbnailPages(_draggedThumbnailPage);
        }

        private void PageThumbnailList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            TryStartThumbnailDrag(e);
        }

        private void PageThumbnailItem_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            TryStartThumbnailDrag(e);
        }

        private void PageThumbnailList_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            FinishThumbnailDrag(e);
        }

        private void PageThumbnailItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            FinishThumbnailDrag(e);
        }

        private void TryStartThumbnailDrag(MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _draggedThumbnailPage == null || _draggedThumbnailPages.Count == 0)
                return;

            var currentPoint = e.GetPosition(PageThumbnailList);
            if (Math.Abs(currentPoint.X - _thumbnailDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(currentPoint.Y - _thumbnailDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            _isThumbnailDragging = true;
            Mouse.Capture(PageThumbnailList);
            PageThumbnailList.Cursor = Cursors.SizeAll;
            TxtStatus.Text = _draggedThumbnailPages.Count == 1
                ? "페이지 이동 중..."
                : $"{_draggedThumbnailPages.Count}개 페이지 이동 중...";
            e.Handled = true;
        }

        private void FinishThumbnailDrag(MouseButtonEventArgs e)
        {
            if (!_isThumbnailDragging || SelectedDocument == null || _draggedThumbnailPage == null || _draggedThumbnailPages.Count == 0)
            {
                ResetThumbnailDragState();
                return;
            }

            var dropPoint = e.GetPosition(PageThumbnailList);
            int targetIndex = GetThumbnailDropIndex(dropPoint);

            if (targetIndex >= 0)
            {
                MoveThumbnailPages(_draggedThumbnailPages, targetIndex);
            }

            ResetThumbnailDragState();
            e.Handled = true;
        }

        private void ResetThumbnailDragState()
        {
            _isThumbnailDragging = false;
            _draggedThumbnailPage = null;
            _draggedThumbnailPages = new List<PdfPageViewModel>();
            PageThumbnailList.Cursor = null;
            if (Mouse.Captured == PageThumbnailList)
                Mouse.Capture(null);
        }

        private void PageThumbnailList_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent("MinsPageThumbnail") || e.Data.GetDataPresent(typeof(PdfPageViewModel))
                ? DragDropEffects.Move
                : DragDropEffects.None;
            e.Handled = true;
        }

        private void PageThumbnailList_Drop(object sender, DragEventArgs e)
        {
            if (SelectedDocument == null)
                return;

            var sourcePage = e.Data.GetData("MinsPageThumbnail") as PdfPageViewModel
                ?? e.Data.GetData(typeof(PdfPageViewModel)) as PdfPageViewModel;
            if (sourcePage == null || !SelectedDocument.Pages.Contains(sourcePage))
                return;

            int sourceIndex = SelectedDocument.Pages.IndexOf(sourcePage);
            int targetIndex = GetThumbnailDropIndex(e);
            if (sourceIndex < 0 || targetIndex < 0)
                return;

            MoveThumbnailPages(new[] { sourcePage }, targetIndex);
            e.Handled = true;
        }

        private List<PdfPageViewModel> GetSelectedThumbnailPages(PdfPageViewModel? fallbackPage)
        {
            if (SelectedDocument == null)
                return new List<PdfPageViewModel>();

            var selectedPages = PageThumbnailList.SelectedItems
                .OfType<PdfPageViewModel>()
                .Where(SelectedDocument.Pages.Contains)
                .OrderBy(p => SelectedDocument.Pages.IndexOf(p))
                .ToList();

            if (fallbackPage != null && !selectedPages.Contains(fallbackPage))
                selectedPages = new List<PdfPageViewModel> { fallbackPage };

            return selectedPages;
        }

        private void MoveThumbnailPages(IEnumerable<PdfPageViewModel> sourcePages, int targetIndex)
        {
            if (SelectedDocument == null)
                return;

            var pagesToMove = sourcePages
                .Where(SelectedDocument.Pages.Contains)
                .Distinct()
                .OrderBy(p => SelectedDocument.Pages.IndexOf(p))
                .ToList();
            if (pagesToMove.Count == 0)
                return;

            var oldOrder = SelectedDocument.Pages.ToList();
            var oldIndexes = pagesToMove.Select(p => oldOrder.IndexOf(p)).Where(i => i >= 0).OrderBy(i => i).ToList();

            int adjustedTargetIndex = targetIndex;
            foreach (int oldIndex in oldIndexes)
            {
                if (oldIndex < targetIndex)
                    adjustedTargetIndex--;
            }

            foreach (var page in pagesToMove)
                SelectedDocument.Pages.Remove(page);

            adjustedTargetIndex = Math.Max(0, Math.Min(adjustedTargetIndex, SelectedDocument.Pages.Count));
            for (int i = 0; i < pagesToMove.Count; i++)
                SelectedDocument.Pages.Insert(adjustedTargetIndex + i, pagesToMove[i]);

            var newOrder = SelectedDocument.Pages.ToList();
            if (oldOrder.SequenceEqual(newOrder))
                return;

            RefreshPageIndexes();
            _syncingThumbnailSelection = true;
            try
            {
                PageThumbnailList.SelectedItems.Clear();
                foreach (var page in pagesToMove)
                    PageThumbnailList.SelectedItems.Add(page);
            }
            finally
            {
                _syncingThumbnailSelection = false;
            }

            var firstMovedPage = pagesToMove[0];
            TxtStatus.Text = pagesToMove.Count == 1 ? "페이지 순서 변경됨" : $"{pagesToMove.Count}개 페이지 순서 변경됨";
            TxtPageInfo.Text = $"{firstMovedPage.PageNumber} / {SelectedDocument.Pages.Count}";

            var previousCursor = Mouse.OverrideCursor;
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                ScrollToPage(firstMovedPage.PageIndex);
            }
            finally
            {
                Mouse.OverrideCursor = previousCursor;
            }
        }

        private int GetThumbnailDropIndex(DragEventArgs e)
        {
            if (SelectedDocument == null)
                return -1;

            var targetItem = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
            if (targetItem == null)
                return SelectedDocument.Pages.Count;

            int targetIndex = PageThumbnailList.ItemContainerGenerator.IndexFromContainer(targetItem);
            if (targetIndex < 0)
                return SelectedDocument.Pages.Count;

            var position = e.GetPosition(targetItem);
            if (position.Y > targetItem.ActualHeight / 2)
                targetIndex++;

            return targetIndex;
        }

        private int GetThumbnailDropIndex(Point position)
        {
            if (SelectedDocument == null)
                return -1;

            var hit = PageThumbnailList.InputHitTest(position) as DependencyObject;
            var targetItem = hit != null ? FindAncestor<ListBoxItem>(hit) : null;
            if (targetItem == null)
                return position.Y < 0 ? 0 : SelectedDocument.Pages.Count;

            int targetIndex = PageThumbnailList.ItemContainerGenerator.IndexFromContainer(targetItem);
            if (targetIndex < 0)
                return SelectedDocument.Pages.Count;

            var itemPoint = PageThumbnailList.TranslatePoint(position, targetItem);
            if (itemPoint.Y > targetItem.ActualHeight / 2)
                targetIndex++;

            return targetIndex;
        }

        private void SyncSelectedThumbnail(int pageIndex)
        {
            if (SelectedDocument == null || pageIndex < 0 || pageIndex >= SelectedDocument.Pages.Count)
                return;

            var page = SelectedDocument.Pages[pageIndex];
            if (ReferenceEquals(PageThumbnailList.SelectedItem, page))
                return;

            _syncingThumbnailSelection = true;
            try
            {
                PageThumbnailList.SelectedItem = page;
            }
            finally
            {
                _syncingThumbnailSelection = false;
            }

            StartThumbnailRendering(SelectedDocument);
        }

        private void StartThumbnailRendering(PdfDocumentModel document)
        {
            _thumbnailRenderCts?.Cancel();
            _thumbnailRenderCts = new CancellationTokenSource();
            var token = _thumbnailRenderCts.Token;

            Dispatcher.InvokeAsync(() =>
            {
                if (!token.IsCancellationRequested && ReferenceEquals(SelectedDocument, document))
                    PageThumbnailList.UpdateLayout();
            }, System.Windows.Threading.DispatcherPriority.ContextIdle);

            var pages = GetThumbnailRenderCandidates(document);
            if (pages.Count == 0)
                return;

            _ = Task.Run(async () =>
            {
                try
                {
                    foreach (var page in pages)
                    {
                        token.ThrowIfCancellationRequested();
                        if (document.IsDisposed)
                            return;

                        await _pdfService.RenderThumbnailAsync(document, page, 140, token);
                    }
                }
                catch (OperationCanceledException)
                {
                }
            }, token);
        }

        private List<PdfPageViewModel> GetThumbnailRenderCandidates(PdfDocumentModel document)
        {
            if (!ReferenceEquals(SelectedDocument, document))
                return new List<PdfPageViewModel>();

            var indexes = new SortedSet<int>();
            int currentIndex = GetCurrentPageIndex();
            AddThumbnailRange(indexes, document, currentIndex - ThumbnailRenderPagesBefore, currentIndex + ThumbnailRenderPagesAfter);

            if (PageThumbnailList.SelectedItem is PdfPageViewModel selectedPage)
                AddThumbnailRange(indexes, document, selectedPage.PageIndex - ThumbnailRenderPagesBefore, selectedPage.PageIndex + ThumbnailRenderPagesAfter);

            if (SidebarBorder.Visibility == Visibility.Visible)
            {
                PageThumbnailList.UpdateLayout();
                for (int i = 0; i < document.Pages.Count; i++)
                {
                    if (PageThumbnailList.ItemContainerGenerator.ContainerFromIndex(i) is ListBoxItem item &&
                        IsElementVisibleInThumbnailList(item))
                    {
                        AddThumbnailRange(indexes, document, i - 2, i + 4);
                    }
                }
            }

            return indexes
                .Where(i => i >= 0 && i < document.Pages.Count)
                .Select(i => document.Pages[i])
                .Where(p => p.ThumbnailSource == null)
                .ToList();
        }

        private static void AddThumbnailRange(SortedSet<int> indexes, PdfDocumentModel document, int start, int end)
        {
            int first = Math.Max(0, start);
            int last = Math.Min(document.Pages.Count - 1, end);
            for (int i = first; i <= last; i++)
                indexes.Add(i);
        }

        private bool IsElementVisibleInThumbnailList(FrameworkElement element)
        {
            if (!element.IsVisible || PageThumbnailList.ActualHeight <= 0)
                return false;

            try
            {
                var bounds = element.TransformToAncestor(PageThumbnailList)
                    .TransformBounds(new Rect(0, 0, element.ActualWidth, element.ActualHeight));
                return bounds.Bottom >= -80 && bounds.Top <= PageThumbnailList.ActualHeight + 160;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        private void PageThumbnailList_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (SelectedDocument != null)
                StartThumbnailRendering(SelectedDocument);
        }
        private void BtnAddBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null) return;

            var bookmark = new PdfBookmarkViewModel
            {
                Title = "새 책갈피",
                PageIndex = GetCurrentPageIndex(),
                IsEditing = true
            };
            SelectedDocument.Bookmarks.Add(bookmark);
            MarkBookmarksChanged();
            BookmarkTree.UpdateLayout();
        }
        private void BtnUpdateBookmarkPage_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { int idx = GetCurrentPageIndex(); bm.PageIndex = idx; if (bm.Title.StartsWith("새 책갈피")) bm.Title = $"새 책갈피 (p.{idx + 1})"; MarkBookmarksChanged(); } }
        private void BookmarkItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e) { if ((sender as FrameworkElement)?.DataContext is PdfBookmarkViewModel bm && SelectedDocument != null) { ScrollToPage(bm.PageIndex); } }
        private void BtnDeleteBookmark_Click(object sender, RoutedEventArgs e) { var t = (sender as MenuItem)?.DataContext as PdfBookmarkViewModel ?? BookmarkTree.SelectedItem as PdfBookmarkViewModel; if (t != null && SelectedDocument != null) { if (SelectedDocument.Bookmarks.Contains(t)) SelectedDocument.Bookmarks.Remove(t); else t.Parent?.Children.Remove(t); MarkBookmarksChanged(); } }
        private void BtnMoveBookmarkUp_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i > 0) { col?.Move(i, i - 1); MarkBookmarksChanged(); } } }
        private void BtnMoveBookmarkDown_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i >= 0 && i < col.Count - 1) { col?.Move(i, i + 1); MarkBookmarksChanged(); } } }
        private void BtnIndentBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i > 0) { var p = col[i - 1]; col.Remove(bm); p.Children.Add(bm); bm.Parent = p; p.IsExpanded = true; MarkBookmarksChanged(); } } }
        private void BtnOutdentBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm && bm.Parent != null) { var oldP = bm.Parent; var list = (oldP.Parent == null) ? SelectedDocument.Bookmarks : oldP.Parent.Children; int i = list.IndexOf(oldP); if (i >= 0) { oldP.Children.Remove(bm); list.Insert(i + 1, bm); bm.Parent = oldP.Parent; MarkBookmarksChanged(); } } }
        private System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>? GetCurrentCollection(PdfBookmarkViewModel bm) => SelectedDocument == null ? null : bm.Parent == null ? SelectedDocument.Bookmarks : bm.Parent.Children;

        private int GetCurrentPageIndex()
        {
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0) return 0;
            var sv = GetCurrentScrollViewer();
            if (sv == null) return 0;

            // [Fix] Calculate based on screen center
            double targetY = sv.VerticalOffset + (sv.ViewportHeight / 2.0);
            double currentY = 0;

            foreach (var p in SelectedDocument.Pages)
            {
                // p.Height is already scaled by Zoom (updated in PropertyChanged handler)
                // Add Margin (10 top + 10 bottom) + Border (1 top + 1 bottom) = 22 approx
                double itemHeight = p.Height + 22;

                if (currentY + itemHeight > targetY)
                    return p.PageIndex;

                currentY += itemHeight;
            }
            return SelectedDocument.Pages.Count - 1;
        }

        public async Task OpenPdfFromPath(string path)
        {
            if (!File.Exists(path)) return;
            Log($"[DEBUG] OpenPdfFromPath: {path}");
            var profileWatch = Stopwatch.StartNew();
            void LogOpenStep(string step)
            {
                Log($"[OPEN_PROFILE] {step} elapsed={profileWatch.ElapsedMilliseconds}ms path='{path}'");
                profileWatch.Restart();
            }

            var existing = Documents.FirstOrDefault(d => d.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                SelectedDocument = existing;
                LogOpenStep("select existing document");
                StartThumbnailRendering(existing);
                return;
            }

            var doc = await _pdfService.LoadPdfAsync(path);
            LogOpenStep("LoadPdfAsync");
            if (doc != null) 
            { 
                // [Fix] Zoom changed event handler to update page sizes
                doc.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(PdfDocumentModel.Zoom))
                    {
                        // 1. Update dimensions for ALL pages (required for scrollbar)
                        foreach (var p in doc.Pages)
                        {
                            p.Width = p.PdfPageWidthPoint * doc.Zoom;
                            p.Height = p.PdfPageHeightPoint * doc.Zoom;
                            
                            // Optional: Clear ImageSource to force refresh? 
                            // No, keep old image (blurry) until new one arrives to avoid flashing white.
                        }

                        // 2. Request re-render ONLY for visible pages (optimization)
                        // Virtualization will handle the rest when scrolling, but we need to update current view immediately.
                        int currentIdx = GetCurrentPageIndex();
                        int start = Math.Max(0, currentIdx - 2);
                        int end = Math.Min(doc.Pages.Count - 1, currentIdx + 5); // Render a bit more downwards

                        for (int i = start; i <= end; i++)
                        {
                             var p = doc.Pages[i];
                             var cts = new System.Threading.CancellationTokenSource();
                             p.RenderCts?.Cancel();
                             p.RenderCts = cts;
                             _pdfService.RenderPageAsync(doc, p, cts.Token);
                        }
                    }
                };
                
                await _pdfService.InitializeDocumentAsync(doc); 
                LogOpenStep("InitializeDocumentAsync");
                
                // Initial size calculation based on default Zoom (1.0)
                foreach(var p in doc.Pages)
                {
                    p.Width = p.PdfPageWidthPoint * doc.Zoom;
                    p.Height = p.PdfPageHeightPoint * doc.Zoom;
                }
                LogOpenStep("initial page size calculation");
                
                Documents.Add(doc); 
                SelectedDocument = doc;
                LogOpenStep("add and select document");
                int initialPageIndex = Math.Max(0, Math.Min(_historyService.GetLastPage(path), doc.Pages.Count - 1));
                RenderPagesAround(initialPageIndex, 0, 2);
                LogOpenStep("queue initial visible render");
                StartThumbnailRendering(doc);
                LogOpenStep("queue lazy thumbnails");
                _ = _pdfService.LoadBookmarksAsync(doc);

                await Dispatcher.InvokeAsync(() =>
                {
                    ScrollToPage(initialPageIndex);
                    TxtStatus.Text = initialPageIndex > 0
                        ? $"{initialPageIndex + 1}페이지에서 다시 열었습니다."
                        : "준비";
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
        }

        private async Task ReloadPdfFromPathAsync(string path, int pageIndex = 0)
        {
            await CloseOpenDocumentsForPathAsync(path);
            PdfService.ClearPageCacheForFile(path);

            _cachedScrollViewer = null;
            MainTabControl.SelectedItem = null;
            MainTabControl.Items.Refresh();
            await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.ContextIdle);

            await OpenPdfFromPath(path);

            await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Loaded);
            MainTabControl.Items.Refresh();

            if (SelectedDocument == null)
                return;

            int targetPageIndex = Math.Max(0, Math.Min(pageIndex, SelectedDocument.Pages.Count - 1));
            ScrollToPage(targetPageIndex);
            RenderPagesAround(targetPageIndex);
        }

        private async Task CloseOpenDocumentsForPathAsync(string path)
        {
            PdfService.ClearPageCacheForFile(path);

            var existingDocuments = Documents
                .Where(d => d.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var document in existingDocuments)
            {
                if (SelectedDocument == document)
                {
                    SelectedDocument = null;
                    MainTabControl.SelectedItem = null;
                }

                await document.CloseAsync();
                Documents.Remove(document);
            }

            if (existingDocuments.Count > 0)
            {
                _cachedScrollViewer = null;
                MainTabControl.Items.Refresh();
                await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.ContextIdle);
            }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e) { var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" }; if (dlg.ShowDialog() == true) OpenPdfFromPath(dlg.FileName); }
        private void Window_Drop(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) { var files = (string[])e.Data.GetData(DataFormats.FileDrop); foreach (var f in files) if (Path.GetExtension(f).ToLower() == ".pdf") OpenPdfFromPath(f); } }
        private void Window_DragOver(object sender, DragEventArgs e) { e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : e.Data.GetDataPresent("MinsBookmark") ? DragDropEffects.Move : DragDropEffects.None; e.Handled = true; }
        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Use Keyboard.IsKeyDown for reliable state check instead of tracking events
            bool isSpaceDown = Keyboard.IsKeyDown(Key.Space);
            bool isTextBoxFocused = Keyboard.FocusedElement is TextBox;

            if (isSpaceDown && e.ChangedButton == MouseButton.Left && !isTextBoxFocused)
            {
                _isPanning = true;
                _lastPanPoint = e.GetPosition(this);
                var listView = FindChild<ListView>(MainTabControl);
                if (listView != null) listView.CaptureMouse();
                e.Handled = true;
            }
        }
        private void Window_PreviewMouseMove(object sender, MouseEventArgs e) { if (_isPanning && e.LeftButton == MouseButtonState.Pressed) { var p = e.GetPosition(this); var d = p - _lastPanPoint; var sv = GetScrollViewer(); if (sv != null) { sv.ScrollToHorizontalOffset(sv.HorizontalOffset - d.X); sv.ScrollToVerticalOffset(sv.VerticalOffset - d.Y); } _lastPanPoint = p; e.Handled = true; } }
        private void Window_PreviewMouseUp(object sender, MouseButtonEventArgs e) { if (_isPanning) { _isPanning = false; GetCurrentListView()?.ReleaseMouseCapture(); e.Handled = true; } }
        private void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control || SelectedDocument == null)
                return;

            if (e.OriginalSource is not DependencyObject source)
                return;

            if (IsDescendantOf(source, PageThumbnailList))
            {
                AdjustThumbnailZoom(e.Delta > 0 ? 1 : -1);
                e.Handled = true;
            }
            else if (IsDescendantOf(source, MainTabControl))
            {
                AdjustDocumentZoom(e.Delta > 0 ? 1 : -1);
                e.Handled = true;
            }
        }

        private static bool IsDescendantOf(DependencyObject source, DependencyObject ancestor)
        {
            DependencyObject? current = source;
            while (current != null)
            {
                if (ReferenceEquals(current, ancestor))
                    return true;

                current = current is Visual or Visual3D
                    ? VisualTreeHelper.GetParent(current)
                    : LogicalTreeHelper.GetParent(current);
            }

            return false;
        }
        private void BtnAddBlankPage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null) return;

            int currentPageIndex = GetCurrentPageIndex();
            int insertIndex = currentPageIndex >= 0
                ? Math.Min(currentPageIndex + 1, SelectedDocument.Pages.Count)
                : SelectedDocument.Pages.Count;

            var referencePage = currentPageIndex >= 0 && currentPageIndex < SelectedDocument.Pages.Count
                ? SelectedDocument.Pages[currentPageIndex]
                : SelectedDocument.Pages.FirstOrDefault();

            double pageWidth = referencePage?.Width > 0 ? referencePage.Width : 595;
            double pageHeight = referencePage?.Height > 0 ? referencePage.Height : 842;
            double pdfWidth = referencePage?.PdfPageWidthPoint > 0 ? referencePage.PdfPageWidthPoint : pageWidth;
            double pdfHeight = referencePage?.PdfPageHeightPoint > 0 ? referencePage.PdfPageHeightPoint : pageHeight;

            var blankPage = new PdfPageViewModel
            {
                OriginalFilePath = SelectedDocument.FilePath,
                OriginalPageIndex = -1,
                IsBlankPage = true,
                PageIndex = insertIndex,
                Width = pageWidth,
                Height = pageHeight,
                PdfPageWidthPoint = pdfWidth,
                PdfPageHeightPoint = pdfHeight,
                CropWidthPoint = pdfWidth,
                CropHeightPoint = pdfHeight
            };

            SelectedDocument.Pages.Insert(insertIndex, blankPage);
            RefreshPageIndexes();
            TxtPageInfo.Text = $"{insertIndex + 1} / {SelectedDocument.Pages.Count}";
            ScrollToPage(insertIndex);
        }
        private void BtnDeletePage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0) return;

            int currentPageIndex = GetCurrentPageIndex();
            if (currentPageIndex < 0 || currentPageIndex >= SelectedDocument.Pages.Count) return;

            var result = MessageBox.Show($"{currentPageIndex + 1}페이지를 삭제하시겠습니까?", "페이지 삭제", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                var pageVM = SelectedDocument.Pages[currentPageIndex];
                
                // 리소스 정리
                pageVM.Unload();
                
                // 컬렉션에서 제거
                SelectedDocument.Pages.RemoveAt(currentPageIndex);
                
                RefreshPageIndexes();
                
                // 상태바 갱신
                TxtPageInfo.Text = $"{Math.Min(currentPageIndex + 1, SelectedDocument.Pages.Count)} / {SelectedDocument.Pages.Count}";
            }
        }
        private void BtnRotateLeft_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) foreach (var p in SelectedDocument.Pages) p.Rotation = (p.Rotation + 270) % 360; }
        private void BtnRotateRight_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) foreach (var p in SelectedDocument.Pages) p.Rotation = (p.Rotation + 90) % 360; }

        private void RefreshPageIndexes()
        {
            if (SelectedDocument == null) return;
            for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                SelectedDocument.Pages[i].PageIndex = i;
        }

        private Canvas? FindPageCanvas(PdfPageViewModel page)
        {
            var listView = GetVisualChild<ListView>(MainTabControl);
            if (listView == null) return null;
            var container = listView.ItemContainerGenerator.ContainerFromItem(page) as DependencyObject;
            if (container != null) return FindChild<Canvas>(container, "AnnotationCanvas");
            return null;
        }

        private T? FindChild<T>(DependencyObject parent, string childName) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T t && (child as FrameworkElement)?.Name == childName) return t;
                var found = FindChild<T>(child, childName);
                if (found != null) return found;
            }
            return null;
        }
    }
}
