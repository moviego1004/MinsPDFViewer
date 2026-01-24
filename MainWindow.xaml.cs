using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using PdfiumViewer;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace MinsPDFViewer
{
    public partial class MainWindow : Window, System.ComponentModel.INotifyPropertyChanged
    {
        private readonly PdfService _pdfService;
        private readonly SearchService _searchService;
        private readonly PdfSignatureService _signatureService;
        private OcrEngine? _ocrEngine;
        private readonly HistoryService _historyService;

        public System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel> Documents
        {
            get; set;
        }
            = new System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel>();

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

        private bool _isSpacePressed = false;
        private bool _isPanning = false;
        private Point _lastPanPoint;

        private ScrollViewer? _cachedScrollViewer;

        public MainWindow()
        {
            // Register encoding provider for proper PDF text handling
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            InitializeComponent();
            DataContext = this;
            _pdfService = new PdfService();
            _searchService = new SearchService();
            _signatureService = new PdfSignatureService();
            _historyService = new HistoryService();

            // Force bind events to ensure they work even if XAML binding fails
            BtnSave.Click += BtnSave_Click;
            BtnSaveAs.Click += BtnSaveAs_Click;

            // Setup global logging and exception handling
            Log("=== Application Starting ===");
            AppDomain.CurrentDomain.UnhandledException += (s, e) => Log($"[CRITICAL] UnhandledException: {e.ExceptionObject}");
            this.Dispatcher.UnhandledException += (s, e) => {
                Log($"[CRITICAL] DispatcherUnhandledException: {e.Exception}");
            };

            try
            {
                if (PdfSharp.Fonts.GlobalFontSettings.FontResolver == null)
                    PdfSharp.Fonts.GlobalFontSettings.FontResolver = new WindowsFontResolver();
            }
            catch { }

            try
            {
                _ocrEngine = OcrEngine.TryCreateFromLanguage(new Language("ko-KR")) ?? OcrEngine.TryCreateFromUserProfileLanguages();
            }
            catch { }

            CbFont.ItemsSource = new string[] { "Malgun Gothic", "Gulim", "Dotum", "Batang" };
            CbFont.SelectedIndex = 0;
            CbSize.ItemsSource = new double[] { 10, 12, 14, 16, 18, 24, 32, 48 };
            CbSize.SelectedIndex = 2;
        }

        private void Log(string message)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                string logMsg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}{Environment.NewLine}";
                File.AppendAllText(logPath, logMsg);
                System.Diagnostics.Debug.Write(logMsg);
                Console.Write(logMsg);
            }
            catch { }
        }

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_selectedAnnotation != null)
            {
                _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = null;
                CheckToolbarVisibility();
            }
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
                // Ensure we don't have multiple subscriptions
                pageVM.Annotations.CollectionChanged -= (s, args) => RefreshCanvas(canvas, pageVM);
                pageVM.Annotations.CollectionChanged += (s, args) => RefreshCanvas(canvas, pageVM);
                
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
                AnnotationType.SignaturePlaceholder => "SignaturePlaceholderTemplate",
                AnnotationType.SignatureField => "SignatureFieldTemplate",
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
                    Task.Run(() => _pdfService.RenderPageImage(SelectedDocument, pageVM));
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
                Log($"[DEBUG] Page_MouseDown called, _currentTool = {_currentTool}");

                // 1. Check if we clicked on an annotation (including TextBox)
                if (IsAnnotationObject(e.OriginalSource))
                {
                    Log("[DEBUG] Page_MouseDown: Clicked on Annotation - Ignoring");
                    return;
                }

                // 2. Only if NOT clicking an annotation, clear focus from TextBox
                if (Keyboard.FocusedElement is TextBox)
                {
                    Log("[DEBUG] Page_MouseDown: Clearing Focus");
                    Keyboard.ClearFocus();
                }

                if (SelectedDocument == null) return;
                var grid = sender as Grid ?? FindAncestor<Grid>(e.OriginalSource as DependencyObject);
                if (grid == null) return;
                var pageVM = grid.DataContext as PdfPageViewModel;
                if (pageVM == null) return;
                _activePageIndex = pageVM.PageIndex;
                _dragStartPoint = e.GetPosition(grid);

                if (_currentTool == "TEXT")
                {
                    if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                    var newAnnot = new PdfAnnotation { Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50, FontSize = _defaultFontSize, FontFamily = _defaultFontFamily, Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold, TextContent = "", IsSelected = true };
                    pageVM.Annotations.Add(newAnnot);
                    var canvas = FindChild<Canvas>(grid, "AnnotationCanvas");
                    if (canvas != null) RefreshCanvas(canvas, pageVM);
                    _selectedAnnotation = newAnnot; _currentTool = "CURSOR"; RbCursor.IsChecked = true;
                    UpdateToolbarFromAnnotation(_selectedAnnotation); CheckToolbarVisibility(); e.Handled = true; return;
                }

                if (_currentTool == "HIGHLIGHT")
                {
                    if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                    foreach (var p in SelectedDocument.Pages) { p.IsSelecting = false; p.IsHighlighting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                    SelectionPopup.IsOpen = false; pageVM.IsSelecting = true; pageVM.IsHighlighting = true; pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y; grid.CaptureMouse();
                    e.Handled = true; return;
                }
                
                // Moved IsAnnotationObject check to top

                if (_selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); }
                if (_currentTool == "CURSOR")
                {
                    foreach (var p in SelectedDocument.Pages) { p.IsSelecting = false; p.IsHighlighting = false; p.SelectionWidth = 0; p.SelectionHeight = 0; }
                    SelectionPopup.IsOpen = false; pageVM.IsSelecting = true; pageVM.IsHighlighting = false; pageVM.SelectionX = _dragStartPoint.X; pageVM.SelectionY = _dragStartPoint.Y; grid.CaptureMouse();
                    e.Handled = true;
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

        public void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1 || SelectedDocument == null)
                return;
            
            var pageVM = (sender as FrameworkElement)?.DataContext as PdfPageViewModel;
            if (pageVM == null) return;

            var relativeTo = sender as IInputElement;

            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null)
            {
                var currentPoint = e.GetPosition(relativeTo);
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
                    CheckTextInSelection(_activePageIndex, rect);
                    SelectionPopup.PlacementTarget = sender as UIElement;
                    SelectionPopup.PlacementRectangle = new Rect(e.GetPosition(sender as IInputElement).X, e.GetPosition(sender as IInputElement).Y + 10, 0, 0);
                    SelectionPopup.IsOpen = true;
                    _selectedPageIndex = _activePageIndex;
                    TxtStatus.Text = string.IsNullOrEmpty(_selectedTextBuffer) ? "영역 선택됨" : "텍스트 선택됨";
                }
                else
                {
                    pageVM.IsSelecting = false;
                    TxtStatus.Text = "준비";
                }
            }
            _activePageIndex = -1;
            _isDraggingAnnotation = false;
            e.Handled = true;
            CheckToolbarVisibility();
        }

        private async void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            _selectedTextBuffer = "";
            if (SelectedDocument?.PdfDocument == null)
                return;

            var doc = SelectedDocument;

            string extractedText = await Task.Run(() =>
            {
                try
                {
                    lock (PdfService.PdfiumLock)
                    {
                        if (doc.PdfDocument == null) return "";
                        return ""; // Simplified for now
                    }
                }
                catch { return ""; }
            });

            _selectedTextBuffer = extractedText;
            if (!string.IsNullOrEmpty(_selectedTextBuffer))
            {
                TxtStatus.Text = "텍스트 선택됨";
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
                    _annotationDragStartOffset = container != null ? e.GetPosition(container) : e.GetPosition(border);
                    parentGrid.CaptureMouse();
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

                        XRect pdfRect = new XRect(pdfX, pdfY, ann.Width * scaleX, ann.Height * scaleY);
                        string tempPath = Path.GetTempFileName();
                        await _pdfService.SavePdf(SelectedDocument, tempPath);
                        _signatureService.SignPdf(tempPath, saveDlg.FileName, config, pageIndex, pdfRect);
                        if (File.Exists(tempPath)) File.Delete(tempPath);
                        targetPage.Annotations.Remove(ann);
                        MessageBox.Show($"전자서명 완료!\n저장됨: {saveDlg.FileName}");
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
            Log($"[DEBUG] AnnotationTextBox_Loaded Fired. DataContext Type: {(dc != null ? dc.GetType().Name : "null")}");

            if (dc is PdfAnnotation ann)
            {
                // [Robustness] Force subscribe to ensure event fires
                tb.TextChanged -= AnnotationTextBox_TextChanged;
                tb.TextChanged += AnnotationTextBox_TextChanged;
                
                // [Sync] Ensure initial text is correct
                if (ann.TextContent != tb.Text) tb.Text = ann.TextContent;

                if (ann.IsSelected) tb.Focus();
                Log($"[DEBUG] AnnotationTextBox_Loaded: TextBox loaded for '{ann.TextContent}'");
            }
            else
            {
                Log($"[CRITICAL] AnnotationTextBox_Loaded: DataContext is NOT PdfAnnotation! It is {dc}");
            }
        }

        private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Log("[DEBUG] AnnotationTextBox_TextChanged Fired (Raw)");
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                // [Direct Sync] Always push text to model, bypassing binding if necessary
                ann.TextContent = tb.Text;
                Log($"[DEBUG] TextChanged: Model updated to '{ann.TextContent}'");

                var formattedText = new FormattedText(tb.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface(new System.Windows.Media.FontFamily(ann.FontFamily), FontStyles.Normal, ann.IsBold ? FontWeights.Bold : FontWeights.Normal, FontStretches.Normal),
                    ann.FontSize, Brushes.Black, VisualTreeHelper.GetDpi(this).PixelsPerDip);

                double newWidth = Math.Max(100, formattedText.Width + 25);
                double newHeight = Math.Max(50, formattedText.Height + 30);

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
                    if (_selectedAnnotation != null && _selectedAnnotation != ann) _selectedAnnotation.IsSelected = false;
                    _selectedAnnotation = ann;
                    _selectedAnnotation.IsSelected = true;
                    UpdateToolbarFromAnnotation(ann);
                    CheckToolbarVisibility();
                }
            }
        }

        private void AnnotationTextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                var newFocus = Keyboard.FocusedElement as DependencyObject;
                if (newFocus != null)
                {
                    var newAnn = FindAncestorDataContext<PdfAnnotation>(newFocus);
                    if (newAnn == ann) return;
                }
                
                if (ann.IsSelected)
                {
                    ann.IsSelected = false;
                    CheckToolbarVisibility();
                }
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
                if (_selectedAnnotation != null) _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = ann;
                _selectedAnnotation.IsSelected = true;
                UpdateToolbarFromAnnotation(ann);
                
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
                    _isDraggingAnnotation = true;
                    parentGrid.CaptureMouse();
                    e.Handled = true;
                }
            }
        }

        private async void BtnOCR_Click(object sender, RoutedEventArgs e)
        {
            if (_ocrEngine == null) { MessageBox.Show("OCR 미지원"); return; }
            if (SelectedDocument == null) return;
            BtnOCR.IsEnabled = false;
            PbStatus.Visibility = Visibility.Visible;
            PbStatus.Maximum = SelectedDocument.Pages.Count;
            PbStatus.Value = 0;
            TxtStatus.Text = "OCR 분석 중...";
            try
            {
                await Task.Run(async () =>
                {
                    for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                    {
                        if (SelectedDocument.IsDisposed) break;
                        var pageVM = SelectedDocument.Pages[i];
                        byte[]? bitmapBytes = null;
                        int bmpWidth = 0, bmpHeight = 0;
                        lock (PdfService.PdfiumLock)
                        {
                            if (SelectedDocument.PdfDocument == null) continue;
                            using (var bitmap = SelectedDocument.PdfDocument.Render(i, (int)pageVM.Width, (int)pageVM.Height, 96, 96, PdfRenderFlags.None))
                            using (var ms = new MemoryStream())
                            {
                                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                bitmapBytes = ms.ToArray();
                                bmpWidth = bitmap.Width; bmpHeight = bitmap.Height;
                            }
                        }
                        if (bitmapBytes != null)
                        {
                            var ibuffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(bitmapBytes);
                            using (var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, bmpWidth, bmpHeight, BitmapAlphaMode.Premultiplied))
                            {
                                softwareBitmap.CopyFromBuffer(ibuffer);
                                var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap);
                                var wordList = new List<OcrWordInfo>();
                                foreach (var line in ocrResult.Lines)
                                    foreach (var word in line.Words)
                                        wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(word.BoundingRect.X, word.BoundingRect.Y, word.BoundingRect.Width, word.BoundingRect.Height) });
                                pageVM.OcrWords = wordList;
                            }
                        }
                        await Application.Current.Dispatcher.InvokeAsync(() => { PbStatus.Value = i + 1; });
                    }
                });
                TxtStatus.Text = "OCR 완료.";
            }
            catch (Exception ex) { Log($"[CRITICAL] OCR Error: {ex}"); TxtStatus.Text = "오류 발생"; }
            finally { BtnOCR.IsEnabled = true; PbStatus.Visibility = Visibility.Collapsed; }
        }

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            // [Fix] Force update bindings by clearing focus (handling IME composition/active TextBox)
            Keyboard.ClearFocus();
            FocusManager.SetFocusedElement(this, this);

            if (SelectedDocument == null) return;

            // [Audit Log] Check memory state before save
            foreach (var p in SelectedDocument.Pages)
            {
                foreach (var a in p.Annotations)
                {
                    if (a.Type == AnnotationType.FreeText)
                    {
                        Log($"[DEBUG] Pre-Save Audit: Page={p.PageIndex}, Text='{a.TextContent}', Rect=({a.X},{a.Y},{a.Width},{a.Height})");
                    }
                }
            }

            // Capture the document to save in a local variable to prevent NRE if selection changes
            var docToSave = SelectedDocument;

            string originalPath = docToSave.FilePath;
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".pdf");
            int currentPageIndex = GetCurrentPageIndex();
            var sv = GetCurrentScrollViewer();
            double vOffset = sv?.VerticalOffset ?? 0;
            double hOffset = sv?.HorizontalOffset ?? 0;

            try
            {
                TxtStatus.Text = "저장 중...";
                await _pdfService.SavePdf(docToSave, tempPath);
                
                docToSave.Dispose();
                Documents.Remove(docToSave);
                if (SelectedDocument == docToSave) SelectedDocument = null;
                
                await Task.Run(() => File.Copy(tempPath, originalPath, true));
                
                _historyService.SetLastPage(originalPath, currentPageIndex);
                _historyService.SaveHistory();
                
                await OpenPdfFromPath(originalPath);
                
                // Allow UI to update bindings
                await Dispatcher.InvokeAsync(() => { }, System.Windows.Threading.DispatcherPriority.Loaded);

                var newSv = GetCurrentScrollViewer();
                if (newSv != null)
                {
                    newSv.ScrollToVerticalOffset(vOffset);
                    newSv.ScrollToHorizontalOffset(hOffset);
                }
                
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
                if (File.Exists(tempPath)) try { File.Delete(tempPath); } catch { } 
            } 
        }

        private async void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null) return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_annotated" };
            if (dlg.ShowDialog() == true)
            {
                try { 
                    await _pdfService.SavePdf(SelectedDocument, dlg.FileName);
                    MessageBox.Show("저장되었습니다.");
                } catch (Exception ex) { Log($"[CRITICAL] SaveAs Error: {ex}"); MessageBox.Show($"저장 실패: {ex.Message}"); }
            }
        }

        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e) { Clipboard.SetText(_selectedTextBuffer); SelectionPopup.IsOpen = false; }
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

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom += 0.1; }
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Zoom -= 0.1; }
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
        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e) {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control && SelectedDocument != null) {
                if (e.Delta > 0) SelectedDocument.Zoom += 0.1; else SelectedDocument.Zoom -= 0.1;
                e.Handled = true;
            }
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
            bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null);
            if (TextStyleToolbar != null) TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            if ((sender as Thumb)?.DataContext is PdfAnnotation ann) {
                ann.Width = Math.Max(50, ann.Width + e.HorizontalChange);
                ann.Height = Math.Max(30, ann.Height + e.VerticalChange);
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
                    else if (content.Contains("텍스트")) _currentTool = "TEXT";
                    CheckToolbarVisibility();
                }
            }
            catch (Exception ex)
            {
                Log($"[CRITICAL] Tool_Checked Error: {ex}");
            }
        }

        private void ListView_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            // Log($"[DEBUG] ListView_PreviewMouseDown: OriginalSource={e.OriginalSource}, Type={e.OriginalSource?.GetType().Name}");
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control) { TxtSearch.Focus(); TxtSearch.SelectAll(); e.Handled = true; }
            else if (e.Key == Key.Delete) BtnDeleteAnnotation_Click(this, new RoutedEventArgs());
            else if (e.Key == Key.Escape && _selectedAnnotation != null) { _selectedAnnotation.IsSelected = false; _selectedAnnotation = null; CheckToolbarVisibility(); }
            if (e.Key == Key.Space && !_isSpacePressed && _currentTool != "TEXT" && !(Keyboard.FocusedElement is TextBox)) { _isSpacePressed = true; Mouse.OverrideCursor = Cursors.Hand; e.Handled = true; }
            else if (e.Key == Key.PageUp || e.Key == Key.PageDown) { var sv = GetScrollViewer(); if (sv != null) { if (e.Key == Key.PageDown) sv.PageDown(); else sv.PageUp(); e.Handled = true; } }
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
                    using (var doc = PdfReader.Open(SelectedDocument.FilePath, PdfDocumentOpenMode.Import)) {
                        PdfDictionary? dict = null;
                        foreach (var page in doc.Pages) {
                            if (page.Annotations != null) {
                                for (int i=0; i<page.Annotations.Count; i++) {
                                    var pa = page.Annotations[i] as PdfDictionary;
                                    if (pa != null && pa.Elements.GetString("/Subtype") == "/Widget" && pa.Elements.GetString("/FT") == "/Sig" && pa.Elements.GetString("/T") == ann.FieldName) {
                                        dict = pa.Elements.GetDictionary("/V"); break;
                                    }
                                }
                            }
                            if (dict != null) break;
                        }
                        if (dict != null) {
                            var res = new SignatureVerificationService().VerifySignature(SelectedDocument.FilePath, dict);
                            new SignatureResultWindow(res) { Owner = this }.ShowDialog();
                        }
                    }
                } catch (Exception ex) { MessageBox.Show($"검증 오류: {ex.Message}"); }
            }
        }

        private void BtnToggleSidebar_Click(object sender, RoutedEventArgs e) => SidebarBorder.Visibility = SidebarBorder.Visibility == Visibility.Visible ? Visibility.Collapsed : Visibility.Visible;
        private void BookmarkTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e) {
            if (e.NewValue is PdfBookmarkViewModel bm && SelectedDocument != null) {
                var target = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex);
                if (target != null) GetCurrentListView()?.ScrollIntoView(target);
            }
        }

        private Point _startPoint; private bool _isDragging = false;
        private void BookmarkTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) => _startPoint = e.GetPosition(null);
        private void BookmarkTree_PreviewMouseMove(object sender, MouseEventArgs e) {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging) {
                Point pos = e.GetPosition(null); Vector diff = _startPoint - pos;
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) {
                    if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm && !bm.IsEditing) {
                        _isDragging = true; DragDrop.DoDragDrop(BookmarkTree, new DataObject("MinsBookmark", bm), DragDropEffects.Move); _isDragging = false;
                    }
                }
            }
        }
        private void BookmarkTree_DragOver(object sender, DragEventArgs e) { e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : e.Data.GetDataPresent("MinsBookmark") ? DragDropEffects.Move : DragDropEffects.None; e.Handled = true; }
        private void BookmarkTree_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) { Window_Drop(sender, e); return; }
            if (!e.Data.GetDataPresent("MinsBookmark") || SelectedDocument == null) return;
            var src = e.Data.GetData("MinsBookmark") as PdfBookmarkViewModel;
            var target = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource)?.DataContext as PdfBookmarkViewModel;
            if (src == null || (target != null && (target == src || IsChildOf(target, src)))) return;
            GetCurrentCollection(src)?.Remove(src);
            if (target == null) { SelectedDocument.Bookmarks.Add(src); src.Parent = null; }
            else { target.Children.Add(src); src.Parent = target; target.IsExpanded = true; }
            e.Handled = true;
        }
        private bool IsChildOf(PdfBookmarkViewModel c, PdfBookmarkViewModel p) { var cur = c.Parent; while (cur != null) { if (cur == p) return true; cur = cur.Parent; } return false; }
        private static T? FindAncestor<T>(DependencyObject cur) where T : DependencyObject { while (cur != null) { if (cur is T t) return t; cur = VisualTreeHelper.GetParent(cur); } return null; }
        private void BookmarkTree_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.F2 && BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) bm.IsEditing = true; }
        private void BtnRenameBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) bm.IsEditing = true; }
        private void BookmarkRename_KeyDown(object sender, KeyEventArgs e) { if (e.Key == Key.Enter && (sender as TextBox)?.DataContext is PdfBookmarkViewModel bm) { bm.IsEditing = false; e.Handled = true; } }
        private void BookmarkRename_LostFocus(object sender, RoutedEventArgs e) { if ((sender as TextBox)?.DataContext is PdfBookmarkViewModel bm) bm.IsEditing = false; }
        private void PdfListView_ScrollChanged(object sender, ScrollChangedEventArgs e) { if (SelectedDocument != null) TxtPageInfo.Text = $"{GetCurrentPageIndex() + 1} / {SelectedDocument.Pages.Count}"; }
        private void BtnAddBookmark_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) SelectedDocument.Bookmarks.Add(new PdfBookmarkViewModel { Title = $"Page {GetCurrentPageIndex() + 1}", PageIndex = GetCurrentPageIndex() }); }
        private void BtnUpdateBookmarkPage_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { int idx = GetCurrentPageIndex(); bm.PageIndex = idx; if (bm.Title.StartsWith("새 책갈피")) bm.Title = $"새 책갈피 (p.{idx + 1})"; } }
        private void BookmarkItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e) { if ((sender as FrameworkElement)?.DataContext is PdfBookmarkViewModel bm && SelectedDocument != null) { var t = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex); if (t != null) GetCurrentListView()?.ScrollIntoView(t); } }
        private void BtnDeleteBookmark_Click(object sender, RoutedEventArgs e) { var t = (sender as MenuItem)?.DataContext as PdfBookmarkViewModel ?? BookmarkTree.SelectedItem as PdfBookmarkViewModel; if (t != null && SelectedDocument != null) { if (SelectedDocument.Bookmarks.Contains(t)) SelectedDocument.Bookmarks.Remove(t); else t.Parent?.Children.Remove(t); } }
        private void BtnMoveBookmarkUp_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i > 0) col?.Move(i, i - 1); } }
        private void BtnMoveBookmarkDown_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i >= 0 && i < col.Count - 1) col?.Move(i, i + 1); } }
        private void BtnIndentBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm) { var col = GetCurrentCollection(bm); int i = col?.IndexOf(bm) ?? -1; if (i > 0) { var p = col[i - 1]; col.Remove(bm); p.Children.Add(bm); bm.Parent = p; p.IsExpanded = true; } } }
        private void BtnOutdentBookmark_Click(object sender, RoutedEventArgs e) { if (BookmarkTree.SelectedItem is PdfBookmarkViewModel bm && bm.Parent != null) { var oldP = bm.Parent; var list = (oldP.Parent == null) ? SelectedDocument.Bookmarks : oldP.Parent.Children; int i = list.IndexOf(oldP); if (i >= 0) { oldP.Children.Remove(bm); list.Insert(i + 1, bm); bm.Parent = oldP.Parent; } } }
        private System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>? GetCurrentCollection(PdfBookmarkViewModel bm) => SelectedDocument == null ? null : bm.Parent == null ? SelectedDocument.Bookmarks : bm.Parent.Children;

        private int GetCurrentPageIndex()
        {
            if (SelectedDocument == null) return 0;
            var sv = GetCurrentScrollViewer();
            if (sv == null) return 0;
            double off = sv.VerticalOffset; double acc = 0;
            foreach (var p in SelectedDocument.Pages) {
                double h = (p.Height * SelectedDocument.Zoom) + 20;
                if (acc + h > off + 50) return p.PageIndex;
                acc += h;
            }
            return 0;
        }

        public async Task OpenPdfFromPath(string path)
        {
            if (!File.Exists(path)) return;
            Log($"[DEBUG] OpenPdfFromPath: {path}");
            var existing = Documents.FirstOrDefault(d => d.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (existing != null) { SelectedDocument = existing; return; }
            var doc = await _pdfService.LoadPdfAsync(path);
            if (doc != null) { await _pdfService.InitializeDocumentAsync(doc); Documents.Add(doc); SelectedDocument = doc; }
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
        private void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e) { } // No-op, handled by PdfListView_PreviewMouseWheel
        private void BtnAddBlankPage_Click(object sender, RoutedEventArgs e) => MessageBox.Show("준비 중");
        private void BtnDeletePage_Click(object sender, RoutedEventArgs e) => MessageBox.Show("준비 중");
        private void BtnRotateLeft_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) foreach (var p in SelectedDocument.Pages) p.Rotation = (p.Rotation + 270) % 360; }
        private void BtnRotateRight_Click(object sender, RoutedEventArgs e) { if (SelectedDocument != null) foreach (var p in SelectedDocument.Pages) p.Rotation = (p.Rotation + 90) % 360; }

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