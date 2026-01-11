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
            InitializeComponent();
            DataContext = this;
            _pdfService = new PdfService();
            _searchService = new SearchService();
            _signatureService = new PdfSignatureService();
            _historyService = new HistoryService();

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

        private void Page_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (IsAnnotationObject(e.OriginalSource))
                return;

            if (_selectedAnnotation != null)
            {
                _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = null;
                CheckToolbarVisibility();
            }

            if (SelectedDocument == null)
                return;

            var grid = sender as Grid;
            if (grid == null)
                return;

            _activePageIndex = (int)grid.Tag;
            _dragStartPoint = e.GetPosition(grid);
            var pageVM = SelectedDocument.Pages[_activePageIndex];

            if (_currentTool == "TEXT")
            {
                var newAnnot = new PdfAnnotation { Type = AnnotationType.FreeText, X = _dragStartPoint.X, Y = _dragStartPoint.Y, Width = 150, Height = 50, FontSize = _defaultFontSize, FontFamily = _defaultFontFamily, Foreground = new SolidColorBrush(_defaultFontColor), IsBold = _defaultIsBold, TextContent = "", IsSelected = true };
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
                foreach (var p in SelectedDocument.Pages)
                {
                    p.IsSelecting = false;
                    p.SelectionWidth = 0;
                    p.SelectionHeight = 0;
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

        private bool IsAnnotationObject(object source)
        {
            if (source is DependencyObject dep)
            {
                DependencyObject current = dep;
                while (current != null)
                {
                    if (current is FrameworkElement fe && fe.DataContext is PdfAnnotation)
                        return true;
                    if (current is Grid g && g.DataContext is PdfPageViewModel)
                        break;
                    current = VisualTreeHelper.GetParent(current);
                }
            }
            return false;
        }

        private void Page_MouseMove(object sender, MouseEventArgs e)
        {
            if (_activePageIndex == -1 || SelectedDocument == null)
                return;
            var grid = sender as Grid;
            if (grid == null)
                return;
            var pageVM = SelectedDocument.Pages[_activePageIndex];
            if (_currentTool == "CURSOR" && _isDraggingAnnotation && _selectedAnnotation != null)
            {
                var currentPoint = e.GetPosition(grid);
                double newX = currentPoint.X - _annotationDragStartOffset.X;
                double newY = currentPoint.Y - _annotationDragStartOffset.Y;
                if (newX < 0)
                    newX = 0;
                if (newY < 0)
                    newY = 0;
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
            if (grid == null || _activePageIndex == -1 || SelectedDocument == null)
                return;
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
                        if (doc.PdfDocument == null)
                            return "";
                        // [임시 처리] PdfiumViewer에서는 좌표 기반 텍스트 추출이 복잡하므로 
                        // 빌드 에러 방지를 위해 우선 빈 문자열 반환 혹은 추후 구현
                        return "";
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
                if (SelectedDocument == null || _selectedPageIndex == -1)
                    return;

                var pageVM = SelectedDocument.Pages[_selectedPageIndex];
                if (pageVM.SelectionWidth <= 0 || pageVM.SelectionHeight <= 0)
                    return;

                TxtStatus.Text = "이미지 처리 중...";
                BitmapSource? croppedBitmap = null;

                await Task.Run(() =>
                {
                    lock (PdfService.PdfiumLock)
                    {
                        if (SelectedDocument.PdfDocument == null)
                            return;

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
                                var cropRect = new System.Drawing.Rectangle(cx, cy, cw, ch);
                                if (cx < 0)
                                    cx = 0;
                                if (cy < 0)
                                    cy = 0;
                                if (cx + cw > fullBmp.Width)
                                    cw = fullBmp.Width - cx;
                                if (cy + ch > fullBmp.Height)
                                    ch = fullBmp.Height - cy;

                                // [수정] 명시적 Bitmap 캐스팅 및 using 구문 수정
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
            if (SelectedDocument == null)
                return;

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

            if (pageIndex == -1 || targetPage == null)
                return;

            var dlg = new CertificateWindow();
            dlg.Owner = this;

            if (dlg.ShowDialog() == true && dlg.ResultConfig != null)
            {
                var config = dlg.ResultConfig;
                config.VisualStampPath = ann.VisualStampPath;
                config.UseVisualStamp = true;

                var saveDlg = new SaveFileDialog
                {
                    Filter = "PDF Files|*.pdf",
                    FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_signed.pdf",
                    Title = "서명된 파일 저장"
                };

                if (saveDlg.ShowDialog() == true)
                {
                    try
                    {
                        double effectivePdfWidth = (targetPage.CropWidthPoint > 0) ? targetPage.CropWidthPoint : targetPage.PdfPageWidthPoint;
                        double effectivePdfHeight = (targetPage.CropHeightPoint > 0) ? targetPage.CropHeightPoint : targetPage.PdfPageHeightPoint;

                        double pdfOriginX = targetPage.CropX;
                        double pdfOriginY = targetPage.CropY;

                        double scaleX = effectivePdfWidth / targetPage.Width;
                        double scaleY = effectivePdfHeight / targetPage.Height;

                        double pdfX = pdfOriginX + (ann.X * scaleX);
                        double pdfY = (pdfOriginY + effectivePdfHeight) - ((ann.Y + ann.Height) * scaleY);

                        XRect pdfRect = new XRect(pdfX, pdfY, ann.Width * scaleX, ann.Height * scaleY);

                        string tempPath = Path.GetTempFileName();
                        await _pdfService.SavePdf(SelectedDocument, tempPath);

                        _signatureService.SignPdf(tempPath, saveDlg.FileName, config, pageIndex, pdfRect);

                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                        targetPage.Annotations.Remove(ann);
                        MessageBox.Show($"전자서명 완료!\n저장됨: {saveDlg.FileName}");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"서명 실패: {ex.Message}");
                    }
                }
            }
        }

        private FrameworkElement? GetParentContentPresenter(DependencyObject child)
        {
            DependencyObject parent = child;
            while (parent != null)
            {
                if (parent is ContentPresenter cp)
                    return cp;
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
                if (ann.IsSelected)
                    tb.Focus();
            }
        }

        private void AnnotationTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox tb && tb.DataContext is PdfAnnotation ann)
            {
                var formattedText = new FormattedText(tb.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
                    new Typeface(new System.Windows.Media.FontFamily(ann.FontFamily),
                                 ann.IsBold ? FontStyles.Normal : FontStyles.Normal,
                                 ann.IsBold ? FontWeights.Bold : FontWeights.Normal,
                                 FontStretches.Normal),
                    ann.FontSize, Brushes.Black, VisualTreeHelper.GetDpi(this).PixelsPerDip);

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
            if (element?.DataContext is PdfAnnotation ann && ann.Type != AnnotationType.FreeText && ann.Type != AnnotationType.SignaturePlaceholder)
            {
                if (_currentTool != "CURSOR")
                {
                    _currentTool = "CURSOR";
                    RbCursor.IsChecked = true;
                    CheckToolbarVisibility();
                }
                if (_selectedAnnotation != null)
                    _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = ann;
                _selectedAnnotation.IsSelected = true;
                UpdateToolbarFromAnnotation(ann);
                DependencyObject parent = element;
                Grid? parentGrid = null;
                while (parent != null)
                {
                    if (parent is Grid g && g.Tag is int idx)
                    {
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
            if (_ocrEngine == null)
            {
                MessageBox.Show("OCR 미지원 (Windows 10/11 기능 필요)");
                return;
            }
            if (SelectedDocument == null || SelectedDocument.Pages.Count == 0)
                return;
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
                        lock (PdfService.PdfiumLock)
                        {
                            if (targetDoc.PdfDocument == null)
                                continue;
                            using (var bitmap = targetDoc.PdfDocument.Render(i, (int)pageVM.Width, (int)pageVM.Height, 96, 96, PdfRenderFlags.None))
                            using (var ms = new MemoryStream())
                            {
                                bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                var bytes = ms.ToArray();
                                var ibuffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(bytes);
                                using (var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, bitmap.Width, bitmap.Height, BitmapAlphaMode.Premultiplied))
                                {
                                    softwareBitmap.CopyFromBuffer(ibuffer);
                                    var ocrResult = _ocrEngine.RecognizeAsync(softwareBitmap).GetAwaiter().GetResult();
                                    var wordList = new List<OcrWordInfo>();
                                    foreach (var line in ocrResult.Lines)
                                    {
                                        foreach (var word in line.Words)
                                        {
                                            wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(word.BoundingRect.X, word.BoundingRect.Y, word.BoundingRect.Width, word.BoundingRect.Height) });
                                        }
                                    }
                                    pageVM.OcrWords = wordList;
                                }
                            }
                        }
                        Application.Current.Dispatcher.Invoke(() => { PbStatus.Value = i + 1; });
                    }
                });
                TxtStatus.Text = "OCR 완료.";
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

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null)
            {
                try
                {
                    await _pdfService.SavePdf(SelectedDocument, SelectedDocument.FilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"저장 오류: {ex.Message}");
                }
            }
        }

        private async void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_annotated" };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    await _pdfService.SavePdf(SelectedDocument, dlg.FileName);
                    SelectedDocument.FilePath = dlg.FileName;
                    SelectedDocument.FileName = Path.GetFileName(dlg.FileName);
                    MessageBox.Show("저장되었습니다.");
                }
                catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
            }
        }

        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(_selectedTextBuffer);
            SelectionPopup.IsOpen = false;
        }

        private void BtnPopupHighlightGreen_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Lime, AnnotationType.Highlight);
        private void BtnPopupHighlightOrange_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Orange, AnnotationType.Highlight);
        private void BtnPopupUnderline_Click(object sender, RoutedEventArgs e) => AddAnnotation(Colors.Black, AnnotationType.Underline);
        private void AddAnnotation(Color color, AnnotationType type)
        {
            if (_selectedPageIndex == -1 || SelectedDocument == null)
                return;
            var p = SelectedDocument.Pages[_selectedPageIndex];
            var ann = new PdfAnnotation { X = p.SelectionX, Y = p.SelectionY, Width = p.SelectionWidth, Height = p.SelectionHeight, Type = type, AnnotationColor = color };
            if (type == AnnotationType.Highlight)
                ann.Background = new SolidColorBrush(Color.FromArgb(80, color.R, color.G, color.B));
            else
            {
                ann.Background = new SolidColorBrush(color);
                ann.Height = 2;
                ann.Y = p.SelectionY + p.SelectionHeight - 2;
            }
            p.Annotations.Add(ann);
            SelectionPopup.IsOpen = false;
            p.IsSelecting = false;
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null)
                SelectedDocument.Zoom += 0.1;
        }
        private void BtnZoomOut_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null)
                SelectedDocument.Zoom -= 0.1;
        }
        private void BtnFitWidth_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null && SelectedDocument.Pages.Count > 0)
            {
                double viewWidth = MainTabControl.ActualWidth - 60;
                if (viewWidth > 0 && SelectedDocument.Pages[0].Width > 0)
                    SelectedDocument.Zoom = viewWidth / SelectedDocument.Pages[0].Width;
            }
        }
        private void BtnFitHeight_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null && SelectedDocument.Pages.Count > 0)
            {
                double viewHeight = MainTabControl.ActualHeight - 60;
                if (viewHeight > 0 && SelectedDocument.Pages[0].Height > 0)
                    SelectedDocument.Zoom = viewHeight / SelectedDocument.Pages[0].Height;
            }
        }
        private void PdfListView_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (SelectedDocument != null)
                {
                    if (e.Delta > 0)
                        SelectedDocument.Zoom += 0.1;
                    else
                        SelectedDocument.Zoom -= 0.1;
                }
                e.Handled = true;
            }
        }

        private void BtnDeleteAnnotation_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedAnnotation != null && SelectedDocument != null)
            {
                foreach (var p in SelectedDocument.Pages)
                {
                    if (p.Annotations.Contains(_selectedAnnotation))
                    {
                        p.Annotations.Remove(_selectedAnnotation);
                        _selectedAnnotation = null;
                        CheckToolbarVisibility();
                        break;
                    }
                }
            }
            else if ((sender as MenuItem)?.CommandParameter is PdfAnnotation a && SelectedDocument != null)
            {
                foreach (var p in SelectedDocument.Pages)
                    if (p.Annotations.Contains(a))
                    {
                        p.Annotations.Remove(a);
                        break;
                    }
            }
        }
        private void CheckToolbarVisibility()
        {
            bool shouldShow = (_currentTool == "TEXT") || (_selectedAnnotation != null);
            if (TextStyleToolbar != null)
                TextStyleToolbar.Visibility = shouldShow ? Visibility.Visible : Visibility.Collapsed;
        }

        private void ResizeThumb_DragDelta(object sender, DragDeltaEventArgs e)
        {
            var thumb = sender as Thumb;
            if (thumb?.DataContext is PdfAnnotation ann)
            {
                ann.Width = Math.Max(50, ann.Width + e.HorizontalChange);
                ann.Height = Math.Max(30, ann.Height + e.VerticalChange);
            }
        }
        private void UpdateToolbarFromAnnotation(PdfAnnotation ann)
        {
            _isUpdatingUiFromSelection = true;
            try
            {
                CbFont.SelectedItem = ann.FontFamily;
                CbSize.SelectedItem = ann.FontSize;
                BtnBold.IsChecked = ann.IsBold;
                if (ann.Foreground is SolidColorBrush brush)
                {
                    foreach (ComboBoxItem item in CbColor.Items)
                    {
                        string cn = item.Tag?.ToString() ?? "";
                        Color c = Colors.Black;
                        if (cn == "Red")
                            c = Colors.Red;
                        else if (cn == "Blue")
                            c = Colors.Blue;
                        else if (cn == "Green")
                            c = Colors.Green;
                        else if (cn == "Orange")
                            c = Colors.Orange;
                        if (brush.Color == c)
                        {
                            CbColor.SelectedItem = item;
                            break;
                        }
                    }
                }
            }
            finally { _isUpdatingUiFromSelection = false; }
        }
        private void StyleChanged(object sender, RoutedEventArgs e)
        {
            if (!IsLoaded || _isUpdatingUiFromSelection)
                return;
            string selectedFont = CbFont.SelectedItem?.ToString() ?? "Malgun Gothic";
            _defaultFontFamily = selectedFont;
            if (CbSize.SelectedItem != null)
                _defaultFontSize = (double)CbSize.SelectedItem;
            _defaultIsBold = BtnBold.IsChecked == true;
            if (CbColor.SelectedItem is ComboBoxItem item && item.Tag != null)
            {
                string cn = item.Tag.ToString() ?? "Black";
                if (cn == "Black")
                    _defaultFontColor = Colors.Black;
                else if (cn == "Red")
                    _defaultFontColor = Colors.Red;
                else if (cn == "Blue")
                    _defaultFontColor = Colors.Blue;
                else if (cn == "Green")
                    _defaultFontColor = Colors.Green;
                else if (cn == "Orange")
                    _defaultFontColor = Colors.Orange;
            }
            if (_selectedAnnotation != null)
            {
                _selectedAnnotation.FontFamily = _defaultFontFamily;
                _selectedAnnotation.FontSize = _defaultFontSize;
                _selectedAnnotation.IsBold = _defaultIsBold;
                _selectedAnnotation.Foreground = new SolidColorBrush(_defaultFontColor);
            }
        }

        private void Tool_Click(object sender, RoutedEventArgs e)
        {
            if (RbCursor.IsChecked == true)
                _currentTool = "CURSOR";
            else if (RbHighlight.IsChecked == true)
                _currentTool = "HIGHLIGHT";
            else if (RbText.IsChecked == true)
                _currentTool = "TEXT";
            CheckToolbarVisibility();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                TxtSearch.Focus();
                TxtSearch.SelectAll();
                e.Handled = true;
                return;
            }
            else if (e.Key == Key.Delete)
                BtnDeleteAnnotation_Click(this, new RoutedEventArgs());
            else if (e.Key == Key.Escape && _selectedAnnotation != null)
            {
                _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = null;
                CheckToolbarVisibility();
            }

            if (e.Key == Key.Space && !_isSpacePressed && _currentTool != "TEXT")
            {
                if (!(Keyboard.FocusedElement is TextBox))
                {
                    _isSpacePressed = true;
                    Mouse.OverrideCursor = Cursors.Hand;
                    e.Handled = true;
                }
            }
            else if (e.Key == Key.PageUp || e.Key == Key.PageDown)
            {
                var sv = GetScrollViewer();
                if (sv != null)
                {
                    if (e.Key == Key.PageDown)
                        sv.PageDown();
                    else
                        sv.PageUp();
                    e.Handled = true;
                }
            }
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            base.OnKeyUp(e);
            if (e.Key == Key.Space)
            {
                _isSpacePressed = false;
                _isPanning = false;
                Mouse.OverrideCursor = null;
                var listView = FindChild<ListView>(MainTabControl);
                if (listView != null)
                    listView.ReleaseMouseCapture();
            }
        }

        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual
        {
            if (parent == null)
                return null;
            T? child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++)
            {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null)
                    child = GetVisualChild<T>(v);
                if (child != null)
                    break;
            }
            return child;
        }

        private ListView? GetCurrentListView() => GetVisualChild<ListView>(MainTabControl);

        private T? FindChild<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
                return null;
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T typedChild)
                    return typedChild;
                var result = FindChild<T>(child);
                if (result != null)
                    return result;
            }
            return null;
        }

        private ScrollViewer? GetScrollViewer()
        {
            if (_cachedScrollViewer == null)
            {
                var listView = FindChild<ListView>(MainTabControl);
                if (listView != null)
                    _cachedScrollViewer = FindChild<ScrollViewer>(listView);
            }
            return _cachedScrollViewer;
        }

        private ScrollViewer? GetCurrentScrollViewer()
        {
            var listView = FindChild<ListView>(MainTabControl);
            if (listView == null)
                return null;
            return FindChild<ScrollViewer>(listView);
        }

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));

        private void BtnSign_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
            {
                MessageBox.Show("서명할 문서가 없습니다.");
                return;
            }

            var result = MessageBox.Show("이미지 도장을 사용하시겠습니까?\n(아니오를 누르면 텍스트 도장이 생성됩니다.)",
                                         "서명 방식 선택", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
                return;

            string? selectedImagePath = null;
            if (result == MessageBoxResult.Yes)
            {
                var fd = new OpenFileDialog { Title = "도장 이미지 선택", Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp" };
                if (fd.ShowDialog() == true)
                    selectedImagePath = fd.FileName;
                else
                    return;
            }

            if (SelectedDocument.Pages.Count == 0)
                return;
            var pageVM = SelectedDocument.Pages[0];

            var placeholder = new PdfAnnotation
            {
                Type = AnnotationType.SignaturePlaceholder,
                Width = 120,
                Height = 60,
                X = (pageVM.Width - 120) / 2,
                Y = (pageVM.Height - 60) / 2,
                IsSelected = true,
                VisualStampPath = selectedImagePath
            };

            if (selectedImagePath != null)
            {
                try
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.UriSource = new Uri(selectedImagePath);
                    bitmap.EndInit();
                    bitmap.Freeze();

                    double ratio = (double)bitmap.PixelWidth / bitmap.PixelHeight;
                    placeholder.Width = 100;
                    placeholder.Height = 100 / ratio;
                }
                catch { }
            }

            pageVM.Annotations.Add(placeholder);
            TxtStatus.Text = "서명 도장을 원하는 위치로 드래그하고 더블 클릭하세요.";
        }

        private void BtnVerifySignature_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            if (sender is FrameworkElement element && element.DataContext is PdfAnnotation ann)
            {
                if (ann.Type == AnnotationType.SignatureField)
                {
                    try
                    {
                        using (var doc = PdfReader.Open(SelectedDocument.FilePath, PdfDocumentOpenMode.Import))
                        {
                            PdfDictionary? targetWidget = null;
                            foreach (var page in doc.Pages)
                            {
                                if (page.Annotations != null)
                                {
                                    for (int i = 0; i < page.Annotations.Count; i++)
                                    {
                                        var pdfAnnot = page.Annotations[i];
                                        if (pdfAnnot.Elements.GetString("/Subtype") == "/Widget" &&
                                            pdfAnnot.Elements.GetString("/FT") == "/Sig")
                                        {
                                            string t = pdfAnnot.Elements.GetString("/T");
                                            if (t == ann.FieldName)
                                            {
                                                targetWidget = pdfAnnot.Elements.GetDictionary("/V");
                                                break;
                                            }
                                        }
                                    }
                                }
                                if (targetWidget != null)
                                    break;
                            }

                            if (targetWidget != null)
                            {
                                var verifier = new SignatureVerificationService();
                                var result = verifier.VerifySignature(SelectedDocument.FilePath, targetWidget);
                                var win = new SignatureResultWindow(result);
                                win.Owner = this;
                                win.ShowDialog();
                            }
                            else
                                MessageBox.Show("서명 데이터를 찾을 수 없습니다.");
                        }
                    }
                    catch (Exception ex) { MessageBox.Show($"검증 중 오류 발생: {ex.Message}"); }
                }
            }
        }

        private void BtnToggleSidebar_Click(object sender, RoutedEventArgs e)
        {
            if (SidebarBorder.Visibility == Visibility.Visible)
                SidebarBorder.Visibility = Visibility.Collapsed;
            else
                SidebarBorder.Visibility = Visibility.Visible;
        }

        private void BookmarkTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is PdfBookmarkViewModel bm && SelectedDocument != null)
            {
                var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex);
                if (targetPage != null)
                {
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null)
                        listView.ScrollIntoView(targetPage);
                }
            }
        }

        private Point _startPoint;
        private bool _isDragging = false;

        private void BookmarkTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _startPoint = e.GetPosition(null);
        }

        private void BookmarkTree_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed && !_isDragging)
            {
                Point mousePos = e.GetPosition(null);
                Vector diff = _startPoint - mousePos;
                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    if (BookmarkTree.SelectedItem is PdfBookmarkViewModel selectedBm)
                    {
                        if (selectedBm.IsEditing)
                            return;
                        _isDragging = true;
                        DataObject dragData = new DataObject("MinsBookmark", selectedBm);
                        DragDrop.DoDragDrop(BookmarkTree, dragData, DragDropEffects.Move);
                        _isDragging = false;
                    }
                }
            }
        }

        private void BookmarkTree_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }
            if (e.Data.GetDataPresent("MinsBookmark") && SelectedDocument != null)
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
                return;
            }
            e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void BookmarkTree_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                Window_Drop(sender, e);
                return;
            }

            if (!e.Data.GetDataPresent("MinsBookmark") || SelectedDocument == null)
                return;
            var sourceBm = e.Data.GetData("MinsBookmark") as PdfBookmarkViewModel;
            if (sourceBm == null)
                return;

            var targetItem = FindAncestor<TreeViewItem>((DependencyObject)e.OriginalSource);
            var targetBm = targetItem?.DataContext as PdfBookmarkViewModel;

            if (targetBm != null && (targetBm == sourceBm || IsChildOf(targetBm, sourceBm)))
            {
                MessageBox.Show("자기 자신이나 하위 항목으로 이동할 수 없습니다.");
                return;
            }

            var oldCollection = GetCurrentCollection(sourceBm);
            if (oldCollection != null)
                oldCollection.Remove(sourceBm);

            if (targetBm == null)
            {
                SelectedDocument.Bookmarks.Add(sourceBm);
                sourceBm.Parent = null;
            }
            else
            {
                targetBm.Children.Add(sourceBm);
                sourceBm.Parent = targetBm;
                targetBm.IsExpanded = true;
            }
            e.Handled = true;
        }

        private bool IsChildOf(PdfBookmarkViewModel potentialChild, PdfBookmarkViewModel potentialParent)
        {
            var current = potentialChild.Parent;
            while (current != null)
            {
                if (current == potentialParent)
                    return true;
                current = current.Parent;
            }
            return false;
        }

        private static T? FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T typed)
                    return typed;
                current = VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private int GetCurrentPageIndex()
        {
            if (SelectedDocument == null)
                return 0;
            var listView = FindChild<ListView>(MainTabControl);
            if (listView == null)
                return 0;
            var scrollViewer = FindChild<ScrollViewer>(listView);
            if (scrollViewer == null)
                return 0;

            double currentOffset = scrollViewer.VerticalOffset;
            double accumulatedHeight = 0;
            foreach (var page in SelectedDocument.Pages)
            {
                double pageTotalHeight = (page.Height * SelectedDocument.Zoom) + 20;
                if (accumulatedHeight + pageTotalHeight > currentOffset + 50)
                    return page.PageIndex;
                accumulatedHeight += pageTotalHeight;
            }
            return 0;
        }

        private void BookmarkTree_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F2 && BookmarkTree.SelectedItem is PdfBookmarkViewModel bm)
                bm.IsEditing = true;
        }

        private void BtnRenameBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (sender is MenuItem item && item.DataContext is PdfBookmarkViewModel bm)
                bm.IsEditing = true;
            else if (BookmarkTree.SelectedItem is PdfBookmarkViewModel sBm)
                sBm.IsEditing = true;
        }

        private void BookmarkRename_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if ((sender as TextBox)?.DataContext is PdfBookmarkViewModel bm)
                    bm.IsEditing = false;
                e.Handled = true;
            }
        }

        private void BookmarkRename_LostFocus(object sender, RoutedEventArgs e)
        {
            if ((sender as TextBox)?.DataContext is PdfBookmarkViewModel bm)
                bm.IsEditing = false;
        }

        private void PdfListView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int currentIndex = GetCurrentPageIndex();
            int totalPages = SelectedDocument.Pages.Count;
            TxtPageInfo.Text = $"{currentIndex + 1} / {totalPages}";
        }

        private void BtnAddBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int currentIndex = GetCurrentPageIndex();
            var newBookmark = new PdfBookmarkViewModel
            {
                Title = $"Page {currentIndex + 1}",
                PageIndex = currentIndex,
                Parent = null
            };
            SelectedDocument.Bookmarks.Add(newBookmark);
        }

        private void BtnUpdateBookmarkPage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            if (BookmarkTree.SelectedItem is PdfBookmarkViewModel selectedBm)
            {
                int currentPageIdx = GetCurrentPageIndex();
                selectedBm.PageIndex = currentPageIdx;
                if (selectedBm.Title.StartsWith("새 책갈피"))
                    selectedBm.Title = $"새 책갈피 (p.{currentPageIdx + 1})";
                MessageBox.Show($"'{selectedBm.Title}'의 위치가 {currentPageIdx + 1}페이지로 변경되었습니다.");
            }
        }

        private void BookmarkItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is FrameworkElement element && element.DataContext is PdfBookmarkViewModel bm)
            {
                if (SelectedDocument != null)
                {
                    var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex);
                    if (targetPage != null)
                    {
                        var listView = GetVisualChild<ListView>(MainTabControl);
                        if (listView != null)
                            listView.ScrollIntoView(targetPage);
                    }
                }
            }
        }

        private void BtnDeleteBookmark_Click(object sender, RoutedEventArgs e)
        {
            PdfBookmarkViewModel? targetBm = null;
            if (sender is MenuItem menuItem && menuItem.DataContext is PdfBookmarkViewModel contextBm)
                targetBm = contextBm;
            else
                targetBm = BookmarkTree.SelectedItem as PdfBookmarkViewModel;

            if (targetBm != null && SelectedDocument != null)
            {
                if (SelectedDocument.Bookmarks.Contains(targetBm))
                    SelectedDocument.Bookmarks.Remove(targetBm);
                else if (targetBm.Parent != null)
                    targetBm.Parent.Children.Remove(targetBm);
            }
        }

        private void BtnMoveBookmarkUp_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;
            var currentList = GetCurrentCollection(selectedBm);
            if (currentList == null)
                return;
            int index = currentList.IndexOf(selectedBm);
            if (index > 0)
            {
                currentList.Move(index, index - 1);
                selectedBm.IsSelected = true;
            }
        }

        private void BtnMoveBookmarkDown_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;
            var currentList = GetCurrentCollection(selectedBm);
            if (currentList == null)
                return;
            int index = currentList.IndexOf(selectedBm);
            if (index >= 0 && index < currentList.Count - 1)
            {
                currentList.Move(index, index + 1);
                selectedBm.IsSelected = true;
            }
        }

        private void BtnIndentBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;
            var currentList = GetCurrentCollection(selectedBm);
            if (currentList == null)
                return;
            int index = currentList.IndexOf(selectedBm);
            if (index > 0)
            {
                var newParent = currentList[index - 1];
                currentList.Remove(selectedBm);
                newParent.Children.Add(selectedBm);
                selectedBm.Parent = newParent;
                newParent.IsExpanded = true;
                selectedBm.IsSelected = true;
            }
        }

        private void BtnOutdentBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;
            if (selectedBm.Parent == null)
                return;
            var oldParent = selectedBm.Parent;
            var targetList = (oldParent.Parent == null) ? SelectedDocument.Bookmarks : oldParent.Parent.Children;
            int parentIndex = targetList.IndexOf(oldParent);
            if (parentIndex >= 0)
            {
                oldParent.Children.Remove(selectedBm);
                targetList.Insert(parentIndex + 1, selectedBm);
                selectedBm.Parent = oldParent.Parent;
                selectedBm.IsSelected = true;
            }
        }

        private System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>? GetCurrentCollection(PdfBookmarkViewModel bm)
        {
            if (SelectedDocument == null)
                return null;
            if (bm.Parent == null)
                return SelectedDocument.Bookmarks;
            else
                return bm.Parent.Children;
        }

        public async void OpenPdfFromPath(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                    return;
                if (Path.GetExtension(filePath).ToLower() != ".pdf")
                {
                    MessageBox.Show("PDF 파일만 지원합니다.");
                    return;
                }
                TxtStatus.Text = $"파일 여는 중: {Path.GetFileName(filePath)}...";
                var docModel = await _pdfService.LoadPdfAsync(filePath);
                if (docModel != null)
                {
                    Documents.Add(docModel);
                    SelectedDocument = docModel;
                    await _pdfService.InitializeDocumentAsync(docModel);
                    int lastPage = _historyService.GetLastPage(filePath);
                    if (lastPage > 0 && lastPage < docModel.Pages.Count)
                    {
                        await Task.Delay(100);
                        var listView = GetVisualChild<ListView>(MainTabControl);
                        if (listView != null)
                        {
                            var targetPage = docModel.Pages[lastPage];
                            listView.ScrollIntoView(targetPage);
                            TxtStatus.Text = $"이전 위치(p.{lastPage + 1})로 복원됨";
                        }
                    }
                    else
                        TxtStatus.Text = "준비 완료";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 열기 실패: {ex.Message}");
                TxtStatus.Text = "오류 발생";
            }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog { Filter = "PDF Files (*.pdf)|*.pdf" };
            if (dlg.ShowDialog() == true)
                OpenPdfFromPath(dlg.FileName);
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string filePath = files[0];
                    if (System.IO.Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                        OpenPdfFromPath(filePath);
                }
            }
        }

        private void BtnRotateLeft_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int idx = GetCurrentPageIndex();
            if (idx >= 0 && idx < SelectedDocument.Pages.Count)
            {
                var page = SelectedDocument.Pages[idx];
                page.Rotation = (page.Rotation - 90);
                if (page.Rotation < 0)
                    page.Rotation += 360;
                double temp = page.Width;
                page.Width = page.Height;
                page.Height = temp;
            }
        }

        private void BtnRotateRight_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int idx = GetCurrentPageIndex();
            if (idx >= 0 && idx < SelectedDocument.Pages.Count)
            {
                var page = SelectedDocument.Pages[idx];
                page.Rotation = (page.Rotation + 90) % 360;
                double temp = page.Width;
                page.Width = page.Height;
                page.Height = temp;
            }
        }

        private void BtnDeletePage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int idx = GetCurrentPageIndex();
            if (idx >= 0 && idx < SelectedDocument.Pages.Count)
            {
                if (MessageBox.Show($"{idx + 1}페이지를 삭제하시겠습니까?\n(저장해야 반영됩니다)", "삭제 확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    SelectedDocument.Pages.RemoveAt(idx);
                    for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                        SelectedDocument.Pages[i].PageIndex = i;
                    TxtStatus.Text = "페이지 삭제됨 (저장 시 반영)";
                    TxtPageInfo.Text = $"{Math.Min(idx + 1, SelectedDocument.Pages.Count)} / {SelectedDocument.Pages.Count}";
                }
            }
        }

        private void BtnAddBlankPage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int idx = GetCurrentPageIndex();
            var newPage = new PdfPageViewModel
            {
                IsBlankPage = true,
                OriginalFilePath = null,
                Width = 595,
                Height = 842,
                PdfPageWidthPoint = 595,
                PdfPageHeightPoint = 842,
                Rotation = 0,
                PageIndex = idx + 1
            };
            SelectedDocument.Pages.Insert(idx + 1, newPage);
            for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                SelectedDocument.Pages[i].PageIndex = i;
            TxtStatus.Text = "빈 페이지 추가됨";
            TxtPageInfo.Text = $"{idx + 2} / {SelectedDocument.Pages.Count}";
        }

        private void PdfListView_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var listView = sender as ListView;
            if (listView == null)
                return;
            if (_isSpacePressed || e.ChangedButton == MouseButton.Middle)
            {
                _isPanning = true;
                _lastPanPoint = e.GetPosition(this);
                listView.CaptureMouse();
                if (e.ChangedButton == MouseButton.Middle)
                    Mouse.OverrideCursor = Cursors.ScrollAll;
                e.Handled = true;
            }
        }

        private void PdfListView_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_isPanning)
            {
                var listView = sender as ListView;
                var sv = GetScrollViewer();
                if (sv != null && listView != null)
                {
                    var currentPoint = e.GetPosition(this);
                    var delta = currentPoint - _lastPanPoint;
                    sv.ScrollToHorizontalOffset(sv.HorizontalOffset - delta.X);
                    sv.ScrollToVerticalOffset(sv.VerticalOffset - delta.Y);
                    _lastPanPoint = currentPoint;
                }
            }
        }

        private void PdfListView_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isPanning)
            {
                _isPanning = false;
                var listView = sender as ListView;
                if (listView != null)
                    listView.ReleaseMouseCapture();
                if (!_isSpacePressed)
                    Mouse.OverrideCursor = null;
                else
                    Mouse.OverrideCursor = Cursors.Hand;
            }
        }

        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_isSpacePressed || e.ChangedButton == MouseButton.Middle)
            {
                _isPanning = true;
                _lastPanPoint = e.GetPosition(this);
                Mouse.Capture(this.Content as UIElement);
                if (e.ChangedButton == MouseButton.Middle)
                    Mouse.OverrideCursor = Cursors.ScrollAll;
                e.Handled = true;
            }
        }

        private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_isPanning)
            {
                var scrollViewer = GetCurrentScrollViewer();
                if (scrollViewer != null)
                {
                    var currentPoint = e.GetPosition(this);
                    var delta = currentPoint - _lastPanPoint;
                    scrollViewer.ScrollToHorizontalOffset(scrollViewer.HorizontalOffset - delta.X);
                    scrollViewer.ScrollToVerticalOffset(scrollViewer.VerticalOffset - delta.Y);
                    _lastPanPoint = currentPoint;
                }
            }
        }

        private void Window_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_isPanning)
            {
                _isPanning = false;
                Mouse.Capture(null);
                if (_isSpacePressed)
                    Mouse.OverrideCursor = Cursors.Hand;
                else
                    Mouse.OverrideCursor = null;
            }
        }

        private void Window_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (SelectedDocument != null)
                {
                    if (e.Delta > 0)
                    {
                        if (SelectedDocument.Zoom < 5.0)
                            SelectedDocument.Zoom += 0.1;
                    }
                    else
                    {
                        if (SelectedDocument.Zoom > 0.2)
                            SelectedDocument.Zoom -= 0.1;
                    }
                }
                e.Handled = true;
            }
        }
    }
}