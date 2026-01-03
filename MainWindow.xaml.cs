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
using Docnet.Core.Models; // 상단에 추가해 주세요
using Microsoft.Win32;
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

        private readonly HistoryService _historyService; // [신규]


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

        // MainWindow.xaml.cs -> UpdateScrollViewerState 메서드 교체

        private void UpdateScrollViewerState(ListView listView, PdfDocumentModel? oldDoc, PdfDocumentModel? newDoc)
        {
            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer == null)
                return;

            // 이벤트 중복 방지
            scrollViewer.ScrollChanged -= ScrollViewer_ScrollChanged;

            // 1. 이전 문서의 스크롤 위치 저장
            if (oldDoc != null)
            {
                oldDoc.SavedVerticalOffset = scrollViewer.VerticalOffset;
                oldDoc.SavedHorizontalOffset = scrollViewer.HorizontalOffset;
            }

            // 2. 새 문서의 스크롤 위치 복원 (비동기 처리)
            if (newDoc != null)
            {
                // [핵심 수정] UI 렌더링이 완료된 후에 스크롤을 이동시켜야 
                // 이전 탭의 위치가 남아있거나 0으로 초기화되는 문제를 막을 수 있음
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    scrollViewer.ScrollToVerticalOffset(newDoc.SavedVerticalOffset);
                    scrollViewer.ScrollToHorizontalOffset(newDoc.SavedHorizontalOffset);

                    // 스크롤 복원 후 이벤트 다시 연결
                    scrollViewer.ScrollChanged += ScrollViewer_ScrollChanged;
                }, System.Windows.Threading.DispatcherPriority.Loaded);
            }
            else
            {
                // 문서가 없을 때는 이벤트만 다시 연결
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

        // // [수정] async void 적용 및 LoadPdfAsync 호출
        // private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        // {
        //     var dlg = new OpenFileDialog { Filter = "PDF Files|*.pdf" };
        //     if (dlg.ShowDialog() == true)
        //     {
        //         // UI 스레드에서 락을 기다리지 않고, 작업이 끝날 때까지 부드럽게 대기(await)
        //         var docModel = await _pdfService.LoadPdfAsync(dlg.FileName);

        //         if (docModel != null)
        //         {
        //             Documents.Add(docModel);
        //             SelectedDocument = docModel;
        //             // 초기화(렌더링 준비) 시작
        //             _ = _pdfService.InitializeDocumentAsync(docModel);
        //         }
        //     }
        // }

        // -------------------------------------------------------------
        // [수정] 탭 닫기 (저장 로직 추가)
        // -------------------------------------------------------------
        private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PdfDocumentModel doc)
            {
                // [신규] 닫기 전에 현재 보고 있는 페이지 저장
                // (주의: 탭이 닫히려는 순간에 이 문서가 'SelectedDocument'가 아닐 수도 있음.
                //  따라서 닫으려는 문서가 활성화되어 있다면 현재 스크롤 위치를 저장)
                if (doc == SelectedDocument)
                {
                    int currentPage = GetCurrentPageIndex();
                    _historyService.SetLastPage(doc.FilePath, currentPage);
                    _historyService.SaveHistory(); // 즉시 파일 쓰기
                }

                doc.Dispose();
                Documents.Remove(doc);
            }
        }

        // -------------------------------------------------------------
        // [신규] 프로그램 종료 시 (열려있는 모든 탭 저장)
        // -------------------------------------------------------------
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 열려있는 모든 문서의 현재 위치를 저장해야 하지만,
            // ListView는 하나뿐이므로 '현재 선택된 탭'의 위치만 정확히 알 수 있음.

            if (SelectedDocument != null)
            {
                int currentPage = GetCurrentPageIndex();
                _historyService.SetLastPage(SelectedDocument.FilePath, currentPage);
            }

            // 전체 기록을 파일로 저장
            _historyService.SaveHistory();

            // 메모리 정리
            foreach (var doc in Documents)
            {
                doc.Dispose();
            }
        }

        // [신규] 페이지가 화면에 보일 때 -> 이미지 렌더링
        private void PageGrid_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement elem && elem.DataContext is PdfPageViewModel pageVM)
            {
                if (SelectedDocument != null)
                {
                    // 비동기로 렌더링 요청 (UI 프리징 방지)
                    Task.Run(() => _pdfService.RenderPageImage(SelectedDocument, pageVM));
                }
            }
        }

        // [신규] 페이지가 화면에서 사라질 때 -> 이미지 메모리 해제
        private void PageGrid_Unloaded(object sender, RoutedEventArgs e)
        {
            if (sender is FrameworkElement elem && elem.DataContext is PdfPageViewModel pageVM)
            {
                pageVM.Unload(); // 이미지 null 처리
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
            if (foundAnnot != null)
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
                var result = MessageBox.Show("문서의 시작입니다. 끝에서부터 다시 찾으시겠습니까?", "검색 완료", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.OK)
                {
                    _searchService.ResetSearch();
                    await _searchService.FindPrevAsync(SelectedDocument, query);
                }
                else
                {
                    TxtStatus.Text = "검색 종료";
                }
            }
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
            if (foundAnnot != null)
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
                var result = MessageBox.Show("문서의 끝입니다. 처음부터 다시 찾으시겠습니까?", "검색 완료", MessageBoxButton.OKCancel, MessageBoxImage.Question);
                if (result == MessageBoxResult.OK)
                {
                    _searchService.ResetSearch();
                    await FindNextSearchResult();
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

        private void PerformSignatureProcess(PdfAnnotation ann)
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

                if (!string.IsNullOrEmpty(ann.VisualStampPath))
                {
                    config.VisualStampPath = ann.VisualStampPath;
                    config.UseVisualStamp = true;
                }
                else
                {
                    config.VisualStampPath = null;
                    config.UseVisualStamp = true;
                }

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
                        _pdfService.SavePdf(SelectedDocument, tempPath);

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
                await Task.Run(async () => { for (int i = 0; i < targetDoc.Pages.Count; i++) { var pageVM = targetDoc.Pages[i]; if (targetDoc.DocReader == null) continue; using (var r = targetDoc.DocReader.GetPageReader(i)) { var rawBytes = r.GetImage(); var w = r.GetPageWidth(); var h = r.GetPageHeight(); using (var stream = new MemoryStream(rawBytes)) { var ibuffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(rawBytes); using (var softwareBitmap = new SoftwareBitmap(BitmapPixelFormat.Bgra8, w, h, BitmapAlphaMode.Premultiplied)) { softwareBitmap.CopyFromBuffer(ibuffer); var ocrResult = await _ocrEngine.RecognizeAsync(softwareBitmap); var wordList = new List<OcrWordInfo>(); foreach (var line in ocrResult.Lines) { foreach (var word in line.Words) { wordList.Add(new OcrWordInfo { Text = word.Text, BoundingBox = new Rect(word.BoundingRect.X, word.BoundingRect.Y, word.BoundingRect.Width, word.BoundingRect.Height) }); } } pageVM.OcrWords = wordList; } } } Application.Current.Dispatcher.Invoke(() => { PbStatus.Value = i + 1; }); } });
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

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument != null)
                _pdfService.SavePdf(SelectedDocument, SelectedDocument.FilePath);
        }
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            var dlg = new SaveFileDialog { Filter = "PDF Files|*.pdf", FileName = Path.GetFileNameWithoutExtension(SelectedDocument.FilePath) + "_annotated" };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    _pdfService.SavePdf(SelectedDocument, dlg.FileName);
                    SelectedDocument.FilePath = dlg.FileName;
                    SelectedDocument.FileName = Path.GetFileName(dlg.FileName);
                    MessageBox.Show("저장되었습니다.");
                }
                catch (Exception ex) { MessageBox.Show($"저장 실패: {ex.Message}"); }
            }
        }

        // [수정] 비동기(async)로 변경하여 UI 멈춤 방지
        // [수정] async 키워드 추가 (비동기 메서드로 변경)
        // MainWindow.xaml.cs

        // [수정] 검색 기능과 동일한 "좌표계 자동 감지" 로직 적용
        private async void CheckTextInSelection(int pageIndex, Rect uiRect)
        {
            _selectedTextBuffer = "";
            if (SelectedDocument?.DocReader == null)
                return;

            var doc = SelectedDocument;

            string extractedText = await Task.Run(() =>
            {
                var sb = new StringBuilder();
                try
                {
                    lock (PdfService.PdfiumLock)
                    {
                        if (doc.DocReader == null)
                            return "";

                        using (var reader = doc.DocReader.GetPageReader(pageIndex))
                        {
                            var chars = reader.GetCharacters().ToList();
                            if (chars.Count == 0)
                                return "";

                            // 1. 좌표계 자동 감지 (SearchService와 동일 로직)
                            bool needsFlip = false;
                            if (chars.Count > 5)
                            {
                                double firstY = (chars[0].Box.Top + chars[0].Box.Bottom) / 2.0;
                                double lastY = (chars.Last().Box.Top + chars.Last().Box.Bottom) / 2.0;
                                if (firstY > lastY)
                                    needsFlip = true; // 표준 PDF (뒤집어야 함)
                            }

                            // 2. 페이지 정보 가져오기
                            var pageVM = doc.Pages[pageIndex];
                            double scaleX = (pageVM.CropWidthPoint > 0) ? pageVM.Width / pageVM.CropWidthPoint : 1.0;
                            double scaleY = (pageVM.CropHeightPoint > 0) ? pageVM.Height / pageVM.CropHeightPoint : 1.0;

                            foreach (var c in chars)
                            {
                                // 3. PDF 좌표 -> UI 좌표로 변환 (검색 로직 역이용)
                                double charMinX = Math.Min(c.Box.Left, c.Box.Right);
                                double charMaxX = Math.Max(c.Box.Left, c.Box.Right);
                                double charMinY = Math.Min(c.Box.Top, c.Box.Bottom);
                                double charMaxY = Math.Max(c.Box.Top, c.Box.Bottom);

                                double finalX = (charMinX - pageVM.CropX) * scaleX;
                                double finalY = 0;
                                double finalW = (charMaxX - charMinX) * scaleX;
                                double finalH = (charMaxY - charMinY) * scaleY;

                                if (needsFlip)
                                {
                                    // 표준 PDF: (PageHeight - MaxY)가 Top이 됨
                                    double pdfTopFromPageTop = pageVM.PdfPageHeightPoint - charMaxY;
                                    finalY = (pdfTopFromPageTop - pageVM.CropY) * scaleY;
                                }
                                else
                                {
                                    // 이미지형 PDF: MinY가 Top이 됨
                                    finalY = (charMinY - pageVM.CropY) * scaleY;
                                }

                                // 4. 변환된 글자 박스와 마우스 선택 영역(uiRect)이 겹치는지 확인
                                var charRect = new Rect(finalX, finalY, finalW, finalH);
                                if (uiRect.IntersectsWith(charRect))
                                {
                                    sb.Append(c.Char);
                                }
                            }
                        }
                    }
                }
                catch { return ""; }

                // OCR 단어 처리 (단순 좌표 비교)
                if (doc.Pages.Count > pageIndex)
                {
                    var pageVM = doc.Pages[pageIndex];
                    if (pageVM.OcrWords != null)
                    {
                        foreach (var word in pageVM.OcrWords)
                        {
                            if (uiRect.IntersectsWith(word.BoundingBox))
                                sb.Append(word.Text + " ");
                        }
                    }
                }

                return sb.ToString();
            });

            _selectedTextBuffer = extractedText;

            if (!string.IsNullOrEmpty(_selectedTextBuffer))
            {
                TxtStatus.Text = "텍스트 선택됨"; // 상태바 메시지
            }
        }

        private void BtnPopupCopy_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(_selectedTextBuffer);
            SelectionPopup.IsOpen = false;
        }
        // [구현] 선택 영역을 고해상도 이미지로 캡처하여 클립보드에 복사
        private async void BtnPopupCopyImage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SelectionPopup.IsOpen = false; // 팝업 닫기

                if (SelectedDocument == null || _selectedPageIndex == -1)
                    return;

                var pageVM = SelectedDocument.Pages[_selectedPageIndex];

                // 선택 영역이 너무 작으면 무시
                if (pageVM.SelectionWidth <= 0 || pageVM.SelectionHeight <= 0)
                    return;

                // 1. 고해상도 설정을 위한 스케일 (2.0 = 약 192 DPI, 3.0 = 약 288 DPI)
                // 너무 높으면 메모리 부족 발생 가능하므로 2.0 권장
                double renderScale = 2.0;

                byte[]? fullPageBytes = null;
                int rawWidth = 0;
                int rawHeight = 0;

                TxtStatus.Text = "이미지 처리 중...";

                // 2. 백그라운드에서 페이지 전체를 고해상도로 렌더링
                await Task.Run(() =>
                {
                    lock (PdfService.PdfiumLock)
                    {
                        // DocLib 인스턴스 확인
                        if (SelectedDocument.DocLib == null)
                            return;

                        try
                        {
                            // 현재 파일을 다시 읽어서 임시 고해상도 Reader 생성
                            // (기존 Reader는 1.0배율로 고정되어 있어서 흐릿함)
                            var fileBytes = File.ReadAllBytes(SelectedDocument.FilePath);

                            using (var reader = SelectedDocument.DocLib.GetDocReader(fileBytes, new PageDimensions(renderScale)))
                            using (var pageReader = reader.GetPageReader(_selectedPageIndex))
                            {
                                rawWidth = pageReader.GetPageWidth();
                                rawHeight = pageReader.GetPageHeight();
                                fullPageBytes = pageReader.GetImage(); // BGRA 포맷 바이트 배열
                            }
                        }
                        catch (Exception ex)
                        {
                            // 파일 접근 오류 등 예외 처리
                            System.Diagnostics.Debug.WriteLine($"Image Copy Error: {ex.Message}");
                        }
                    }
                });

                if (fullPageBytes == null)
                {
                    MessageBox.Show("이미지 생성에 실패했습니다.");
                    return;
                }

                // 3. 바이트 배열 -> BitmapSource 변환
                var fullBitmap = BitmapSource.Create(
                    rawWidth, rawHeight,
                    96 * renderScale, 96 * renderScale,
                    PixelFormats.Bgra32, null,
                    fullPageBytes, rawWidth * 4);

                // 4. 선택 영역에 맞춰 자르기 (Crop)
                // 화면상의 비율(Ratio)을 계산하여 실제 고해상도 이미지 좌표로 변환
                double ratioX = pageVM.SelectionX / pageVM.Width;
                double ratioY = pageVM.SelectionY / pageVM.Height;
                double ratioW = pageVM.SelectionWidth / pageVM.Width;
                double ratioH = pageVM.SelectionHeight / pageVM.Height;

                int cropX = (int)(rawWidth * ratioX);
                int cropY = (int)(rawHeight * ratioY);
                int cropW = (int)(rawWidth * ratioW);
                int cropH = (int)(rawHeight * ratioH);

                // 좌표 보정 (이미지 범위 벗어남 방지)
                if (cropX < 0)
                    cropX = 0;
                if (cropY < 0)
                    cropY = 0;
                if (cropX + cropW > rawWidth)
                    cropW = rawWidth - cropX;
                if (cropY + cropH > rawHeight)
                    cropH = rawHeight - cropY;

                if (cropW > 0 && cropH > 0)
                {
                    // 자르기 실행
                    var croppedBitmap = new CroppedBitmap(fullBitmap, new Int32Rect(cropX, cropY, cropW, cropH));

                    // 클립보드에 설정
                    Clipboard.SetImage(croppedBitmap);
                    TxtStatus.Text = "이미지가 클립보드에 복사되었습니다.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"이미지 복사 실패: {ex.Message}");
                TxtStatus.Text = "오류 발생";
            }
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
        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                TxtSearch.Focus();
                TxtSearch.SelectAll();
                e.Handled = true;
            }
            else if (e.Key == Key.Delete)
                BtnDeleteAnnotation_Click(this, new RoutedEventArgs());
            else if (e.Key == Key.Escape && _selectedAnnotation != null)
            {
                _selectedAnnotation.IsSelected = false;
                _selectedAnnotation = null;
                CheckToolbarVisibility();
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
        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(name));

        // [복구] 전자서명 버튼 (누락됐던 부분)
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
                var fd = new OpenFileDialog
                {
                    Title = "도장 이미지 선택",
                    Filter = "Image Files|*.png;*.jpg;*.jpeg;*.bmp"
                };

                if (fd.ShowDialog() == true)
                {
                    selectedImagePath = fd.FileName;
                }
                else
                {
                    return;
                }
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
                                    // [수정] foreach 대신 for문을 사용하고 인덱스로 접근하여
                                    // PdfItem이 아닌 PdfAnnotation(혹은 PdfDictionary) 타입으로 확실하게 받음
                                    for (int i = 0; i < page.Annotations.Count; i++)
                                    {
                                        var pdfAnnot = page.Annotations[i];

                                        // Null 체크 및 요소 접근
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
                            {
                                MessageBox.Show("서명 데이터를 찾을 수 없습니다.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"검증 중 오류 발생: {ex.Message}");
                    }
                }
            }
        }

        // [신규] 책갈피 사이드바 열기/닫기
        private void BtnToggleSidebar_Click(object sender, RoutedEventArgs e)
        {
            if (SidebarBorder.Visibility == Visibility.Visible)
                SidebarBorder.Visibility = Visibility.Collapsed;
            else
                SidebarBorder.Visibility = Visibility.Visible;
        }

        // [신규] PDF 로드 시 책갈피도 같이 로드 (LoadPdfAsync 성공 후 호출 권장)
        // (기존 BtnOpen_Click에서 InitializeDocumentAsync 호출 후 _pdfService.LoadBookmarks(docModel); 추가 필요)

        // [신규] 트리뷰 선택 시 해당 페이지로 이동
        private void BookmarkTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is PdfBookmarkViewModel bm && SelectedDocument != null)
            {
                // 1. 해당 페이지 객체 찾기
                var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex);
                if (targetPage != null)
                {
                    // 2. 리스트뷰 스크롤 이동
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null)
                        listView.ScrollIntoView(targetPage);
                }
            }
        }

        // [신규] 현재 화면에 보이는 페이지 번호를 계산하는 헬퍼 메서드
        private int GetCurrentPageIndex()
        {
            if (SelectedDocument == null)
                return 0;

            // 리스트뷰 내부의 스크롤뷰어 찾기
            var listView = GetVisualChild<ListView>(MainTabControl); // 탭 안에 있는 경우
            if (listView == null)
                return 0;

            var scrollViewer = GetVisualChild<ScrollViewer>(listView);
            if (scrollViewer == null)
                return 0;

            double currentOffset = scrollViewer.VerticalOffset;
            double accumulatedHeight = 0;

            // 가변 높이 페이지 대응: 위에서부터 높이를 더해가며 현재 스크롤 위치와 비교
            foreach (var page in SelectedDocument.Pages)
            {
                // 페이지 높이 + 여백(Margin 10+10=20) * 줌
                // 주의: 뷰모델의 Height는 '현재 렌더링된 크기'가 아닐 수 있으므로 
                // 렌더링 스케일 로직과 맞춰야 하지만, 보통 Width/Height 바인딩 된 값을 씁니다.
                // 여기서는 간단히 뷰모델 값 사용

                double pageTotalHeight = (page.Height * SelectedDocument.Zoom) + 20;

                // 현재 스크롤 위치가 이 페이지 범위 안에 걸치면 당첨
                // (조금 넉넉하게 currentOffset + 50 정도 줌)
                if (accumulatedHeight + pageTotalHeight > currentOffset + 50)
                {
                    return page.PageIndex;
                }
                accumulatedHeight += pageTotalHeight;
            }

            return 0; // 못 찾으면 0
        }

        private void PdfListView_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (SelectedDocument == null)
                return;

            // 현재 보고 있는 페이지 인덱스 계산 (0부터 시작)
            int currentIndex = GetCurrentPageIndex();
            int totalPages = SelectedDocument.Pages.Count;

            // UI 업데이트 (사용자에게는 1부터 시작하는 숫자로 보여줌)
            TxtPageInfo.Text = $"{currentIndex + 1} / {totalPages}";
        }


        // [수정] 책갈피 추가 (현재 페이지 자동 인식)
        private void BtnAddBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;

            // [에러 해결] 이 줄이 빠져 있었습니다.
            int currentIndex = GetCurrentPageIndex();

            // 제목은 사람용(1부터), 내부는 기계용(0부터)
            var newBookmark = new PdfBookmarkViewModel
            {
                Title = $"Page {currentIndex + 1}",
                PageIndex = currentIndex,
                Parent = null
            };

            SelectedDocument.Bookmarks.Add(newBookmark);
        }

        //[신규] 우클릭 -> "현재 페이지로 위치 갱신" 기능
        private void BtnUpdateBookmarkPage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;

            // 우클릭한 대상이 아니라 '현재 선택된' 아이템을 기준으로 처리
            // (TreeView는 우클릭 시 선택 상태가 변하지 않을 수 있으므로 주의 필요하지만, 
            //  보통 우클릭 전 좌클릭을 하므로 SelectedItem 사용)
            if (BookmarkTree.SelectedItem is PdfBookmarkViewModel selectedBm)
            {
                int currentPageIdx = GetCurrentPageIndex();
                selectedBm.PageIndex = currentPageIdx;

                // 이름이 기본값(새 책갈피...)이면 센스있게 페이지 번호 업데이트, 
                // 사용자가 이름을 바꿨다면 이름은 유지
                if (selectedBm.Title.StartsWith("새 책갈피"))
                {
                    selectedBm.Title = $"새 책갈피 (p.{currentPageIdx + 1})";
                }

                MessageBox.Show($"'{selectedBm.Title}'의 위치가 {currentPageIdx + 1}페이지로 변경되었습니다.");
            }
        }

        // MainWindow.xaml.cs

        // [신규] 책갈피 항목 클릭 시 페이지 이동 (선택 여부와 관계없이 동작)
        private void BookmarkItem_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            // 클릭된 요소의 데이터(ViewModel) 가져오기
            if (sender is FrameworkElement element && element.DataContext is PdfBookmarkViewModel bm)
            {
                if (SelectedDocument != null)
                {
                    var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.PageIndex == bm.PageIndex);
                    if (targetPage != null)
                    {
                        var listView = GetVisualChild<ListView>(MainTabControl);
                        if (listView != null)
                        {
                            // [핵심] 해당 페이지로 즉시 스크롤
                            listView.ScrollIntoView(targetPage);

                            // (선택 사항) 약간의 강제 스크롤 보정
                            // ScrollIntoView는 이미 보이면 동작 안 할 수 있으므로, 확실한 이동을 위해 추가 처리 가능
                        }
                    }
                }
            }
            // 주의: e.Handled = true를 하지 않아야 텍스트박스 포커스나 트리 선택이 정상 동작함
        }


        // [수정] 삭제 버튼 (툴바 버튼 + 우클릭 메뉴 공용)
        private void BtnDeleteBookmark_Click(object sender, RoutedEventArgs e)
        {
            // 1. 툴바 버튼 클릭 시: TreeView.SelectedItem 사용
            // 2. ContextMenu 클릭 시: DataContext 사용

            PdfBookmarkViewModel? targetBm = null;

            if (sender is MenuItem menuItem && menuItem.DataContext is PdfBookmarkViewModel contextBm)
            {
                targetBm = contextBm;
            }
            else
            {
                targetBm = BookmarkTree.SelectedItem as PdfBookmarkViewModel;
            }

            if (targetBm != null && SelectedDocument != null)
            {
                // 루트에서 삭제
                if (SelectedDocument.Bookmarks.Contains(targetBm))
                {
                    SelectedDocument.Bookmarks.Remove(targetBm);
                }
                // 부모의 자식 목록에서 삭제
                else if (targetBm.Parent != null)
                {
                    targetBm.Parent.Children.Remove(targetBm);
                }
            }
        }
        // =========================================================
        // [신규 구현] 책갈피 트리 구조 조작 (위/아래/들여쓰기/내어쓰기)
        // =========================================================

        // 1. 위로 이동
        private void BtnMoveBookmarkUp_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;

            // 현재 리스트(형제들) 찾기
            var currentList = GetCurrentCollection(selectedBm);
            if (currentList == null)
                return;

            int index = currentList.IndexOf(selectedBm);
            if (index > 0)
            {
                // UI 반영을 위해 Move 사용 (ObservableCollection 기능)
                currentList.Move(index, index - 1);
                // 포커스 유지
                selectedBm.IsSelected = true;
            }
        }

        // 2. 아래로 이동
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

        // 3. 들여쓰기 (자식으로 만들기: 바로 위 형제의 자식이 됨)
        private void BtnIndentBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;

            var currentList = GetCurrentCollection(selectedBm);
            if (currentList == null)
                return;

            int index = currentList.IndexOf(selectedBm);
            if (index > 0) // 바로 위에 형제가 있어야 함
            {
                var newParent = currentList[index - 1];

                // 이동 처리
                currentList.Remove(selectedBm);
                newParent.Children.Add(selectedBm);

                // 부모 참조 갱신
                selectedBm.Parent = newParent;

                // 새 부모 펼치기 및 선택 유지
                newParent.IsExpanded = true;
                selectedBm.IsSelected = true;
            }
        }

        // 4. 내어쓰기 (부모 레벨로 올리기: 현재 부모의 다음 형제가 됨)
        private void BtnOutdentBookmark_Click(object sender, RoutedEventArgs e)
        {
            if (BookmarkTree.SelectedItem is not PdfBookmarkViewModel selectedBm || SelectedDocument == null)
                return;

            // 이미 최상위 루트라면 내어쓰기 불가능
            if (selectedBm.Parent == null)
                return;

            var oldParent = selectedBm.Parent;

            // 이동할 목표 리스트 (할아버지의 자식들 or 최상위 루트)
            var targetList = (oldParent.Parent == null)
                ? SelectedDocument.Bookmarks
                : oldParent.Parent.Children;

            // 현재 부모가 목표 리스트에서 몇 번째인지 확인 (그 바로 밑으로 가야 함)
            int parentIndex = targetList.IndexOf(oldParent);
            if (parentIndex >= 0)
            {
                // 이동 처리
                oldParent.Children.Remove(selectedBm);
                targetList.Insert(parentIndex + 1, selectedBm);

                // 부모 참조 갱신
                selectedBm.Parent = oldParent.Parent;

                // 선택 유지
                selectedBm.IsSelected = true;
            }
        }

        // [헬퍼] 현재 아이템이 속한 컬렉션(형제 리스트)을 찾아주는 메서드
        private System.Collections.ObjectModel.ObservableCollection<PdfBookmarkViewModel>? GetCurrentCollection(PdfBookmarkViewModel bm)
        {
            if (SelectedDocument == null)
                return null;

            if (bm.Parent == null)
            {
                // 부모가 없으면 최상위 루트 리스트
                return SelectedDocument.Bookmarks;
            }
            else
            {
                // 부모가 있으면 부모의 자식 리스트
                return bm.Parent.Children;
            }
        }

        // =========================================================
        // [핵심 수정] 파일 열기 통합 메서드 (순서 개선)
        // =========================================================
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

                    // [신규] 마지막으로 읽던 페이지로 이동
                    int lastPage = _historyService.GetLastPage(filePath);

                    // 유효한 페이지 범위인지 확인 (파일 내용이 바뀔 수 있으므로)
                    if (lastPage > 0 && lastPage < docModel.Pages.Count)
                    {
                        // 약간의 딜레이를 주어야 UI(ListView)가 준비된 후 이동됨
                        await Task.Delay(100);

                        var listView = GetVisualChild<ListView>(MainTabControl);
                        if (listView != null)
                        {
                            var targetPage = docModel.Pages[lastPage];
                            // [확실한 이동] 가상화된 리스트에서 강제 스크롤
                            listView.ScrollIntoView(targetPage);
                            TxtStatus.Text = $"이전 위치(p.{lastPage + 1})로 복원됨";
                        }
                    }
                    else
                    {
                        TxtStatus.Text = "준비 완료";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"파일 열기 실패: {ex.Message}");
                TxtStatus.Text = "오류 발생";
            }
        }

        // =========================================================
        // [수정] 열기 버튼도 위 통합 메서드를 사용하도록 변경
        // =========================================================
        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "PDF Files (*.pdf)|*.pdf"
            };

            if (dlg.ShowDialog() == true)
            {
                OpenPdfFromPath(dlg.FileName);
            }
        }

        // =========================================================
        // 드래그 앤 드롭 이벤트 핸들러
        // =========================================================
        private void Window_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void Window_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string file in files)
                {
                    OpenPdfFromPath(file);
                }
            }
        }

        // 1. 왼쪽 회전
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

                // 회전 시 가로/세로 스왑 (UI 레이아웃 갱신용)
                double temp = page.Width;
                page.Width = page.Height;
                page.Height = temp;
            }
        }

        // 2. 오른쪽 회전
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

        // 3. 페이지 삭제
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

                    // 인덱스 재정렬
                    for (int i = 0; i < SelectedDocument.Pages.Count; i++)
                    {
                        SelectedDocument.Pages[i].PageIndex = i;
                    }
                    TxtStatus.Text = "페이지 삭제됨 (저장 시 반영)";
                    // 페이지 정보 갱신을 위해 강제로 스크롤 이벤트 호출하거나 텍스트 갱신
                    TxtPageInfo.Text = $"{Math.Min(idx + 1, SelectedDocument.Pages.Count)} / {SelectedDocument.Pages.Count}";
                }
            }
        }

        // 4. 빈 페이지 추가
        private void BtnAddBlankPage_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedDocument == null)
                return;
            int idx = GetCurrentPageIndex();

            // 현재 페이지 뒤에 추가
            var newPage = new PdfPageViewModel
            {
                IsBlankPage = true,
                OriginalFilePath = null, // 빈 페이지는 원본 파일 없음
                Width = 595,  // A4 기본 너비 (포인트)
                Height = 842, // A4 기본 높이
                PdfPageWidthPoint = 595,
                PdfPageHeightPoint = 842,
                Rotation = 0,
                PageIndex = idx + 1,
                // ImageSource는 null (하얀색으로 나옴)
            };

            SelectedDocument.Pages.Insert(idx + 1, newPage);

            // 인덱스 재정렬
            for (int i = 0; i < SelectedDocument.Pages.Count; i++)
            {
                SelectedDocument.Pages[i].PageIndex = i;
            }
            TxtStatus.Text = "빈 페이지 추가됨";
            TxtPageInfo.Text = $"{idx + 2} / {SelectedDocument.Pages.Count}";
        }

    }
}