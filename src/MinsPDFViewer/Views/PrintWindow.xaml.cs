using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace MinsPDFViewer
{
    /// <summary>
    /// 인쇄 설정 + 미리보기. 미리보기는 실제 인쇄와 같은 배치 계산(<see cref="PrintLayout"/>)과 같은 렌더 옵션을 쓴다.
    /// 확인하면 <see cref="Options"/>에 결과가 들어간다 (인쇄는 호출한 쪽에서 백그라운드로).
    /// </summary>
    public partial class PrintWindow : Window
    {
        private readonly PrintSource _source;
        private readonly int _currentPageIndex;
        private PrinterSettings? _printerSettings;
        private PageSettings? _pageSettings;
        private List<int> _pages = new();
        private int _previewPos;
        private int _renderGeneration;
        private bool _initializing = true;
        private readonly DispatcherTimer _previewDebounce;

        public PrintOptions? Options { get; private set; }

        public PrintWindow(PrintSource source, int currentPageIndex, Window? owner)
        {
            _source = source;
            _currentPageIndex = Math.Max(0, Math.Min(currentPageIndex, source.PageCount - 1));
            InitializeComponent();
            if (owner != null)
                Owner = owner;

            _previewDebounce = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(120) };
            _previewDebounce.Tick += (_, _) =>
            {
                _previewDebounce.Stop();
                _ = UpdatePreviewAsync();
            };

            RbAll.Content = $"전체 ({source.PageCount}쪽)";
            RbCurrent.Content = $"현재 페이지 ({_currentPageIndex + 1})";
            ChipUnlocked.Visibility = source.PrintRestricted ? Visibility.Visible : Visibility.Collapsed;
            ChipEdits.Visibility = source.FromPendingEdits ? Visibility.Visible : Visibility.Collapsed;

            Loaded += (_, _) =>
            {
                LoadPrinters();
                _initializing = false;
                RefreshPages();
                // 현재 페이지가 범위에 있으면 거기서 미리보기 시작
                int pos = _pages.IndexOf(_currentPageIndex);
                _previewPos = pos >= 0 ? pos : 0;
                SchedulePreview();
            };
        }

        // ── 프린터·용지 ──────────────────────────────────────────────────

        private void LoadPrinters()
        {
            CbPrinter.Items.Clear();
            string? defaultPrinter = null;
            try
            {
                defaultPrinter = new PrinterSettings().PrinterName;
                foreach (string name in PrinterSettings.InstalledPrinters)
                    CbPrinter.Items.Add(name);
            }
            catch (Exception ex)
            {
                DebugLog.Append("PrintWindow", $"Printer enumeration failed: {ex.Message}");
            }

            if (CbPrinter.Items.Count == 0)
            {
                CbPrinter.Items.Add("프린터 없음");
                CbPrinter.SelectedIndex = 0;
                CbPrinter.IsEnabled = false;
                BtnProperties.IsEnabled = false;
                BtnPrint.IsEnabled = false;
                return;
            }

            int index = defaultPrinter != null ? CbPrinter.Items.IndexOf(defaultPrinter) : -1;
            CbPrinter.SelectedIndex = index >= 0 ? index : 0;
        }

        private void CbPrinter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!CbPrinter.IsEnabled || CbPrinter.SelectedItem is not string name)
                return;

            try
            {
                var settings = new PrinterSettings { PrinterName = name };
                if (!settings.IsValid)
                {
                    _printerSettings = null;
                    _pageSettings = null;
                    BtnPrint.IsEnabled = false;
                    CbPaper.ItemsSource = null;
                    SchedulePreview();
                    return;
                }

                _printerSettings = settings;
                _pageSettings = (PageSettings)settings.DefaultPageSettings.Clone();
                LoadPapers(_pageSettings.PaperSize);
                ChkGrayscale.IsChecked = ChkGrayscale.IsChecked == true || !settings.SupportsColor;
                ChkGrayscale.IsEnabled = settings.SupportsColor;
            }
            catch (Exception ex)
            {
                DebugLog.Append("PrintWindow", $"Printer '{name}' settings failed: {ex.Message}");
                _printerSettings = null;
                _pageSettings = null;
            }

            UpdatePrintButton();
            SchedulePreview();
        }

        private void LoadPapers(PaperSize? select)
        {
            var papers = new List<PaperSize>();
            try
            {
                if (_printerSettings != null)
                    papers.AddRange(_printerSettings.PaperSizes.Cast<PaperSize>().Where(p => p.Width > 0 && p.Height > 0));
            }
            catch
            {
            }

            CbPaper.ItemsSource = papers;
            PaperSize? match = null;
            if (select != null)
            {
                match = papers.FirstOrDefault(p => p.RawKind == select.RawKind && p.RawKind != 0 && p.RawKind != (int)PaperKind.Custom)
                        ?? papers.FirstOrDefault(p => string.Equals(p.PaperName, select.PaperName, StringComparison.OrdinalIgnoreCase))
                        ?? papers.FirstOrDefault(p => Math.Abs(p.Width - select.Width) <= 2 && Math.Abs(p.Height - select.Height) <= 2);
            }
            CbPaper.SelectedItem = match ?? papers.FirstOrDefault();
        }

        private void CbPaper_SelectionChanged(object sender, SelectionChangedEventArgs e) => SchedulePreview();

        private void BtnProperties_Click(object sender, RoutedEventArgs e)
        {
            if (_printerSettings == null || _pageSettings == null)
                return;

            if (CbPaper.SelectedItem is PaperSize paper)
                _pageSettings.PaperSize = paper;
            if (RbLandscape.IsChecked == true) _pageSettings.Landscape = true;
            else if (RbPortrait.IsChecked == true) _pageSettings.Landscape = false;
            _printerSettings.Copies = ParseCopies() ?? 1;

            try
            {
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (!PrinterPropertiesDialog.Show(hwnd, _printerSettings, _pageSettings))
                    return;
            }
            catch (Exception ex)
            {
                DebugLog.Append("PrintWindow", $"Printer properties failed: {ex.Message}");
                return;
            }

            // 드라이버 창에서 바꾼 값을 화면에 반영
            LoadPapers(_pageSettings.PaperSize);
            TxtCopies.Text = Math.Max((short)1, _printerSettings.Copies).ToString(CultureInfo.InvariantCulture);
            ChkCollate.IsChecked = _printerSettings.Collate;
            if (_printerSettings.SupportsColor)
                ChkGrayscale.IsChecked = !_pageSettings.Color;
            if (RbAuto.IsChecked != true)
            {
                RbLandscape.IsChecked = _pageSettings.Landscape;
                RbPortrait.IsChecked = !_pageSettings.Landscape;
            }
            SchedulePreview();
        }

        // ── 옵션 ─────────────────────────────────────────────────────────

        private void Option_Changed(object sender, RoutedEventArgs e)
        {
            if (_initializing)
                return;
            if (ReferenceEquals(sender, RbAll) || ReferenceEquals(sender, RbCurrent) || ReferenceEquals(sender, RbRange))
                RefreshPages(resetPosition: true);
            SchedulePreview();
        }

        private void TxtRange_GotFocus(object sender, RoutedEventArgs e)
        {
            RbRange.IsChecked = true;
        }

        private void TxtRange_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_initializing)
                return;
            if (RbRange.IsChecked != true)
                RbRange.IsChecked = true;
            RefreshPages(resetPosition: true);
            SchedulePreview();
        }

        private void TxtCopies_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_initializing)
                return;
            bool ok = ParseCopies() != null;
            TxtCopies.BorderBrush = ok ? SystemColors.ControlDarkBrush : Brushes.IndianRed;
            UpdatePrintButton();
            UpdateSummary();
        }

        private short? ParseCopies()
        {
            if (int.TryParse(TxtCopies.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) && n >= 1 && n <= 999)
                return (short)n;
            return null;
        }

        private void RefreshPages(bool resetPosition = false)
        {
            string? error = null;
            if (RbCurrent.IsChecked == true)
            {
                _pages = new List<int> { _currentPageIndex };
            }
            else if (RbRange.IsChecked == true)
            {
                if (!PrintPageRange.TryParse(TxtRange.Text, _source.PageCount, out var parsed, out error))
                    parsed = new List<int>();
                _pages = parsed;
            }
            else
            {
                _pages = Enumerable.Range(0, _source.PageCount).ToList();
            }

            bool rangeError = RbRange.IsChecked == true && error != null && !string.IsNullOrWhiteSpace(TxtRange.Text);
            TxtRangeError.Text = rangeError ? error : "";
            TxtRangeError.Visibility = rangeError ? Visibility.Visible : Visibility.Collapsed;
            TxtRange.BorderBrush = rangeError ? Brushes.IndianRed : SystemColors.ControlDarkBrush;

            if (resetPosition)
                _previewPos = 0;
            _previewPos = Math.Max(0, Math.Min(_previewPos, _pages.Count - 1));
            UpdatePrintButton();
            UpdateSummary();
        }

        private void UpdatePrintButton()
        {
            BtnPrint.IsEnabled = _printerSettings != null && _pages.Count > 0 && ParseCopies() != null;
        }

        private void UpdateSummary()
        {
            short copies = ParseCopies() ?? 1;
            TxtSummary.Text = _pages.Count == 0
                ? ""
                : copies > 1 ? $"{_pages.Count}쪽 × {copies}부" : $"{_pages.Count}쪽";
            TxtPreviewPos.Text = _pages.Count == 0 ? "0 / 0" : $"{_previewPos + 1} / {_pages.Count}";
            BtnPrev.IsEnabled = _previewPos > 0;
            BtnNext.IsEnabled = _previewPos < _pages.Count - 1;
        }

        private PrintScaleMode SelectedScaleMode =>
            RbShrink.IsChecked == true ? PrintScaleMode.ShrinkOversized :
            RbActual.IsChecked == true ? PrintScaleMode.ActualSize : PrintScaleMode.Fit;

        private PrintOrientationMode SelectedOrientation =>
            RbPortrait.IsChecked == true ? PrintOrientationMode.Portrait :
            RbLandscape.IsChecked == true ? PrintOrientationMode.Landscape : PrintOrientationMode.Auto;

        // ── 미리보기 ─────────────────────────────────────────────────────

        private void BtnPrev_Click(object sender, RoutedEventArgs e) => MovePreview(-1);
        private void BtnNext_Click(object sender, RoutedEventArgs e) => MovePreview(1);

        private void MovePreview(int delta)
        {
            int next = Math.Max(0, Math.Min(_previewPos + delta, _pages.Count - 1));
            if (next == _previewPos)
                return;
            _previewPos = next;
            UpdateSummary();
            _ = UpdatePreviewAsync();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (Keyboard.FocusedElement is TextBox)
                return;
            if (e.Key is Key.PageDown or Key.Right or Key.Down) { MovePreview(1); e.Handled = true; }
            else if (e.Key is Key.PageUp or Key.Left or Key.Up) { MovePreview(-1); e.Handled = true; }
            else if (e.Key == Key.Home) { MovePreview(-_pages.Count); e.Handled = true; }
            else if (e.Key == Key.End) { MovePreview(_pages.Count); e.Handled = true; }
        }

        private void PreviewHost_SizeChanged(object sender, SizeChangedEventArgs e) => SchedulePreview();

        private void SchedulePreview()
        {
            if (_initializing)
                return;
            UpdateSummary();
            _previewDebounce.Stop();
            _previewDebounce.Start();
        }

        /// <summary>미리보기용 용지 크기(1/100 인치, 방향 적용)와 인쇄 가능 영역.</summary>
        private (double PaperW, double PaperH, PrintRect Printable) GetPaperGeometry(bool landscape)
        {
            double w = 850, h = 1100; // Letter 기본값
            double mx = 0, my = 0;
            if (CbPaper.SelectedItem is PaperSize paper && paper.Width > 0 && paper.Height > 0)
            {
                w = paper.Width;
                h = paper.Height;
            }

            if (_pageSettings != null)
            {
                try
                {
                    mx = _pageSettings.HardMarginX;
                    my = _pageSettings.HardMarginY;
                }
                catch
                {
                    mx = my = 0;
                }
            }

            if (landscape)
            {
                (w, h) = (h, w);
                (mx, my) = (my, mx);
            }

            return (w, h, new PrintRect(mx, my, Math.Max(1, w - 2 * mx), Math.Max(1, h - 2 * my)));
        }

        private async Task UpdatePreviewAsync()
        {
            if (_pages.Count == 0)
            {
                PreviewImage.Source = null;
                PreviewHost.Visibility = Visibility.Hidden;
                TxtNoPreview.Visibility = Visibility.Visible;
                return;
            }

            PreviewHost.Visibility = Visibility.Visible;
            TxtNoPreview.Visibility = Visibility.Collapsed;

            int pageIndex = _pages[_previewPos];
            var size = _source.PageSizes[pageIndex];
            bool defaultLandscape = _pageSettings?.Landscape ?? false;
            bool landscape = PrintLayout.IsLandscape(size.Width, size.Height, SelectedOrientation, defaultLandscape);
            var (paperW, paperH, printable) = GetPaperGeometry(landscape);
            var rect = PrintLayout.Compute(size.Width, size.Height, paperW, paperH, printable, SelectedScaleMode);

            PaperCanvas.Width = paperW;
            PaperCanvas.Height = paperH;
            PaperCanvas.Clip = new RectangleGeometry(new Rect(0, 0, paperW, paperH)); // 실제 크기에서 넘치는 부분은 잘림
            foreach (FrameworkElement element in new FrameworkElement[] { PreviewImage, PageOutline })
            {
                Canvas.SetLeft(element, rect.X);
                Canvas.SetTop(element, rect.Y);
                element.Width = rect.Width;
                element.Height = rect.Height;
            }

            // 화면에 실제로 보이는 크기만큼만 그린다
            double hostW = Math.Max(1, PreviewHost.ActualWidth);
            double hostH = Math.Max(1, PreviewHost.ActualHeight);
            double viewScale = Math.Min(hostW / paperW, hostH / paperH);
            double dpiScale = VisualTreeHelper.GetDpi(this).DpiScaleX;
            int pixelW = (int)Math.Clamp(Math.Round(rect.Width * viewScale * dpiScale), 1, 3000);
            int pixelH = (int)Math.Clamp(Math.Round(rect.Height * viewScale * dpiScale), 1, 3000);

            bool annotations = ChkAnnotations.IsChecked == true;
            bool grayscale = ChkGrayscale.IsChecked == true;
            int generation = ++_renderGeneration;
            try
            {
                var bitmap = await Task.Run(() => PrintService.RenderPreviewPage(_source, pageIndex, pixelW, pixelH, annotations, grayscale));
                if (generation != _renderGeneration)
                    return;
                PreviewImage.Source = bitmap;
            }
            catch (Exception ex)
            {
                DebugLog.Append("PrintWindow", $"Preview render failed (page {pageIndex + 1}): {ex.Message}");
            }
        }

        // ── 인쇄 ─────────────────────────────────────────────────────────

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            RefreshPages();
            short? copies = ParseCopies();
            if (_printerSettings == null || _pages.Count == 0 || copies == null || CbPrinter.SelectedItem is not string printer)
                return;

            var settings = (PrinterSettings)_printerSettings.Clone();
            Options = new PrintOptions
            {
                PrinterName = printer,
                PrinterSettings = settings,
                PaperSize = CbPaper.SelectedItem as PaperSize,
                Copies = copies.Value,
                Collate = ChkCollate.IsChecked == true,
                PageIndices = new List<int>(_pages),
                ScaleMode = SelectedScaleMode,
                Orientation = SelectedOrientation,
                Grayscale = ChkGrayscale.IsChecked == true,
                IncludeAnnotations = ChkAnnotations.IsChecked == true
            };
            DialogResult = true;
        }
    }
}
