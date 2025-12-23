using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace MinsPDFViewer
{
    public partial class MainWindow : Window
    {
        // 서비스 인스턴스
        private readonly PdfService _pdfService;
        private readonly SearchService _searchService;

        public System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel> Documents { get; set; } 
            = new System.Collections.ObjectModel.ObservableCollection<PdfDocumentModel>();
        
        private PdfDocumentModel? _selectedDocument;
        public PdfDocumentModel? SelectedDocument
        {
            get => _selectedDocument;
            set 
            { 
                _selectedDocument = value; 
                // 탭 변경 시 툴바 상태 등 업데이트 필요시 호출
            }
        }

        // 검색 상태 관리
        private List<PdfAnnotation> _searchResults = new List<PdfAnnotation>();
        private int _currentSearchIndex = -1;
        private string _lastSearchQuery = "";

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            
            // 서비스 초기화
            _pdfService = new PdfService();
            _searchService = new SearchService();
            
            // 폰트 리졸버 설정 (한 번만 실행)
            try { 
                if (PdfSharp.Fonts.GlobalFontSettings.FontResolver == null) 
                    PdfSharp.Fonts.GlobalFontSettings.FontResolver = new WindowsFontResolver(); 
            } catch { }
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "PDF Files|*.pdf" };
            if (dlg.ShowDialog() == true)
            {
                var docModel = _pdfService.LoadPdf(dlg.FileName);
                if (docModel != null)
                {
                    Documents.Add(docModel);
                    SelectedDocument = docModel;
                    // 비동기 렌더링 시작
                    _ = _pdfService.RenderPagesAsync(docModel);
                }
            }
        }

        private void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            DoSearch(TxtSearch.Text);
        }

        private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                bool isShift = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
                if (TxtSearch.Text == _lastSearchQuery && _searchResults.Count > 0)
                {
                    NavigateSearchResult(!isShift);
                }
                else
                {
                    DoSearch(TxtSearch.Text);
                }
            }
        }

        private void DoSearch(string query)
        {
            if (SelectedDocument == null) return;

            _lastSearchQuery = query;
            // SearchService에 위임
            _searchResults = _searchService.PerformSearch(SelectedDocument, query);
            
            TxtStatus.Text = $"검색 결과: {_searchResults.Count}건";
            
            if (_searchResults.Count > 0)
            {
                _currentSearchIndex = -1;
                NavigateSearchResult(true);
            }
            else
            {
                MessageBox.Show("검색 결과가 없습니다.");
            }
        }

        private void NavigateSearchResult(bool next)
        {
            if (_searchResults.Count == 0) return;

            // 이전 강조 해제
            if (_currentSearchIndex >= 0 && _currentSearchIndex < _searchResults.Count)
                _searchResults[_currentSearchIndex].Background = new SolidColorBrush(Color.FromArgb(60, 0, 255, 255));

            // 인덱스 이동
            if (next)
            {
                _currentSearchIndex++;
                if (_currentSearchIndex >= _searchResults.Count) _currentSearchIndex = 0;
            }
            else
            {
                _currentSearchIndex--;
                if (_currentSearchIndex < 0) _currentSearchIndex = _searchResults.Count - 1;
            }

            // 현재 항목 강조
            var currentAnnot = _searchResults[_currentSearchIndex];
            currentAnnot.Background = new SolidColorBrush(Color.FromArgb(120, 255, 0, 255));

            // 스크롤 이동
            if (SelectedDocument != null)
            {
                var targetPage = SelectedDocument.Pages.FirstOrDefault(p => p.Annotations.Contains(currentAnnot));
                if (targetPage != null)
                {
                    // 탭 컨트롤 내부의 ListView 찾기 (VisualTreeHelper 필요)
                    var listView = GetVisualChild<ListView>(MainTabControl);
                    if (listView != null)
                    {
                        listView.ScrollIntoView(targetPage);
                    }
                    TxtStatus.Text = $"검색: {_currentSearchIndex + 1} / {_searchResults.Count}";
                }
            }
        }
        
        // 유틸리티: VisualTree에서 자식 찾기
        private static T? GetVisualChild<T>(DependencyObject parent) where T : Visual
        {
            if (parent == null) return null;
            T? child = default(T);
            int numVisuals = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < numVisuals; i++) {
                Visual v = (Visual)VisualTreeHelper.GetChild(parent, i);
                child = v as T;
                if (child == null) child = GetVisualChild<T>(v);
                if (child != null) break;
            }
            return child;
        }

        // 탭 닫기
        private void BtnCloseTab_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is PdfDocumentModel doc)
            {
                doc.DocReader?.Dispose();
                Documents.Remove(doc);
                if (Documents.Count == 0) SelectedDocument = null;
            }
        }
        
        // 줌 및 기타 버튼 이벤트 핸들러들은 여기에 그대로 유지하거나 필요 시 추가...
        private void BtnPrevSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(false);
        private void BtnNextSearch_Click(object sender, RoutedEventArgs e) => NavigateSearchResult(true);
    }
}