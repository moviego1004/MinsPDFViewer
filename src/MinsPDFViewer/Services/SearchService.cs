using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace MinsPDFViewer
{
    public class SearchService
    {
        private readonly PdfService _pdfService;
        private int _currentPageIndex = -1;
        private int _currentMatchIndex = -1;
        private List<Rect> _currentMatches = new List<Rect>();
        private string _lastQuery = "";
        private PdfDocumentModel? _lastDocument = null;
        private bool _cancelSearch = false;

        // Highlights
        private readonly SolidColorBrush _brushActive = new SolidColorBrush(Color.FromArgb(140, 255, 140, 0)); // Dark Orange
        private readonly SolidColorBrush _brushNormal = new SolidColorBrush(Color.FromArgb(90, 255, 235, 59)); // Bright Yellow

        public SearchService(PdfService pdfService)
        {
            _pdfService = pdfService;
            _brushActive.Freeze();
            _brushNormal.Freeze();
        }

        private static void RunOnUiDispatcherOrDirect(Action action)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
            {
                action();
                return;
            }
            dispatcher.Invoke(action);
        }

        public void ResetSearch()
        {
            _currentPageIndex = -1;
            _currentMatchIndex = -1;
            _currentMatches.Clear();
            _cancelSearch = false;
        }

        public void CancelSearch()
        {
            _cancelSearch = true;
        }

        public void ClearAllHighlights(PdfDocumentModel document)
        {
            if (document == null) return;
            RunOnUiDispatcherOrDirect(() =>
            {
                foreach (var page in document.Pages)
                {
                    var toRemove = page.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                    foreach (var a in toRemove) page.Annotations.Remove(a);
                }
            });
        }

        public async Task<PdfAnnotation?> FindNextAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.Pages.Count == 0) return null;

            if (query != _lastQuery || document != _lastDocument)
            {
                _lastDocument = document;
                _lastQuery = query;
                ResetSearch();
                ClearAllHighlights(document);
            }

            return await Task.Run(() =>
            {
                int pageCount = document.Pages.Count;

                // 1. Check if we have more matches in current page
                if (_currentPageIndex >= 0 && _currentPageIndex < pageCount && _currentMatchIndex + 1 < _currentMatches.Count)
                {
                    _currentMatchIndex++;
                    return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                }

                // 2. Search subsequent pages sequentially
                int startPage = _currentPageIndex + 1;
                for (int pageIdx = startPage; pageIdx < pageCount; pageIdx++)
                {
                    if (_cancelSearch) break;

                    var pageVM = document.Pages[pageIdx];
                    var rects = SearchInPage(document, pageVM, query);

                    if (rects.Count > 0)
                    {
                        _currentMatches = rects;
                        _currentMatchIndex = 0;
                        _currentPageIndex = pageIdx;
                        return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                    }
                }

                return null;
            });
        }

        public async Task<PdfAnnotation?> FindPrevAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.Pages.Count == 0) return null;

            if (query != _lastQuery || document != _lastDocument)
            {
                _lastDocument = document;
                _lastQuery = query;
                _currentPageIndex = document.Pages.Count; // Start after end
                _currentMatchIndex = -1;
                _currentMatches.Clear();
                _cancelSearch = false;
                ClearAllHighlights(document);
            }

            return await Task.Run(() =>
            {
                int pageCount = document.Pages.Count;

                // 1. Check if we have prev matches in current page
                if (_currentPageIndex >= 0 && _currentPageIndex < pageCount && _currentMatchIndex - 1 >= 0 && _currentMatchIndex - 1 < _currentMatches.Count)
                {
                    _currentMatchIndex--;
                    return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                }

                // 2. Search previous pages sequentially backwards
                int startIdx = _currentPageIndex - 1;
                if (startIdx >= pageCount) startIdx = pageCount - 1;

                for (int pageIdx = startIdx; pageIdx >= 0; pageIdx--)
                {
                    if (_cancelSearch) break;

                    var pageVM = document.Pages[pageIdx];
                    var rects = SearchInPage(document, pageVM, query);

                    if (rects.Count > 0)
                    {
                        _currentMatches = rects;
                        _currentMatchIndex = rects.Count - 1;
                        _currentPageIndex = pageIdx;
                        return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                    }
                }

                return null;
            });
        }

        private List<Rect> SearchInPage(PdfDocumentModel document, PdfPageViewModel pageVM, string query)
        {
            var results = new List<Rect>();
            if (pageVM.IsBlankPage) return results;

            // 1. 텍스트 층 (열린 문서에서 바로, 결과는 페이지 화면 좌표)
            var textRects = _pdfService.FindTextRectsInView(document, pageVM, query);
            if (textRects.Count > 0)
                results.AddRange(textRects);

            // 2. If no text rects found from PDF layer, fallback to in-memory OcrWords (if OCR was run)
            if (results.Count == 0 && pageVM.OcrWords != null && pageVM.OcrWords.Count > 0)
            {
                results.AddRange(SearchOcrWords(pageVM, query));
            }

            return results;
        }

        private static List<Rect> SearchOcrWords(PdfPageViewModel pageVM, string query)
        {
            var rects = new List<Rect>();
            if (pageVM.OcrWords == null || pageVM.OcrWords.Count == 0) return rects;

            // OcrWords는 원본 페이지 화면(앱 회전 전) 포인트, 왼쪽 위 원점 → 페이지 화면 좌표(DIP)로
            Rect ToView(double x, double y, double w, double h) => OcrRectToView(pageVM, new Rect(x, y, w, h));

            // Single word check
            string trimmedQuery = query.Trim();
            foreach (var word in pageVM.OcrWords)
            {
                if (word.Text.IndexOf(trimmedQuery, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    var box = word.BoundingBox;
                    rects.Add(ToView(box.X, box.Y, box.Width, box.Height));
                }
            }

            // Multi-word phrase check (across sequential words)
            if (rects.Count == 0 && trimmedQuery.Contains(' '))
            {
                var words = pageVM.OcrWords;
                for (int i = 0; i < words.Count; i++)
                {
                    string combined = "";
                    int j = i;
                    while (j < words.Count && combined.Length < trimmedQuery.Length * 2)
                    {
                        combined += (combined.Length == 0 ? "" : " ") + words[j].Text;
                        if (combined.IndexOf(trimmedQuery, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Merge bounding boxes from i to j
                            double minX = double.MaxValue, minY = double.MaxValue;
                            double maxX = double.MinValue, maxY = double.MinValue;
                            for (int k = i; k <= j; k++)
                            {
                                minX = Math.Min(minX, words[k].BoundingBox.X);
                                minY = Math.Min(minY, words[k].BoundingBox.Y);
                                maxX = Math.Max(maxX, words[k].BoundingBox.Right);
                                maxY = Math.Max(maxY, words[k].BoundingBox.Bottom);
                            }

                            rects.Add(ToView(minX, minY, maxX - minX, maxY - minY));
                            break;
                        }
                        j++;
                    }
                }
            }

            return rects;
        }

        /// <summary>원본 페이지 화면(포인트, 앱 회전 전) 사각형 → 현재 페이지 화면(DIP, 앱 회전·줌 적용).</summary>
        internal static Rect OcrRectToView(PdfPageViewModel pageVM, Rect source)
        {
            double sw = pageVM.SourceViewWidthPoint > 0 ? pageVM.SourceViewWidthPoint : 595.0;
            double sh = pageVM.SourceViewHeightPoint > 0 ? pageVM.SourceViewHeightPoint : 842.0;
            int quarter = PdfService.ToPdfiumRotate(pageVM.Rotation);
            bool turned = quarter % 2 == 1;
            double scale = pageVM.Width > 0 ? pageVM.Width / (turned ? sh : sw) : 1.0;

            (double X, double Y) Map(double x, double y) => quarter switch
            {
                1 => (sh - y, x),          // 시계 방향 90°
                2 => (sw - x, sh - y),
                3 => (y, sw - x),
                _ => (x, y)
            };

            var a = Map(source.Left, source.Top);
            var b = Map(source.Right, source.Bottom);
            return new Rect(
                Math.Min(a.X, b.X) * scale, Math.Min(a.Y, b.Y) * scale,
                Math.Abs(b.X - a.X) * scale, Math.Abs(b.Y - a.Y) * scale);
        }

        private PdfAnnotation? HighlightMatch(PdfDocumentModel document, int pageIndex, int matchIndex)
        {
            if (pageIndex < 0 || pageIndex >= document.Pages.Count) return null;
            if (matchIndex < 0 || matchIndex >= _currentMatches.Count) return null;

            // 결과는 이미 페이지 화면 좌표(DIP)
            var rect = _currentMatches[matchIndex];
            var pageVM = document.Pages[pageIndex];
            double uiX = rect.X, uiY = rect.Y, uiW = rect.Width, uiH = rect.Height;

            PdfAnnotation? activeAnnot = null;

            RunOnUiDispatcherOrDirect(() =>
            {
                // Remove previous active highlights from all pages
                foreach (var p in document.Pages)
                {
                    var actives = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight && object.ReferenceEquals(a.Background, _brushActive)).ToList();
                    foreach (var a in actives) p.Annotations.Remove(a);
                }

                // Add Normal highlights for all matches in the current page
                foreach (var r in _currentMatches)
                {
                    double rX = r.X, rY = r.Y, rW = r.Width, rH = r.Height;

                    bool exists = pageVM.Annotations.Any(a => a.Type == AnnotationType.SearchHighlight &&
                                                              Math.Abs(a.X - rX) < 1 &&
                                                              Math.Abs(a.Y - rY) < 1);
                    if (!exists)
                    {
                        pageVM.Annotations.Add(new PdfAnnotation
                        {
                            Type = AnnotationType.SearchHighlight,
                            X = rX,
                            Y = rY,
                            Width = rW,
                            Height = rH,
                            Background = _brushNormal
                        });
                    }
                }

                // Add active focus highlight
                activeAnnot = new PdfAnnotation
                {
                    Type = AnnotationType.SearchHighlight,
                    X = uiX,
                    Y = uiY,
                    Width = uiW,
                    Height = uiH,
                    Background = _brushActive
                };
                pageVM.Annotations.Add(activeAnnot);
            });

            return activeAnnot;
        }
    }
}
