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

            // 1. Search text layer via PDFium
            if (!string.IsNullOrEmpty(document.FilePath))
            {
                var textRects = _pdfService.FindTextRects(document.FilePath, pageVM.OriginalPageIndex, query, document.Password);
                if (textRects != null && textRects.Count > 0)
                {
                    results.AddRange(textRects);
                }
            }

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

            // OcrWords는 이미 PDF 포인트(원본 페이지 화면 기준, 왼쪽 위 원점)이므로 줌 환산이 필요 없다.
            // 반환 형식은 FindTextRects와 같게 (left, top-from-bottom, width, height).
            double pdfH = pageVM.SourceViewHeightPoint > 0 ? pageVM.SourceViewHeightPoint : 842.0;
            const double scaleX = 1.0;
            const double scaleY = 1.0;

            // Single word check
            string trimmedQuery = query.Trim();
            foreach (var word in pageVM.OcrWords)
            {
                if (word.Text.IndexOf(trimmedQuery, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    double left = word.BoundingBox.X * scaleX;
                    double width = word.BoundingBox.Width * scaleX;
                    double top = pdfH - (word.BoundingBox.Y * scaleY);
                    double height = word.BoundingBox.Height * scaleY;
                    rects.Add(new Rect(left, top, width, height));
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

                            double left = minX * scaleX;
                            double width = (maxX - minX) * scaleX;
                            double top = pdfH - (minY * scaleY);
                            double height = (maxY - minY) * scaleY;
                            rects.Add(new Rect(left, top, width, height));
                            break;
                        }
                        j++;
                    }
                }
            }

            return rects;
        }

        private PdfAnnotation? HighlightMatch(PdfDocumentModel document, int pageIndex, int matchIndex)
        {
            if (pageIndex < 0 || pageIndex >= document.Pages.Count) return null;
            if (matchIndex < 0 || matchIndex >= _currentMatches.Count) return null;

            var rect = _currentMatches[matchIndex];
            var pageVM = document.Pages[pageIndex];

            double pdfH = pageVM.PdfPageHeightPoint > 0 ? pageVM.PdfPageHeightPoint : 842.0;
            double pdfW = pageVM.PdfPageWidthPoint > 0 ? pageVM.PdfPageWidthPoint : 595.0;

            double uiScale = pageVM.Width > 0 ? pageVM.Width / pdfW : 1.0;

            double uiX = rect.X * uiScale;
            double uiY = (pdfH - rect.Y) * uiScale;
            double uiW = rect.Width * uiScale;
            double uiH = rect.Height * uiScale;

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
                    double rX = r.X * uiScale;
                    double rY = (pdfH - r.Y) * uiScale;
                    double rW = r.Width * uiScale;
                    double rH = r.Height * uiScale;

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
