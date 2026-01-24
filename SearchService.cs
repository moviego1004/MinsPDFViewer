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
        private int _currentPageIndex = 0;
        private int _currentMatchIndex = -1;
        private List<Rect> _currentMatches = new List<Rect>();
        private string _lastQuery = "";
        private PdfDocumentModel? _lastDocument = null;
        private bool _cancelSearch = false;
        
        // Highlights
        private readonly SolidColorBrush _brushActive = new SolidColorBrush(Color.FromArgb(120, 255, 165, 0)); // Orange
        private readonly SolidColorBrush _brushNormal = new SolidColorBrush(Color.FromArgb(80, 255, 255, 0)); // Yellow

        public SearchService(PdfService pdfService)
        {
            _pdfService = pdfService;
            _brushActive.Freeze();
            _brushNormal.Freeze();
        }

        public void ResetSearch()
        {
            _currentPageIndex = 0;
            _currentMatchIndex = -1;
            _currentMatches.Clear();
            _cancelSearch = false;
        }

        public void CancelSearch()
        {
            _cancelSearch = true;
        }

        private void ClearAllHighlights(PdfDocumentModel document)
        {
            Application.Current.Dispatcher.Invoke(() =>
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
            if (string.IsNullOrWhiteSpace(query) || document == null) return null;

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
                if (_currentMatchIndex + 1 < _currentMatches.Count)
                {
                    _currentMatchIndex++;
                    return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                }

                // 2. Search next pages
                for (int i = 0; i < pageCount; i++)
                {
                    if (_cancelSearch) break;

                    // Calculate next page index (wrap around logic can be added if needed, currently linear)
                    int targetPageIdx = (_currentPageIndex + 1 + i);
                    if (targetPageIdx >= pageCount) 
                    {
                         // Wrap around or stop? Let's stop at end for now, or wrap?
                         // User requirement: "move to next page". Usually wraps or stops.
                         // Let's stop and let UI handle "End of document".
                         break; 
                    }
                    
                    // Actually, let's fix the loop logic. We need to start from next page.
                    // If we are at page 0, we searched matches. If exhausted, go to page 1.
                    // Logic above: `_currentPageIndex` is where we are.
                    
                    // Reset matches for new page
                    _currentMatches.Clear();
                    _currentMatchIndex = -1;
                    _currentPageIndex = targetPageIdx; // Move to next page

                    var rects = _pdfService.FindTextRects(document.FilePath, _currentPageIndex, query);
                    
                    if (rects.Count > 0)
                    {
                        _currentMatches = rects;
                        _currentMatchIndex = 0;
                        return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                    }
                }

                return null;
            });
        }
        
        public async Task<PdfAnnotation?> FindPrevAsync(PdfDocumentModel document, string query)
        {
            // Similar logic but backwards
             if (string.IsNullOrWhiteSpace(query) || document == null) return null;

            if (query != _lastQuery || document != _lastDocument)
            {
                _lastDocument = document;
                _lastQuery = query;
                _currentPageIndex = document.Pages.Count; // Start from end
                _currentMatchIndex = -1;
                _currentMatches.Clear();
                ClearAllHighlights(document);
            }

            return await Task.Run(() =>
            {
                // 1. Check if we have prev matches in current page
                if (_currentMatchIndex - 1 >= 0)
                {
                    _currentMatchIndex--;
                    return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                }

                // 2. Search prev pages
                int pageCount = document.Pages.Count;
                // If we are starting fresh (FindPrev from scratch), _currentPageIndex might be count.
                // We need to loop backwards.
                
                int startIdx = _currentPageIndex;
                if (_currentMatches.Count == 0) startIdx = _currentPageIndex - 1; // Move to prev page if current exhausted
                
                for (int i = startIdx; i >= 0; i--)
                {
                    if (_cancelSearch) break;
                    
                    _currentPageIndex = i;
                    var rects = _pdfService.FindTextRects(document.FilePath, _currentPageIndex, query);
                    
                    if (rects.Count > 0)
                    {
                        _currentMatches = rects;
                        _currentMatchIndex = rects.Count - 1; // Last match in page
                        return HighlightMatch(document, _currentPageIndex, _currentMatchIndex);
                    }
                }
                return null;
            });
        }

        private PdfAnnotation? HighlightMatch(PdfDocumentModel document, int pageIndex, int matchIndex)
        {
            if (pageIndex < 0 || pageIndex >= document.Pages.Count) return null;
            if (matchIndex < 0 || matchIndex >= _currentMatches.Count) return null;

            var rect = _currentMatches[matchIndex];
            var pageVM = document.Pages[pageIndex];

            // Convert PDF coords to UI coords
            // PDF: Origin Bottom-Left. Y increases Up.
            // UI: Origin Top-Left. Y increases Down.
            
            // However, FPDFText_GetRect returns page coordinates.
            // We need page height to flip Y.
            
            double pdfH = pageVM.PdfPageHeightPoint;
            double pdfW = pageVM.PdfPageWidthPoint;
            
            // Rect from FPDFText_GetRect is (left, top, right, bottom) in PDF coords.
            // Width = right - left
            // Height = top - bottom
            // UI X = left * Scale
            // UI Y = (PageHeight - top) * Scale
            
            double uiScale = pageVM.Width / pdfW; // Current UI scale relative to PDF point size
            
            // Since pageVM.Width includes Zoom, we should use that.
            // Note: pageVM.Width is the actual display width (e.g. 1000px).
            // pdfW is 595pt (A4).
            
            double uiX = rect.X * uiScale;
            double uiY = (pdfH - rect.Y) * uiScale; 
            // Wait, rect.Y in my FindTextRects implementation was "top". 
            // FPDFText_GetRect returns (left, top, right, bottom).
            // I stored Rect(left, top, width, height).
            // So rect.Y is PDF Top.
            // UI Y = (pdfH - pdfTop) * scale.
            
            double uiW = rect.Width * uiScale;
            double uiH = rect.Height * uiScale; // Height was (top - bottom)
            
            PdfAnnotation? annot = null;
            
            Application.Current.Dispatcher.Invoke(() =>
            {
                // 1. Remove ONLY the active highlight from ALL pages (to move focus)
                // Or better: Mark all existing search highlights as Normal, then add/update Active.
                
                // Let's remove previous active highlights
                foreach(var p in document.Pages) 
                {
                    var actives = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight && object.ReferenceEquals(a.Background, _brushActive)).ToList();
                    foreach(var a in actives) p.Annotations.Remove(a);
                    
                    // Also clear "Normal" highlights if we want fresh start? No, keep them?
                    // User wants: "Highlight words in page, and active one different color".
                    // So we should add ALL matches as Normal, then overlay Active?
                    // For now, simpler: Just show current match as Active.
                    // If user wants ALL matches highlighted simultaneously:
                    // When loading a page, we should add annotations for ALL matches.
                    
                    // Implementing "Highlight ALL in current page + Active":
                    // Whenever we change page, we add all matches as Normal.
                }

                // Add Normal highlights for all matches in current page
                // Check if already added?
                // For simplicity/performance: Just show current active match.
                // Re-reading user request: "텍스트 검색하면 페이지 내에서 단어를 하이라이트해주고, 현재 포커스 단어는 다른 색상으로..."
                // => All matches in page should be highlighted.
                
                // Add all matches as Normal first (if not exists)
                foreach(var r in _currentMatches)
                {
                    double rX = r.X * uiScale;
                    double rY = (pdfH - r.Y) * uiScale;
                    double rW = r.Width * uiScale;
                    double rH = r.Height * uiScale;
                    
                    // Check existence? O(N^2) but N is small (matches in page)
                    bool exists = pageVM.Annotations.Any(a => a.Type == AnnotationType.SearchHighlight && Math.Abs(a.X - rX) < 1 && Math.Abs(a.Y - rY) < 1);
                    if (!exists)
                    {
                         pageVM.Annotations.Add(new PdfAnnotation
                         {
                             Type = AnnotationType.SearchHighlight,
                             X = rX, Y = rY, Width = rW, Height = rH,
                             Background = _brushNormal
                         });
                    }
                }
                
                // Add Active highlight (overlay)
                annot = new PdfAnnotation
                {
                    Type = AnnotationType.SearchHighlight,
                    X = uiX, Y = uiY, Width = uiW, Height = uiH,
                    Background = _brushActive
                };
                pageVM.Annotations.Add(annot);
            });
            
            return annot;
        }
    }
}