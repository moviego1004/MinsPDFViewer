using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using PdfiumViewer;

namespace MinsPDFViewer
{
    public class SearchService
    {
        private int _lastPageIndex = 0;
        private int _lastCharIndex = 0;
        private string _lastQuery = "";
        private PdfDocumentModel? _lastDocument = null;
        private bool _cancelSearch = false;

        private readonly SolidColorBrush _brushContext = new SolidColorBrush(Color.FromArgb(60, 0, 255, 0));
        private readonly SolidColorBrush _brushActive = new SolidColorBrush(Color.FromArgb(100, 255, 165, 0));

        public void ResetSearch()
        {
            _lastPageIndex = 0;
            _lastCharIndex = 0;
            _cancelSearch = false;
        }

        public void CancelSearch()
        {
            _cancelSearch = true;
        }

        public async Task<PdfAnnotation?> FindNextAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null)
                return null;

            if (query != _lastQuery || document != _lastDocument)
            {
                _lastDocument = document;
                _lastQuery = query;
                ResetSearch();
                ClearAllSearchHighlights(document);
            }

            return await Task.Run(() =>
            {
                int pageCount = 0;
                lock (PdfService.PdfiumLock)
                {
                    if (document.PdfDocument == null)
                        return (PdfAnnotation?)null;
                    pageCount = document.PdfDocument.PageCount;
                }

                for (int i = _lastPageIndex; i < pageCount; i++)
                {
                    if (_cancelSearch)
                        break;

                    var matchesOnPage = new List<SearchMatchInfo>();
                    int findIndex = -1;

                    lock (PdfService.PdfiumLock)
                    {
                        if (document.PdfDocument is PdfiumViewer.PdfDocument specificDoc)
                        {
                            // [수정] GetPdfText는 string을 반환함
                            string pageText = specificDoc.GetPdfText(i);

                            int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                            findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                            if (findIndex != -1)
                            {
                                int scanIndex = 0;
                                while (true)
                                {
                                    int idx = pageText.IndexOf(query, scanIndex, StringComparison.OrdinalIgnoreCase);
                                    if (idx == -1)
                                        break;

                                    // [수정] GetTextBounds는 Document에서 호출
                                    var bounds = specificDoc.GetTextBounds(i, idx, query.Length);
                                    bool isActive = (idx == findIndex);
                                    matchesOnPage.Add(new SearchMatchInfo { Rects = bounds, IsActive = isActive });

                                    scanIndex = idx + 1;
                                }
                            }
                        }
                    }

                    if (findIndex != -1 && matchesOnPage.Count > 0)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;

                        return Application.Current.Dispatcher.Invoke(() =>
                            RenderPageSearchResults(document, i, matchesOnPage));
                    }

                    if (i == _lastPageIndex)
                        _lastCharIndex = 0;
                }
                return null;
            });
        }

        public async Task<PdfAnnotation?> FindPrevAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null)
                return null;

            if (query != _lastQuery || document != _lastDocument)
            {
                _lastDocument = document;
                _lastQuery = query;
                _lastPageIndex = document.Pages.Count - 1;
                _lastCharIndex = -1;
                ClearAllSearchHighlights(document);
            }

            return await Task.Run(() =>
            {
                for (int i = _lastPageIndex; i >= 0; i--)
                {
                    var matchesOnPage = new List<SearchMatchInfo>();
                    int findIndex = -1;

                    lock (PdfService.PdfiumLock)
                    {
                        if (document.PdfDocument is PdfiumViewer.PdfDocument specificDoc)
                        {
                            // [수정] GetPdfText는 string을 반환
                            string pageText = specificDoc.GetPdfText(i);

                            int startIndex = (i == _lastPageIndex) ? ((_lastCharIndex == -1 || _lastCharIndex == 0) ? pageText.Length - 1 : Math.Max(0, _lastCharIndex - 2)) : pageText.Length - 1;
                            if (startIndex < 0 || startIndex >= pageText.Length)
                                startIndex = pageText.Length - 1;

                            findIndex = pageText.LastIndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                            if (findIndex != -1)
                            {
                                int scanIndex = 0;
                                while (true)
                                {
                                    int idx = pageText.IndexOf(query, scanIndex, StringComparison.OrdinalIgnoreCase);
                                    if (idx == -1)
                                        break;

                                    // [수정] GetTextBounds는 Document에서 호출
                                    var bounds = specificDoc.GetTextBounds(i, idx, query.Length);
                                    bool isActive = (idx == findIndex);
                                    matchesOnPage.Add(new SearchMatchInfo { Rects = bounds, IsActive = isActive });

                                    scanIndex = idx + 1;
                                }
                            }
                        }
                    }

                    if (findIndex != -1 && matchesOnPage.Count > 0)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex;
                        return Application.Current.Dispatcher.Invoke(() =>
                            RenderPageSearchResults(document, i, matchesOnPage));
                    }
                }
                return null;
            });
        }

        private void ClearAllSearchHighlights(PdfDocumentModel document)
        {
            foreach (var p in document.Pages)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                    foreach (var r in toRemove)
                        p.Annotations.Remove(r);
                });
            }
        }

        private PdfAnnotation? RenderPageSearchResults(PdfDocumentModel doc, int pageIndex, List<SearchMatchInfo> matches)
        {
            var pageVM = doc.Pages[pageIndex];

            var toRemove = pageVM.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
            foreach (var r in toRemove)
                pageVM.Annotations.Remove(r);

            PdfAnnotation? activeAnnotation = null;

            foreach (var match in matches)
            {
                foreach (var rect in match.Rects)
                {
                    var bounds = rect.Bounds;
                    var ann = new PdfAnnotation
                    {
                        X = bounds.X,
                        Y = bounds.Y,
                        Width = bounds.Width,
                        Height = bounds.Height,
                        Type = AnnotationType.SearchHighlight
                    };

                    if (match.IsActive)
                    {
                        ann.Background = _brushActive;
                        activeAnnotation = ann;
                    }
                    else
                    {
                        ann.Background = _brushContext;
                    }

                    pageVM.Annotations.Add(ann);
                }
            }

            return activeAnnotation;
        }

        // [내부 클래스] 매치 정보 전달용 DTO
        private class SearchMatchInfo
        {
            public IList<PdfRectangle> Rects { get; set; } = new List<PdfRectangle>();
            public bool IsActive
            {
                get; set;
            }
        }
    }
}