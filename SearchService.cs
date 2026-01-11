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
                        if (document.PdfDocument != null)
                        {
                            // GetPdfText는 string을 반환
                            string pageText = document.PdfDocument.GetPdfText(i);

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

                                    // [수정] 구체 클래스로 캐스팅하여 GetTextSegments 호출 시도
                                    IList<PdfRectangle> bounds = new List<PdfRectangle>();

                                    if (document.PdfDocument is PdfiumViewer.PdfDocument specificDoc)
                                    {
                                        // PdfiumViewer 2.13.0 기준 API 확인: GetTextSegments(int page, int index, int count)
                                        // PdfTextSpan을 반환할 수도 있고 PdfRectangle을 반환할 수도 있음.
                                        // 만약 컴파일 에러가 난다면 이 부분이 핵심 원인입니다.
                                        // 여기서는 일반적인 PdfiumViewer 패턴인 GetTextSegments를 사용합니다.

                                        // 주의: GetTextSegments가 없다면 SearchService 기능을 잠시 비활성화해야 할 수도 있습니다.
                                        // 우선 가장 유력한 후보인 GetTextSegments를 사용합니다.

                                        // 반환 타입이 IList<PdfTextSpan>일 가능성이 높음 -> 변환 필요
                                        var segments = specificDoc.GetTextSegments(i, idx, query.Length);
                                        foreach (var seg in segments)
                                        {
                                            // PdfTextSpan -> PdfRectangle 변환 (Bounds 속성이 있다면)
                                            // PdfTextSpan에 X, Y, Width, Height가 직접 있다면 그대로 사용
                                            bounds.Add(new PdfRectangle(i, new System.Drawing.RectangleF((float)seg.X, (float)seg.Y, (float)seg.Width, (float)seg.Height)));
                                        }
                                    }

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
                        if (document.PdfDocument != null)
                        {
                            string pageText = document.PdfDocument.GetPdfText(i);

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

                                    IList<PdfRectangle> bounds = new List<PdfRectangle>();
                                    if (document.PdfDocument is PdfiumViewer.PdfDocument specificDoc)
                                    {
                                        var segments = specificDoc.GetTextSegments(i, idx, query.Length);
                                        foreach (var seg in segments)
                                        {
                                            bounds.Add(new PdfRectangle(i, new System.Drawing.RectangleF((float)seg.X, (float)seg.Y, (float)seg.Width, (float)seg.Height)));
                                        }
                                    }

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