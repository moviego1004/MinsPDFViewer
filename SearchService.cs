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

        // 하이라이팅용 브러시는 당분간 사용 안 함
        // private readonly SolidColorBrush _brushContext = new SolidColorBrush(Color.FromArgb(60, 0, 255, 0));
        // private readonly SolidColorBrush _brushActive = new SolidColorBrush(Color.FromArgb(100, 255, 165, 0));

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
                // ClearAllSearchHighlights(document); // 하이라이트 기능 제거
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

                    int findIndex = -1;

                    lock (PdfService.PdfiumLock)
                    {
                        if (document.PdfDocument != null)
                        {
                            // 텍스트 가져오기
                            string pageText = document.PdfDocument.GetPdfText(i);

                            int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                            findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);
                        }
                    }

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;

                        // 하이라이트 없이 페이지 정보만 담은 더미 주석 반환
                        // MainWindow에서 이 주석이 포함된 페이지로 스크롤함
                        var dummyAnnot = new PdfAnnotation { Type = AnnotationType.SearchHighlight };

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            document.Pages[i].Annotations.Add(dummyAnnot);
                        });

                        return dummyAnnot;
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
                // ClearAllSearchHighlights(document);
            }

            return await Task.Run(() =>
            {
                for (int i = _lastPageIndex; i >= 0; i--)
                {
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
                        }
                    }

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex;

                        var dummyAnnot = new PdfAnnotation { Type = AnnotationType.SearchHighlight };

                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            document.Pages[i].Annotations.Add(dummyAnnot);
                        });

                        return dummyAnnot;
                    }
                }
                return null;
            });
        }

        // 하이라이트 기능 임시 제거로 인해 사용 안 함
        /*
        private void ClearAllSearchHighlights(PdfDocumentModel document)
        {
            foreach (var p in document.Pages)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                    foreach (var r in toRemove) p.Annotations.Remove(r);
                });
            }
        }
        */
    }
}