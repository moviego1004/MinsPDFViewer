using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Docnet.Core.Models;

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
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null)
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
                if (document.DocReader == null)
                    return null;

                for (int i = _lastPageIndex; i < document.Pages.Count; i++)
                {
                    if (_cancelSearch)
                        break;

                    var pageVM = document.Pages[i];
                    string pageText = "";
                    List<Character> pageChars = null;

                    lock (PdfService.PdfiumLock)
                    {
                        if (document.DocReader == null)
                            return null;

                        using (var pageReader = document.DocReader.GetPageReader(i))
                        {
                            // [수정] Null 경고 해결 (?? "")
                            pageText = pageReader.GetText() ?? "";
                            pageChars = pageReader.GetCharacters().ToList();
                        }
                    }

                    int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                    int findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;

                        return Application.Current.Dispatcher.Invoke(() =>
                            RenderPageSearchResults(pageVM, pageText, pageChars!, query, findIndex));
                    }

                    if (i == _lastPageIndex)
                        _lastCharIndex = 0;
                }
                return null;
            });
        }

        public async Task<PdfAnnotation?> FindPrevAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null)
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
                    var pageVM = document.Pages[i];
                    string pageText = "";
                    List<Character> pageChars = null;

                    lock (PdfService.PdfiumLock)
                    {
                        if (document.DocReader == null)
                            return null;

                        using (var pageReader = document.DocReader.GetPageReader(i))
                        {
                            // [수정] Null 경고 해결 (?? "")
                            pageText = pageReader.GetText() ?? "";
                            pageChars = pageReader.GetCharacters().ToList();
                        }
                    }

                    int startIndex = (i == _lastPageIndex) ? ((_lastCharIndex == -1 || _lastCharIndex == 0) ? pageText.Length - 1 : Math.Max(0, _lastCharIndex - 2)) : pageText.Length - 1;
                    if (startIndex < 0)
                        continue;
                    if (startIndex >= pageText.Length)
                        startIndex = pageText.Length - 1;

                    int findIndex = pageText.LastIndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex;
                        return Application.Current.Dispatcher.Invoke(() =>
                            RenderPageSearchResults(pageVM, pageText, pageChars!, query, findIndex));
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

        private PdfAnnotation? RenderPageSearchResults(PdfPageViewModel pageVM, string pageText, List<Character> chars, string query, int activeIndex)
        {
            var toRemove = pageVM.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
            foreach (var r in toRemove)
                pageVM.Annotations.Remove(r);

            PdfAnnotation? activeAnnotation = null;
            int currentIndex = 0;

            while (true)
            {
                int index = pageText.IndexOf(query, currentIndex, StringComparison.OrdinalIgnoreCase);
                if (index == -1)
                    break;

                bool isFound = GetWordRect(pageVM, chars, index, query.Length, out Rect rect);

                if (isFound)
                {
                    var ann = new PdfAnnotation { X = rect.X, Y = rect.Y, Width = rect.Width, Height = rect.Height, Type = AnnotationType.SearchHighlight };
                    if (index == activeIndex)
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
                currentIndex = index + 1;
            }
            return activeAnnotation;
        }

        private bool GetWordRect(PdfPageViewModel pageVM, List<Character> chars, int startIndex, int length, out Rect result)
        {
            result = new Rect();
            if (chars == null)
                return false;

            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            bool found = false;

            for (int c = 0; c < length; c++)
            {
                if (startIndex + c < chars.Count)
                {
                    var box = chars[startIndex + c].Box;
                    minX = Math.Min(minX, Math.Min(box.Left, box.Right));
                    minY = Math.Min(minY, Math.Min(box.Top, box.Bottom));
                    maxX = Math.Max(maxX, Math.Max(box.Left, box.Right));
                    maxY = Math.Max(maxY, Math.Max(box.Top, box.Bottom));
                    found = true;
                }
            }

            if (!found)
                return false;

            double scaleX = (pageVM.CropWidthPoint > 0) ? pageVM.CropWidthPoint / pageVM.Width : 1.0;
            double scaleY = (pageVM.CropHeightPoint > 0) ? pageVM.CropHeightPoint / pageVM.Height : 1.0;

            double topMargin = pageVM.PdfPageHeightPoint - (pageVM.CropY + pageVM.CropHeightPoint);
            double leftMargin = pageVM.CropX;

            double finalX, finalY;

            if (topMargin > 20)
            {
                finalY = (minY + (topMargin / 2)) / scaleY;
            }
            else
            {
                finalY = (minY - topMargin) / scaleY;
            }

            if (leftMargin > 20)
            {
                finalX = (minX - (leftMargin / 3)) / scaleX;
            }
            else
            {
                finalX = (minX - leftMargin) / scaleX;
            }

            double finalW = (maxX - minX) / scaleX;
            double finalH = (maxY - minY) / scaleY;

            if (finalX < 0)
                finalX = 0;
            if (finalY < 0)
                finalY = 0;

            result = new Rect(finalX, finalY, finalW, finalH);
            return true;
        }
    }
}