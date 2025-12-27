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

                    // 충돌 방지 Lock
                    lock (PdfService.PdfiumLock)
                    {
                        if (document.DocReader == null)
                            return null;
                        using (var pageReader = document.DocReader.GetPageReader(i))
                        {
                            pageText = pageReader.GetText() ?? "";
                            pageChars = pageReader.GetCharacters().ToList();
                        }
                    }

                    // [좌표계 자동 감지] 이 페이지가 뒤집혔는지 스스로 판단
                    bool needsFlip = DetectFlipNeed(pageChars);

                    int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                    int findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;

                        return Application.Current.Dispatcher.Invoke(() =>
                            RenderPageSearchResults(pageVM, pageText, pageChars!, query, findIndex, needsFlip));
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

                    // 충돌 방지 Lock
                    lock (PdfService.PdfiumLock)
                    {
                        if (document.DocReader == null)
                            return null;
                        using (var pageReader = document.DocReader.GetPageReader(i))
                        {
                            pageText = pageReader.GetText() ?? "";
                            pageChars = pageReader.GetCharacters().ToList();
                        }
                    }

                    // [좌표계 자동 감지]
                    bool needsFlip = DetectFlipNeed(pageChars);

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
                            RenderPageSearchResults(pageVM, pageText, pageChars!, query, findIndex, needsFlip));
                    }
                }
                return null;
            });
        }

        // [핵심 로직] 글자의 흐름을 보고 좌표계를 자동 판단
        private bool DetectFlipNeed(List<Character> chars)
        {
            // 글자가 너무 적으면 판단 불가 -> 기본값(표준)으로 가정
            if (chars == null || chars.Count < 5)
                return true;

            // 문서의 맨 앞 글자와 맨 뒤 글자의 Y좌표 비교
            var firstChar = chars[0];
            var lastChar = chars[chars.Count - 1];

            double firstY = (firstChar.Box.Top + firstChar.Box.Bottom) / 2.0;
            double lastY = (lastChar.Box.Top + lastChar.Box.Bottom) / 2.0;

            // 앞쪽 글자의 Y값이 더 크면? (아래로 갈수록 작아짐) -> 표준 좌표계(Bottom-Left) -> 뒤집기 필요(True)
            if (firstY > lastY)
                return true;

            // 앞쪽 글자의 Y값이 더 작으면? (아래로 갈수록 커짐) -> 이미지 좌표계(Top-Left) -> 뒤집기 불필요(False)
            return false;
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

        private PdfAnnotation? RenderPageSearchResults(PdfPageViewModel pageVM, string pageText, List<Character> chars, string query, int activeIndex, bool needsFlip)
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

                bool isFound = GetWordRect(pageVM, chars, index, query.Length, needsFlip, out Rect rect);

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

        private bool GetWordRect(PdfPageViewModel pageVM, List<Character> chars, int startIndex, int length, bool needsFlip, out Rect result)
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

            // CropBox 0 방지 처리가 PdfService에서 되었으므로 안전하게 계산
            double scaleX = (pageVM.CropWidthPoint > 0) ? pageVM.CropWidthPoint / pageVM.Width : 1.0;
            double scaleY = (pageVM.CropHeightPoint > 0) ? pageVM.CropHeightPoint / pageVM.Height : 1.0;

            double finalX, finalY, finalW, finalH;

            // X축: 좌측 기준 (공통)
            finalX = (minX - pageVM.CropX) / scaleX;
            finalW = (maxX - minX) / scaleX;
            finalH = (maxY - minY) / scaleY;

            // Y축: 자동 감지된 결과에 따라 분기
            if (needsFlip)
            {
                // 표준 PDF (Bottom-Left)인 경우 -> 뒤집어서 계산
                double pdfTopFromPageTop = pageVM.PdfPageHeightPoint - maxY;
                finalY = (pdfTopFromPageTop - pageVM.CropY) / scaleY;
            }
            else
            {
                // 이미지형 PDF (Top-Left)인 경우 -> 그대로 사용
                finalY = (minY - pageVM.CropY) / scaleY;
            }

            if (finalX < 0)
                finalX = 0;
            if (finalY < 0)
                finalY = 0;



            System.Diagnostics.Debug.WriteLine("========================================");
            System.Diagnostics.Debug.WriteLine($"[Page {pageVM.PageIndex}] Debug Info");
            System.Diagnostics.Debug.WriteLine($" - Rotation : {pageVM.Rotation}");
            System.Diagnostics.Debug.WriteLine($" - MediaBox : {pageVM.MediaBoxInfo}");
            System.Diagnostics.Debug.WriteLine($" - CropBox  : {pageVM.CropX}, {pageVM.CropY}, {pageVM.CropWidthPoint}, {pageVM.CropHeightPoint}");
            System.Diagnostics.Debug.WriteLine($" - UI Size  : {pageVM.Width} x {pageVM.Height}");
            System.Diagnostics.Debug.WriteLine($" - Raw Char Box (First) : {minX}, {minY}, {maxX}, {maxY}");
            System.Diagnostics.Debug.WriteLine($" - Final Result Rect    : {finalX}, {finalY}, {finalW}, {finalH}");
            System.Diagnostics.Debug.WriteLine("========================================");




            result = new Rect(finalX, finalY, finalW, finalH);
            return true;
        }
    }
}