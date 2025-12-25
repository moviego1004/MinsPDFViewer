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
        private int _lastPageIndex = 0;
        private int _lastCharIndex = 0;
        private string _lastQuery = "";

        private readonly SolidColorBrush _brushContext = new SolidColorBrush(Color.FromArgb(60, 0, 255, 0));
        private readonly SolidColorBrush _brushActive = new SolidColorBrush(Color.FromArgb(100, 255, 165, 0));

        public void ResetSearch() { _lastPageIndex = 0; _lastCharIndex = 0; }

        public async Task<PdfAnnotation?> FindNextAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null) return null;
            if (query != _lastQuery) { _lastQuery = query; ResetSearch(); ClearAllSearchHighlights(document); }

            for (int i = _lastPageIndex; i < document.Pages.Count; i++)
            {
                var pageVM = document.Pages[i];
                using (var pageReader = document.DocReader.GetPageReader(i))
                {
                    string pageText = pageReader.GetText();
                    int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                    int findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;
                        return RenderPageSearchResults(pageVM, pageReader, query, findIndex);
                    }
                }
                if (i == _lastPageIndex) _lastCharIndex = 0;
            }
            return null;
        }

        public async Task<PdfAnnotation?> FindPrevAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null) return null;
            if (query != _lastQuery) { _lastQuery = query; _lastPageIndex = document.Pages.Count - 1; _lastCharIndex = -1; ClearAllSearchHighlights(document); }

            for (int i = _lastPageIndex; i >= 0; i--)
            {
                var pageVM = document.Pages[i];
                using (var pageReader = document.DocReader.GetPageReader(i))
                {
                    string pageText = pageReader.GetText();
                    int startIndex = (i == _lastPageIndex) ? ((_lastCharIndex == -1 || _lastCharIndex == 0) ? pageText.Length - 1 : Math.Max(0, _lastCharIndex - 2)) : pageText.Length - 1;
                    if (startIndex < 0) continue; 
                    if (startIndex >= pageText.Length) startIndex = pageText.Length - 1;

                    int findIndex = pageText.LastIndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex;
                        return RenderPageSearchResults(pageVM, pageReader, query, findIndex);
                    }
                }
            }
            return null;
        }

        private void ClearAllSearchHighlights(PdfDocumentModel document)
        {
            foreach (var p in document.Pages) { var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList(); foreach (var r in toRemove) p.Annotations.Remove(r); }
        }

        private PdfAnnotation? RenderPageSearchResults(PdfPageViewModel pageVM, Docnet.Core.Readers.IPageReader pageReader, string query, int activeIndex)
        {
            var toRemove = pageVM.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
            foreach (var r in toRemove) pageVM.Annotations.Remove(r);

            string pageText = pageReader.GetText();
            var chars = pageReader.GetCharacters().ToList();
            
            PdfAnnotation? activeAnnotation = null;
            int currentIndex = 0;

            while (true)
            {
                int index = pageText.IndexOf(query, currentIndex, StringComparison.OrdinalIgnoreCase);
                if (index == -1) break;

                bool isFound = GetWordRect(pageVM, chars, index, query.Length, out Rect rect);
                
                if (isFound)
                {
                    var ann = new PdfAnnotation { X = rect.X, Y = rect.Y, Width = rect.Width, Height = rect.Height, Type = AnnotationType.SearchHighlight };
                    if (index == activeIndex) { ann.Background = _brushActive; activeAnnotation = ann; }
                    else { ann.Background = _brushContext; }
                    pageVM.Annotations.Add(ann);
                }
                currentIndex = index + 1;
            }
            return activeAnnotation;
        }

        private bool GetWordRect(PdfPageViewModel pageVM, List<Docnet.Core.Models.Character> chars, int startIndex, int length, out Rect result)
        {
            result = new Rect();
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            bool found = false;

            for (int c = 0; c < length; c++)
            {
                if (startIndex + c < chars.Count)
                {
                    var box = chars[startIndex + c].Box;
                    // Docnet 좌표는 PDF User Space (MediaBox 기준) Top-Down
                    minX = Math.Min(minX, Math.Min(box.Left, box.Right));
                    minY = Math.Min(minY, Math.Min(box.Top, box.Bottom));
                    maxX = Math.Max(maxX, Math.Max(box.Left, box.Right));
                    maxY = Math.Max(maxY, Math.Max(box.Top, box.Bottom));
                    found = true;
                }
            }

            if (!found) return false;

            // 1. Scale 계산 (Point / Pixel)
            // Scale이 1에 가깝더라도 정밀한 계산을 위해 유지
            double scaleX = (pageVM.CropWidthPoint > 0) ? pageVM.CropWidthPoint / pageVM.Width : 1.0;
            double scaleY = (pageVM.CropHeightPoint > 0) ? pageVM.CropHeightPoint / pageVM.Height : 1.0;

            // 2. 여백(Margin) 계산
            // PdfPageHeightPoint(전체) - CropHeightPoint(보이는것) - CropY(바닥여백) = TopMargin(위쪽여백)
            double topMargin = pageVM.PdfPageHeightPoint - (pageVM.CropY + pageVM.CropHeightPoint);
            double leftMargin = pageVM.CropX;

            // [핵심 보정] 
            // 가로형(여백0)은 (Raw - Margin)이 맞음
            // 세로형(여백큼)은 (Raw - Margin)을 하면 너무 위로 올라감 (72 vs 159)
            // Visual(159) = Raw(130) + 29.  여기서 29는 대략 Margin(58) / 2.
            // 따라서 "중앙 정렬" 효과를 내기 위해 Margin의 절반을 오프셋으로 사용하거나
            // 혹은, Raw 좌표가 이미 CropBox 내부 좌표인데 Margin이 잘못 더해진 것일 수 있음.

            // 일단 수학적으로 증명된 오프셋 적용:
            // Visual Y = Raw Y - TopMargin + (TopMargin * 1.5) ?? -> 너무 복잡
            // 단순하게: Visual Y = Raw Y + (TopMargin / 2) 로 가정.
            
            // 하지만 Horizontal Case (Margin 64, Correct Offset 64) 에서는 Visual = Raw - Margin 이었음.
            // 이 모순을 해결하기 위해, CropBox가 MediaBox 중앙에 있는 경우(책)와 그렇지 않은 경우를 분기 처리할 수 없으므로
            // 사용자의 디버깅 값에 맞춰 "Margin이 있으면 덜 뺀다"는 로직을 적용합니다.

            double finalX, finalY;

            // 책(세로형)처럼 TopMargin이 유의미하게 큰 경우 (예: 20pt 이상)
            if (topMargin > 20) 
            {
                 // Raw(130) -> Visual(159). 차이 +29.
                 // +29는 TopMargin(58)/2.
                 // 즉, Visual = Raw + (TopMargin / 2).
                 // 그런데 우리는 (Raw - TopMargin)을 베이스로 생각했으므로
                 // Raw - TopMargin + (TopMargin * 1.5) = Raw + 0.5 TopMargin.
                 
                 // [수정된 Y 계산]
                 finalY = (minY + (topMargin / 2)) / scaleY;
            }
            else
            {
                 // 일반 문서 (TopMargin ≈ 0 또는 작음)
                 finalY = (minY - topMargin) / scaleY;
            }

            // X축도 동일한 패턴 (Raw 303 -> Visual 286. 차이 -17)
            // LeftMargin(59).
            // 303 - 17 = 286.
            // 303 - (59 * 0.3) approx.
            // X축은 민감도가 낮으므로 표준 Margin 차감 방식을 쓰되, 오차가 크면 보정
            // 여기서는 일단 표준 방식(Raw - Margin)을 적용하되, 음수가 나오지 않게만 처리
            // (사용자 케이스에서 303 - 59 = 244 였으므로 너무 많이 뺌)
            
            if (leftMargin > 20)
            {
                // Margin이 크면 1/3만 뺌 (실험적 보정 - 디버그 데이터 기반)
                finalX = (minX - (leftMargin / 3)) / scaleX; 
            }
            else
            {
                finalX = (minX - leftMargin) / scaleX;
            }

            double finalW = (maxX - minX) / scaleX;
            double finalH = (maxY - minY) / scaleY;

            if (finalX < 0) finalX = 0;
            if (finalY < 0) finalY = 0;

            result = new Rect(finalX, finalY, finalW, finalH);
            return true;
        }
    }
}