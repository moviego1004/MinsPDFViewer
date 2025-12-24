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
        // 검색 진행 상태 저장 (현재 페이지, 현재 글자 위치)
        private int _lastPageIndex = 0;
        private int _lastCharIndex = 0;
        private string _lastQuery = "";

        // 검색 상태 초기화 (처음부터 다시 찾기)
        public void ResetSearch()
        {
            _lastPageIndex = 0;
            _lastCharIndex = 0;
        }

        // [핵심] 다음 검색 결과 하나만 찾기 (페이지별 순차 검색)
        public async Task<PdfAnnotation?> FindNextAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null)
                return null;

            // 검색어가 변경되면 처음부터 다시 시작
            if (query != _lastQuery)
            {
                _lastQuery = query;
                ResetSearch();
                // 기존 하이라이트 제거
                foreach (var p in document.Pages) 
                {
                    var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                    foreach (var r in toRemove) p.Annotations.Remove(r);
                }
            }

            // 현재 위치(_lastPageIndex)부터 문서 끝까지 순차 검색
            for (int i = _lastPageIndex; i < document.Pages.Count; i++)
            {
                var pageVM = document.Pages[i];
                
                using (var pageReader = document.DocReader.GetPageReader(i))
                {
                    string pageText = pageReader.GetText();
                    
                    // 현재 페이지에서 검색 (첫 진입 페이지는 저장된 위치부터, 이후 페이지는 0부터)
                    int startIndex = (i == _lastPageIndex) ? _lastCharIndex : 0;
                    int findIndex = pageText.IndexOf(query, startIndex, StringComparison.OrdinalIgnoreCase);

                    if (findIndex != -1)
                    {
                        // 찾음: 다음 검색을 위해 위치 저장 (현재 위치 + 1)
                        _lastPageIndex = i;
                        _lastCharIndex = findIndex + 1;

                        // 하이라이트 생성 및 반환
                        return CreateHighlight(pageVM, pageReader, findIndex, query);
                    }
                }
                
                // 페이지를 넘어가면 글자 인덱스 초기화
                // (다음 루프의 페이지는 0부터 검색해야 하므로)
                if (i == _lastPageIndex) 
                {
                    // 현재 페이지 루프가 끝나면 다음 페이지로 넘어가기 전 인덱스 리셋이 필요없음
                    // (startIndex 로직에서 처리됨)
                }
            }

            // 문서 끝까지 찾았으나 결과가 없음 -> null 반환 (UI에서 메시지박스 처리)
            return null; 
        }

        private PdfAnnotation CreateHighlight(PdfPageViewModel pageVM, Docnet.Core.Readers.IPageReader pageReader, int index, string query)
        {
            var chars = pageReader.GetCharacters().ToList();
            
            // 검색된 단어 영역 계산
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            bool found = false;

            for (int c = 0; c < query.Length; c++)
            {
                if (index + c < chars.Count)
                {
                    var box = chars[index + c].Box;
                    
                    // Docnet은 Top-Left 좌표계를 사용하므로 안전하게 Min/Max만 추출
                    double cLeft = Math.Min(box.Left, box.Right);
                    double cRight = Math.Max(box.Left, box.Right);
                    double cTop = Math.Min(box.Top, box.Bottom);
                    double cBottom = Math.Max(box.Top, box.Bottom);

                    minX = Math.Min(minX, cLeft);
                    minY = Math.Min(minY, cTop);
                    maxX = Math.Max(maxX, cRight);
                    maxY = Math.Max(maxY, cBottom);
                    
                    found = true;
                }
            }

            if (!found) return null;

            // [사용자 검증 반영] 배율을 1.0으로 고정
            // WPF Canvas가 이미 뷰박스 등을 통해 배율을 처리하거나, 
            // Docnet 좌표가 화면 좌표와 1:1로 매칭되는 상황입니다.
            double scaleX = 1.0;
            double scaleY = 1.0;

            // 좌표 계산 (CropBox 오프셋 보정, Y축 반전 없음)
            double finalX = (minX - pageVM.CropX) * scaleX;
            double finalY = (minY - pageVM.CropY) * scaleY;
            double finalW = (maxX - minX) * scaleX;
            double finalH = (maxY - minY) * scaleY;

            // 음수 좌표 방지
            if (finalX < 0) finalX = 0;
            if (finalY < 0) finalY = 0;

            var ann = new PdfAnnotation
            {
                X = finalX,
                Y = finalY,
                Width = finalW,
                Height = finalH,
                Background = new SolidColorBrush(Color.FromArgb(120, 255, 255, 0)), // 형광색
                Type = AnnotationType.SearchHighlight
            };
            
            // 뷰모델에 즉시 추가하여 화면에 표시
            pageVM.Annotations.Add(ann);

            return ann;
        }
    }
}