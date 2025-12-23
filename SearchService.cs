using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Windows; // Rect 사용을 위해

namespace MinsPDFViewer
{
    public class SearchService
    {
        // 검색 결과 리스트 반환
        public List<PdfAnnotation> PerformSearch(PdfDocumentModel document, string query)
        {
            var results = new List<PdfAnnotation>();

            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null) 
                return results;

            // 1. 기존 검색 하이라이트 제거
            foreach (var p in document.Pages)
            {
                var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                foreach (var r in toRemove) p.Annotations.Remove(r);
            }

            // 2. 페이지 순회하며 검색
            for (int i = 0; i < document.Pages.Count; i++)
            {
                var pageVM = document.Pages[i];

                // *** 핵심: 배율 계산 ***
                // Docnet이 반환하는 텍스트 좌표는 Point 단위(72dpi)입니다.
                // 렌더링된 이미지는 2.0배 확대(PageDimensions(2.0))되어 있으므로, 좌표도 비율만큼 늘려야 합니다.
                // 보통 scale = ViewWidth / PdfPointWidth 입니다.
                double scaleX = pageVM.Width / pageVM.PdfPageWidthPoint;
                double scaleY = pageVM.Height / pageVM.PdfPageHeightPoint;

                // [OCR 검색 부분 생략 가능 - 필요시 추가]

                // [Docnet 텍스트 검색]
                using (var pageReader = document.DocReader.GetPageReader(i))
                {
                    string pageText = pageReader.GetText();
                    var chars = pageReader.GetCharacters().ToList();

                    int index = 0;
                    while ((index = pageText.IndexOf(query, index, StringComparison.OrdinalIgnoreCase)) != -1)
                    {
                        double minX = double.MaxValue, minY = double.MaxValue;
                        double maxX = double.MinValue, maxY = double.MinValue;
                        bool found = false;

                        // 검색된 단어의 각 글자들의 바운딩 박스를 합칩니다.
                        for (int c = 0; c < query.Length; c++)
                        {
                            if (index + c < chars.Count)
                            {
                                var box = chars[index + c].Box;
                                
                                // *** 중요: Docnet 좌표는 Top-Left 기준이므로 그대로 사용합니다. ***
                                // Y축 반전(Flip)이나 CropBox 오프셋 계산을 하지 않습니다.
                                // '좌표잘찾는_MainWindow.xaml.cs'와 동일한 로직입니다.
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

                        if (found)
                        {
                            // *** 핵심: 좌표에 배율만 적용 ***
                            double finalX = minX * scaleX;
                            double finalY = minY * scaleY;
                            double finalW = (maxX - minX) * scaleX;
                            double finalH = (maxY - minY) * scaleY;

                            var ann = new PdfAnnotation
                            {
                                X = finalX,
                                Y = finalY,
                                Width = finalW,
                                Height = finalH,
                                Background = new SolidColorBrush(Color.FromArgb(120, 0, 255, 255)), // 진한 하늘색
                                Type = AnnotationType.SearchHighlight
                            };
                            pageVM.Annotations.Add(ann);
                            results.Add(ann);
                        }
                        index += query.Length;
                    }
                }
            }

            return results;
        }
    }
}