using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Media;
using System.Windows; 

namespace MinsPDFViewer
{
    public class SearchService
    {
        public List<PdfAnnotation> PerformSearch(PdfDocumentModel document, string query)
        {
            var results = new List<PdfAnnotation>();

            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null) 
                return results;

            foreach (var p in document.Pages)
            {
                var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                foreach (var r in toRemove) p.Annotations.Remove(r);
            }

            for (int i = 0; i < document.Pages.Count; i++)
            {
                var pageVM = document.Pages[i];
                double scaleX = pageVM.Width / pageVM.PdfPageWidthPoint;
                double scaleY = pageVM.Height / pageVM.PdfPageHeightPoint;

                // 1. OCR 검색
                if (pageVM.OcrWords != null)
                {
                    foreach (var word in pageVM.OcrWords)
                    {
                        if (word.Text.Contains(query, StringComparison.OrdinalIgnoreCase))
                        {
                            var ann = new PdfAnnotation
                            {
                                X = word.BoundingBox.X,
                                Y = word.BoundingBox.Y,
                                Width = word.BoundingBox.Width,
                                Height = word.BoundingBox.Height,
                                Background = new SolidColorBrush(Color.FromArgb(120, 255, 255, 0)),
                                Type = AnnotationType.SearchHighlight
                            };
                            pageVM.Annotations.Add(ann);
                            results.Add(ann);
                        }
                    }
                }

                // 2. Docnet 텍스트 검색 (보정 로직 적용)
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

                        for (int c = 0; c < query.Length; c++)
                        {
                            if (index + c < chars.Count)
                            {
                                var box = chars[index + c].Box;
                                
                                // 좌표 수집 (Bottom-Left 기준)
                                double cLeft = Math.Min(box.Left, box.Right);
                                double cRight = Math.Max(box.Left, box.Right);
                                double cTop = Math.Max(box.Top, box.Bottom);    // 큰 값이 Top (Y축 위쪽)
                                double cBottom = Math.Min(box.Top, box.Bottom); // 작은 값이 Bottom

                                minX = Math.Min(minX, cLeft);
                                minY = Math.Min(minY, cBottom);
                                maxX = Math.Max(maxX, cRight);
                                maxY = Math.Max(maxY, cTop);
                                found = true;
                            }
                        }

                        if (found)
                        {
                            // [핵심] 회전 감지 및 높이 기준 설정
                            // 만약 maxY가 현재 페이지 높이보다 크다면, 회전되어 Width가 Height 역할을 하는 것임.
                            // 또는 Rotation이 90/270도인 경우.
                            double flipBaseHeight = pageVM.PdfPageHeightPoint;
                            if (pageVM.Rotation == 90 || pageVM.Rotation == 270 || maxY > flipBaseHeight)
                            {
                                flipBaseHeight = pageVM.PdfPageWidthPoint;
                            }

                            // 1. X축: (절대좌표 - CropBox시작점) * 배율
                            double finalX = (minX - pageVM.CropX) * scaleX;
                            double finalW = (maxX - minX) * scaleX;

                            // 2. Y축: (기준높이 + CropBox시작점 - 절대좌표Top) * 배율
                            // maxY가 710이면 flipBaseHeight는 792(Width)가 되어 정상적인 양수 좌표가 나옴
                            double cropTopY = pageVM.CropY + flipBaseHeight;
                            double finalY = (cropTopY - maxY) * scaleY;
                            double finalH = (maxY - minY) * scaleY;

                            var ann = new PdfAnnotation
                            {
                                X = finalX,
                                Y = finalY,
                                Width = finalW,
                                Height = finalH,
                                Background = new SolidColorBrush(Color.FromArgb(120, 0, 255, 255)),
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