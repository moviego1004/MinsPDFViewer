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

            // 기존 하이라이트 제거
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

                // 2. Docnet 텍스트 검색
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
                                // PDF 좌표 수집 (Bottom-Left 기준)
                                double cLeft = Math.Min(box.Left, box.Right);
                                double cRight = Math.Max(box.Left, box.Right);
                                double cTop = Math.Max(box.Top, box.Bottom);
                                double cBottom = Math.Min(box.Top, box.Bottom);

                                minX = Math.Min(minX, cLeft);
                                minY = Math.Min(minY, cBottom);
                                maxX = Math.Max(maxX, cRight);
                                maxY = Math.Max(maxY, cTop);
                                found = true;
                            }
                        }

                        if (found)
                        {
                            double finalX, finalY, finalW, finalH;

                            // [회전 보정 로직]
                            // 90도 회전된 경우: Y축이 화면의 X축이 되고, X축이 화면의 Y축이 됨
                            // 또는 좌표값(maxY)이 페이지 높이(612)를 초과하는 경우 회전된 것으로 간주
                            bool isRotated90 = (pageVM.Rotation == 90) || (maxY > pageVM.PdfPageHeightPoint);

                            if (isRotated90)
                            {
                                // 90도 회전 공식:
                                // View X = (PageWidth - PDF_MaxY) * Scale
                                // View Y = (PDF_MinX) * Scale
                                
                                // Width는 회전된 상태의 너비(792)를 사용해야 함
                                double pageWidth = pageVM.PdfPageWidthPoint; 

                                finalX = (pageWidth - maxY) * scaleX; 
                                finalY = (minX) * scaleY;
                                
                                // 너비/높이도 스왑
                                finalW = (maxY - minY) * scaleX; 
                                finalH = (maxX - minX) * scaleY;
                            }
                            else
                            {
                                // 표준 (회전 없음):
                                // View X = (PDF_MinX) * Scale
                                // View Y = (PageHeight - PDF_MaxY) * Scale (Y축 반전)
                                
                                finalX = (minX - pageVM.CropX) * scaleX;
                                finalY = (pageVM.PdfPageHeightPoint + pageVM.CropY - maxY) * scaleY;
                                finalW = (maxX - minX) * scaleX;
                                finalH = (maxY - minY) * scaleY;
                            }

                            // 음수 보정 (화면 밖으로 나가는 경우 방지)
                            if (finalX < 0) finalX = 0;
                            if (finalY < 0) finalY = 0;

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