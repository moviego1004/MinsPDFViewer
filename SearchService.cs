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

        public void ResetSearch()
        {
            _lastPageIndex = 0;
            _lastCharIndex = 0;
        }

        public async Task<PdfAnnotation?> FindNextAsync(PdfDocumentModel document, string query)
        {
            if (string.IsNullOrWhiteSpace(query) || document == null || document.DocReader == null)
                return null;

            if (query != _lastQuery)
            {
                _lastQuery = query;
                ResetSearch();
                foreach (var p in document.Pages) 
                {
                    var toRemove = p.Annotations.Where(a => a.Type == AnnotationType.SearchHighlight).ToList();
                    foreach (var r in toRemove) p.Annotations.Remove(r);
                }
            }

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
                        return CreateHighlight(pageVM, pageReader, findIndex, query);
                    }
                }
            }
            return null; 
        }

        // [수정] 반환 타입을 PdfAnnotation? (nullable)로 변경하여 CS8603 경고 해결
        private PdfAnnotation? CreateHighlight(PdfPageViewModel pageVM, Docnet.Core.Readers.IPageReader pageReader, int index, string query)
        {
            var chars = pageReader.GetCharacters().ToList();
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;
            bool found = false;

            for (int c = 0; c < query.Length; c++)
            {
                if (index + c < chars.Count)
                {
                    var box = chars[index + c].Box;
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

            // [사용자 요청] Scale 1.0 고정
            double scaleX = 1.0;
            double scaleY = 1.0;

            // Docnet 좌표는 Top-Left 기준이므로 Y반전 없이 CropBox 오프셋만 보정
            double finalX = (minX - pageVM.CropX) * scaleX;
            double finalY = (minY - pageVM.CropY) * scaleY;
            double finalW = (maxX - minX) * scaleX;
            double finalH = (maxY - minY) * scaleY;

            if (finalX < 0) finalX = 0;
            if (finalY < 0) finalY = 0;

            var ann = new PdfAnnotation
            {
                X = finalX,
                Y = finalY,
                Width = finalW,
                Height = finalH,
                Background = new SolidColorBrush(Color.FromArgb(120, 255, 255, 0)),
                Type = AnnotationType.SearchHighlight
            };
            
            pageVM.Annotations.Add(ann);
            return ann;
        }
    }
}