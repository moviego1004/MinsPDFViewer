using System;
using System.IO;
using System.Collections.Generic;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf.Annotations; // 주석 처리를 위해 필요

namespace MinsPDFViewer
{
    public class PdfService
    {
        public void SavePdf(PdfDocumentModel document, string savePath)
        {
            if (document == null) return;

            try
            {
                // 원본 파일을 읽어서 수정 모드로 엽니다.
                var originalBytes = File.ReadAllBytes(document.FilePath);
                using (var ms = new MemoryStream(originalBytes))
                using (var doc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                {
                    // PDF 버전 호환성 확보
                    if (doc.Version < 14) doc.Version = 14;

                    for (int i = 0; i < doc.PageCount && i < document.Pages.Count; i++)
                    {
                        var pdfPage = doc.Pages[i];
                        var pageVM = document.Pages[i];

                        // [좌표 보정] 스케일 계산: (PDF 실제 크기) / (뷰어/이미지 크기)
                        double scaleX = pdfPage.Width.Point / pageVM.Width;
                        double scaleY = pdfPage.Height.Point / pageVM.Height;

                        // ---------------------------------------------------------
                        // 1. 기존 주석(Highlight, Underline, FreeText) 저장 로직 보존
                        // ---------------------------------------------------------
                        if (pdfPage.Annotations != null)
                        {
                            // 기존 뷰어에서 관리하는 타입의 주석만 제거 (중복 방지)
                            var toRemove = new List<PdfAnnotation>();
                            for (int k = 0; k < pdfPage.Annotations.Count; k++)
                            {
                                string st = pdfPage.Annotations[k].Elements.GetString("/Subtype");
                                if (st == "/FreeText" || st == "/Highlight" || st == "/Underline")
                                    toRemove.Add(pdfPage.Annotations[k]);
                            }
                            foreach (var a in toRemove) pdfPage.Annotations.Remove(a);
                        }

                        // 화면에 있는 주석들을 PDF에 다시 쓰기
                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SearchHighlight || ann.Type == AnnotationType.Other) continue;

                            double ax = ann.X * scaleX;
                            double ay = ann.Y * scaleY;
                            double aw = ann.Width * scaleX;
                            double ah = ann.Height * scaleY;
                            double pdfY_BottomUp = pdfPage.Height.Point - (ay + ah); // PDF는 Y축이 아래에서 위로

                            if (ann.Type == AnnotationType.FreeText)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Elements.SetName("/Subtype", "/FreeText");
                                pdfAnnot.Rectangle = new PdfRectangle(new XRect(ax, pdfY_BottomUp, aw, ah));
                                pdfAnnot.Contents = ann.TextContent;
                                
                                var color = (ann.Foreground as System.Windows.Media.SolidColorBrush)?.Color ?? System.Windows.Media.Colors.Black;
                                double r = color.R / 255.0; double g = color.G / 255.0; double b = color.B / 255.0;
                                
                                // 폰트 및 색상 설정
                                pdfAnnot.Elements.SetString("/DA", $"/Helv {ann.FontSize} Tf {r:0.##} {g:0.##} {b:0.##} rg");
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else // Highlight, Underline
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                string subtype = (ann.Type == AnnotationType.Highlight) ? "/Highlight" : "/Underline";
                                
                                if (ann.Type == AnnotationType.Underline) 
                                    pdfY_BottomUp = pdfPage.Height.Point - (ay + 2); // 밑줄 위치 보정

                                var rect = new XRect(ax, pdfY_BottomUp, aw, ah);
                                pdfAnnot.Rectangle = new PdfRectangle(rect);
                                pdfAnnot.Elements.SetName("/Subtype", subtype);

                                // QuadPoints 설정 (하이라이트 영역 지정)
                                double qLeft = rect.X; double qRight = rect.X + rect.Width;
                                double qBottom = rect.Y; double qTop = rect.Y + rect.Height;
                                var quadPoints = new PdfArray(doc, new PdfReal(qLeft), new PdfReal(qTop), new PdfReal(qRight), new PdfReal(qTop), new PdfReal(qLeft), new PdfReal(qBottom), new PdfReal(qRight), new PdfReal(qBottom));
                                pdfAnnot.Elements.Add("/QuadPoints", quadPoints);

                                // 색상 설정
                                double r = ann.AnnotationColor.R / 255.0;
                                double g = ann.AnnotationColor.G / 255.0;
                                double b = ann.AnnotationColor.B / 255.0;
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(r), new PdfReal(g), new PdfReal(b));
                                
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }

                        // ---------------------------------------------------------
                        // 2. [수정됨] OCR 결과 투명 텍스트 저장 (Searchable PDF)
                        // ---------------------------------------------------------
                        if (pageVM.OcrWords != null && pageVM.OcrWords.Count > 0)
                        {
                            using (var gfx = XGraphics.FromPdfPage(pdfPage))
                            {
                                // [중요 1] 하얀색 겹침 방지: Alpha가 0인 진짜 투명 브러시 사용
                                var invisibleColor = XColor.FromArgb(0, 0, 0, 0);
                                var invisibleBrush = new XSolidBrush(invisibleColor);

                                // [중요 2] 한글 깨짐 방지: Unicode 옵션 적용
                                var fontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);

                                foreach (var word in pageVM.OcrWords)
                                {
                                    // 좌표 변환
                                    double x = word.BoundingBox.X * scaleX;
                                    double y = word.BoundingBox.Y * scaleY;
                                    double w = word.BoundingBox.Width * scaleX;
                                    double h = word.BoundingBox.Height * scaleY;

                                    // 폰트 크기 계산 (높이 기준 75%)
                                    double fSize = h * 0.75;
                                    if (fSize < 1) fSize = 1;

                                    // 폰트 생성 (Unicode 옵션 필수)
                                    var font = new XFont("Malgun Gothic", fSize, XFontStyle.Regular, fontOptions);

                                    // 투명 텍스트 그리기 (가운데 정렬)
                                    gfx.DrawString(word.Text, font, invisibleBrush,
                                        new XRect(x, y, w, h), XStringFormats.Center);
                                }
                            }
                        }
                    }
                    doc.Save(savePath);
                }
                // 성공 메시지는 MainWindow에서 띄우거나 여기서 처리
                System.Windows.MessageBox.Show($"저장 완료: {savePath}");
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"저장 실패: {ex.Message}");
            }
        }
    }
}