using System;
using System.IO;
using System.Windows.Media.Imaging;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Drawing; 
using System.Windows.Media; 

namespace MinsPDFViewer
{
    public class PdfService
    {
        private readonly IDocLib _docLib;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        public PdfDocumentModel? LoadPdf(string filePath)
        {
            if (!File.Exists(filePath)) return null;

            try
            {
                // [1] 파일 잠금 방지: 메모리로 읽어서 Docnet에 전달
                var bytes = File.ReadAllBytes(filePath);
                var docReader = _docLib.GetDocReader(bytes, new PageDimensions(1.0));
                
                var model = new PdfDocumentModel
                {
                    FilePath = filePath,
                    FileName = Path.GetFileName(filePath),
                    DocLib = _docLib,
                    DocReader = docReader
                };

                return model;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                return null;
            }
        }

        public async Task RenderPagesAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null) return;

            double renderScale = 3.0;

            await Task.Run(() =>
            {
                // [2] 렌더링 시에도 파일 잠금 방지를 위해 bytes로 읽음
                byte[]? fileBytes = null;
                try 
                { 
                    fileBytes = File.ReadAllBytes(model.FilePath); 
                } 
                catch { return; } // 파일 읽기 실패 시 중단

                PdfDocument? sharpDoc = null;
                try
                {
                    // PdfSharp도 스트림으로 열기
                    using (var ms = new MemoryStream(fileBytes))
                    {
                        sharpDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Import);
                    }
                }
                catch { }

                // Docnet도 바이트 배열로 열기 (파일 경로 X)
                using (var renderReader = _docLib.GetDocReader(fileBytes, new PageDimensions(renderScale)))
                {
                    for (int i = 0; i < renderReader.GetPageCount(); i++)
                    {
                        using (var pageReader = renderReader.GetPageReader(i))
                        {
                            var rawWidth = pageReader.GetPageWidth();
                            var rawHeight = pageReader.GetPageHeight();
                            
                            var rawBytes = pageReader.GetImage(RenderFlags.RenderAnnotations);

                            BitmapSource? source = null;
                            
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                source = BitmapSource.Create(rawWidth, rawHeight, 96, 96, 
                                    System.Windows.Media.PixelFormats.Bgra32, null, 
                                    rawBytes, rawWidth * 4);
                                source.Freeze();
                            });

                            double uiWidth = rawWidth / renderScale;
                            double uiHeight = rawHeight / renderScale;

                            var pageVM = new PdfPageViewModel
                            {
                                PageIndex = i,
                                Width = uiWidth,   
                                Height = uiHeight, 
                                ImageSource = source 
                            };

                            if (sharpDoc != null && i < sharpDoc.PageCount)
                            {
                                var p = sharpDoc.Pages[i];
                                pageVM.PdfPageWidthPoint = p.Width.Point;
                                pageVM.PdfPageHeightPoint = p.Height.Point;
                                pageVM.CropX = p.CropBox.X1;
                                pageVM.CropY = p.CropBox.Y1; 
                                pageVM.CropWidthPoint = p.CropBox.Width;
                                pageVM.CropHeightPoint = p.CropBox.Height;

                                LoadSignatureFields(p, pageVM, (int)uiWidth, (int)uiHeight);
                            }

                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                model.Pages.Add(pageVM);
                            });
                        }
                    }
                }
                sharpDoc?.Dispose();
            });
        }

        private void LoadSignatureFields(PdfPage page, PdfPageViewModel pageVM, int uiWidth, int uiHeight)
        {
            if (page.Annotations == null) return;

            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k];
                
                if (annot.Elements.ContainsKey("/Subtype") && annot.Elements.GetString("/Subtype") == "/Widget" && 
                    annot.Elements.ContainsKey("/FT") && annot.Elements.GetString("/FT") == "/Sig")
                {
                    var rect = annot.Rectangle.ToXRect(); 
                    
                    double scaleX = (double)uiWidth / page.Width.Point;
                    double scaleY = (double)uiHeight / page.Height.Point;

                    double finalX = rect.X * scaleX;
                    double finalY = (page.Height.Point - (rect.Y + rect.Height)) * scaleY;
                    double finalW = rect.Width * scaleX;
                    double finalH = rect.Height * scaleY;

                    var sigDict = annot.Elements.GetDictionary("/V");
                    
                    var sigAnnot = new PdfAnnotation
                    {
                        Type = AnnotationType.SignatureField,
                        X = finalX,
                        Y = finalY,
                        Width = finalW,
                        Height = finalH,
                        SignatureData = sigDict 
                    };

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        pageVM.Annotations.Add(sigAnnot);
                    });
                }
            }
        }

        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || string.IsNullOrEmpty(model.FilePath)) return;

            string tempOutputPath = Path.GetTempFileName();

            try
            {
                // 원본 파일을 수정 모드로 엽니다. (메모리에 로드된 상태라 충돌 없음)
                using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Modify))
                {
                    foreach (var pageVM in model.Pages)
                    {
                        if (pageVM.PageIndex >= doc.PageCount) continue;
                        var pdfPage = doc.Pages[pageVM.PageIndex];

                        foreach (var ann in pageVM.Annotations)
                        {
                            if (ann.Type == AnnotationType.SignatureField || 
                                ann.Type == AnnotationType.SignaturePlaceholder) continue;

                            double effectivePdfWidth = (pageVM.CropWidthPoint > 0) ? pageVM.CropWidthPoint : pageVM.PdfPageWidthPoint;
                            double effectivePdfHeight = (pageVM.CropHeightPoint > 0) ? pageVM.CropHeightPoint : pageVM.PdfPageHeightPoint;
                            
                            double pdfOriginX = pageVM.CropX;
                            double pdfOriginY = pageVM.CropY;

                            double scaleX = effectivePdfWidth / pageVM.Width;
                            double scaleY = effectivePdfHeight / pageVM.Height;

                            double pdfX = pdfOriginX + (ann.X * scaleX);
                            double pdfY = (pdfOriginY + effectivePdfHeight) - ((ann.Y + ann.Height) * scaleY);
                            double pdfW = ann.Width * scaleX;
                            double pdfH = ann.Height * scaleY;

                            var rect = new PdfRectangle(new XRect(pdfX, pdfY, pdfW, pdfH));

                            // GenericPdfAnnotation 사용 (보호 수준 에러 해결)
                            if (ann.Type == AnnotationType.FreeText)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/FreeText");
                                pdfAnnot.Elements["/Contents"] = new PdfString(ann.TextContent);
                                
                                double r = ann.Foreground is SolidColorBrush b ? b.Color.R / 255.0 : 0;
                                double g = ann.Foreground is SolidColorBrush b2 ? b2.Color.G / 255.0 : 0;
                                double b_ = ann.Foreground is SolidColorBrush b3 ? b3.Color.B / 255.0 : 0;

                                string da = $"{r} {g} {b_} rg /Helv {ann.FontSize} Tf";
                                pdfAnnot.Elements["/DA"] = new PdfString(da);
                                
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else if (ann.Type == AnnotationType.Highlight)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/Highlight");
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(1), new PdfReal(1), new PdfReal(0));
                                pdfAnnot.Elements["/CA"] = new PdfReal(0.5);

                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                            else if (ann.Type == AnnotationType.Underline)
                            {
                                var pdfAnnot = new GenericPdfAnnotation(doc);
                                pdfAnnot.Rectangle = rect;
                                pdfAnnot.Elements["/Subtype"] = new PdfName("/Underline");
                                pdfAnnot.Elements["/C"] = new PdfArray(doc, new PdfReal(0), new PdfReal(0), new PdfReal(0));
                                
                                pdfPage.Annotations.Add(pdfAnnot);
                            }
                        }
                    }

                    // 임시 파일에 저장
                    doc.Save(tempOutputPath);
                }

                // 원본 덮어쓰기
                if (File.Exists(outputPath)) File.Delete(outputPath);
                File.Move(tempOutputPath, outputPath);
            }
            catch
            {
                if (File.Exists(tempOutputPath)) File.Delete(tempOutputPath);
                throw;
            }
        }
    }
}