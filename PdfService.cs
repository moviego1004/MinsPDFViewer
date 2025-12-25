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
using PdfSharp.Pdf.Annotations; // [추가]

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
                var docReader = _docLib.GetDocReader(filePath, new PageDimensions(1.0));
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

            await Task.Run(() =>
            {
                PdfDocument? sharpDoc = null;
                try
                {
                    // [수정] ReadOnly -> Import 사용 (경고 해결)
                    sharpDoc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Import);
                }
                catch { }

                for (int i = 0; i < model.DocReader.GetPageCount(); i++)
                {
                    using (var pageReader = model.DocReader.GetPageReader(i))
                    {
                        var width = pageReader.GetPageWidth();
                        var height = pageReader.GetPageHeight();
                        var rawBytes = pageReader.GetImage();

                        BitmapSource? source = null;
                        
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            source = BitmapSource.Create(width, height, 96, 96, 
                                System.Windows.Media.PixelFormats.Bgra32, null, 
                                rawBytes, width * 4);
                            source.Freeze();
                        });

                        var pageVM = new PdfPageViewModel
                        {
                            PageIndex = i,
                            Width = width,
                            Height = height,
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

                            LoadSignatureFields(p, pageVM, width, height);
                        }

                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            model.Pages.Add(pageVM);
                        });
                    }
                }
                sharpDoc?.Dispose();
            });
        }

        private void LoadSignatureFields(PdfPage page, PdfPageViewModel pageVM, int renderPixelWidth, int renderPixelHeight)
        {
            if (page.Annotations == null) return;

            // [수정] PdfAnnotation으로 안전하게 접근
            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k]; // 여기서 annot은 PdfAnnotation 타입
                
                // Elements 속성 접근 가능해짐
                if (annot.Elements.ContainsKey("/Subtype") && annot.Elements.GetString("/Subtype") == "/Widget" && 
                    annot.Elements.ContainsKey("/FT") && annot.Elements.GetString("/FT") == "/Sig")
                {
                    var rect = annot.Rectangle; 
                    
                    double scaleX = renderPixelWidth / page.Width.Point;
                    double scaleY = renderPixelHeight / page.Height.Point;

                    double finalX = rect.X * scaleX;
                    double finalY = (page.Height.Point - (rect.Y + rect.Height)) * scaleY;
                    double finalW = rect.Width * scaleX;
                    double finalH = rect.Height * scaleY;

                    // 검증용 딕셔너리 추출
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
            File.Copy(model.FilePath, outputPath, true);
        }
    }
}