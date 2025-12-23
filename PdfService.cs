using Docnet.Core;
using Docnet.Core.Models;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MinsPDFViewer
{
    public class PdfService
    {
        private readonly IDocLib _docLib;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        public PdfDocumentModel? LoadPdf(string path)
        {
            try
            {
                var newDoc = new PdfDocumentModel { FilePath = path, FileName = Path.GetFileName(path) };
                var fileBytes = File.ReadAllBytes(path);

                // 1. PdfSharp으로 주석 추출 (좌표 변환 필요 시 여기서 처리)
                var extractedRawData = new Dictionary<int, List<RawAnnotationInfo>>();
                var pdfPageSizes = new Dictionary<int, PdfSharp.Drawing.XSize>();

                using (var msInput = new MemoryStream(fileBytes))
                using (var doc = PdfReader.Open(msInput, PdfDocumentOpenMode.Modify))
                {
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        var page = doc.Pages[i];
                        pdfPageSizes[i] = new XSize(page.Width.Point, page.Height.Point);
                        extractedRawData[i] = new List<RawAnnotationInfo>();

                        if (page.Annotations != null)
                        {
                            // 주석 추출 로직 (단순화: 여기서는 추출만 하고 뷰모델 변환은 아래에서)
                            // 기존 코드의 주석 추출 로직 유지...
                        }
                    }
                    
                    // Docnet용 클린 스트림 생성
                    var cleanStream = new MemoryStream(); 
                    doc.Save(cleanStream);
                    
                    // *** 중요: 2.0 배율로 로드 (고해상도) ***
                    newDoc.DocReader = _docLib.GetDocReader(cleanStream.ToArray(), new PageDimensions(2.0));
                }

                if (newDoc.DocReader != null)
                {
                    int pc = newDoc.DocReader.GetPageCount();
                    for (int i = 0; i < pc; i++)
                    {
                        using (var r = newDoc.DocReader.GetPageReader(i))
                        {
                            // Docnet이 렌더링한 이미지 크기 (2배 확대된 크기)
                            double viewW = r.GetPageWidth(); 
                            double viewH = r.GetPageHeight();
                            
                            var pvm = new PdfPageViewModel 
                            { 
                                PageIndex = i, 
                                Width = viewW, 
                                Height = viewH,
                                // 원본 PDF 크기 (Point 단위)
                                PdfPageWidthPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Width : viewW / 2.0,
                                PdfPageHeightPoint = pdfPageSizes.ContainsKey(i) ? pdfPageSizes[i].Height : viewH / 2.0
                            };

                            // 주석 뷰모델 변환 및 추가 로직은 필요하다면 여기에 복원
                            newDoc.Pages.Add(pvm);
                        }
                    }
                }

                return newDoc;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 로드 실패: {ex.Message}");
                return null;
            }
        }

        public async Task RenderPagesAsync(PdfDocumentModel doc)
        {
            await Task.Run(() =>
            {
                Parallel.For(0, doc.Pages.Count, new ParallelOptions { MaxDegreeOfParallelism = 4 }, i =>
                {
                    if (doc.DocReader == null) return;
                    using (var r = doc.DocReader.GetPageReader(i))
                    {
                        var bytes = r.GetImage();
                        var w = r.GetPageWidth();
                        var h = r.GetPageHeight();
                        
                        Application.Current.Dispatcher.Invoke(() => 
                        { 
                            if (i < doc.Pages.Count) 
                                doc.Pages[i].ImageSource = RawBytesToBitmapImage(bytes, w, h); 
                        });
                    }
                });
            });
        }

        private BitmapImage RawBytesToBitmapImage(byte[] b, int w, int h)
        {
            var bm = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
            bm.WritePixels(new Int32Rect(0, 0, w, h), b, w * 4, 0);
            if (bm.CanFreeze) bm.Freeze();
            
            using (var ms = new MemoryStream())
            {
                var enc = new PngBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bm));
                enc.Save(ms);
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.StreamSource = ms;
                img.EndInit();
                if (img.CanFreeze) img.Freeze();
                return img;
            }
        }
    }
}