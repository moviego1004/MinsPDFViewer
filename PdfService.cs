using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MinsPDFViewer
{
    public class PdfService
    {
        public static readonly object PdfiumLock = new object();
        private readonly IDocLib _docLib;

        public PdfService()
        {
            _docLib = DocLib.Instance;
        }

        // [수정] LoadPdf를 비동기(Async)로 변경하여 UI 스레드 차단 방지
        public async Task<PdfDocumentModel?> LoadPdfAsync(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            return await Task.Run(() =>
            {
                try
                {
                    var bytes = File.ReadAllBytes(filePath);

                    IDocReader docReader;
                    // 여기서 락을 대기하더라도 백그라운드 스레드이므로 UI는 멈추지 않음
                    lock (PdfiumLock)
                    {
                        docReader = _docLib.GetDocReader(bytes, new PageDimensions(1.0));
                    }

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
            });
        }

        // [이름 변경] 전체 렌더링 -> 초기화 (이미지 생성 X)
        public async Task InitializeDocumentAsync(PdfDocumentModel model)
        {
            if (model.DocReader == null)
                return;

            // 203MB 파일 대응: 렌더링 스케일을 1.5로 줄임 (3.0은 메모리 폭발 원인)
            double renderScale = 1.5;

            await Task.Run(() =>
            {
                byte[]? originalBytes = null;
                try
                {
                    originalBytes = File.ReadAllBytes(model.FilePath);
                }
                catch { return; }

                // 1. Clean PDF 생성
                byte[] cleanPdfBytes = originalBytes;
                try
                {
                    using (var ms = new MemoryStream(originalBytes))
                    using (var sharpDoc = PdfReader.Open(ms, PdfDocumentOpenMode.Modify))
                    {
                        bool modified = false;
                        for (int i = 0; i < sharpDoc.PageCount; i++)
                        {
                            var page = sharpDoc.Pages[i];
                            if (page.Annotations != null)
                            {
                                for (int k = page.Annotations.Count - 1; k >= 0; k--)
                                {
                                    var annot = page.Annotations[k];
                                    if (annot == null)
                                        continue;
                                    var subtype = annot.Elements.GetString("/Subtype");
                                    if (subtype == "/FreeText" || subtype == "/Highlight" || subtype == "/Underline")
                                    {
                                        page.Annotations.Elements.RemoveAt(k);
                                        modified = true;
                                    }
                                }
                            }
                        }
                        if (modified)
                        {
                            using (var outMs = new MemoryStream())
                            {
                                sharpDoc.Save(outMs);
                                cleanPdfBytes = outMs.ToArray();
                            }
                        }
                    }
                }
                catch { }

                // 2. 렌더링용 CleanDocReader 생성 및 보관
                IDocReader cleanReader;
                lock (PdfiumLock)
                {
                    cleanReader = _docLib.GetDocReader(cleanPdfBytes, new PageDimensions(renderScale));
                }
                model.CleanDocReader = cleanReader;

                // 3. 페이지 껍데기(ViewModel) 생성 (이미지는 비워둠)
                // 원본 Reader를 사용하여 메타데이터 추출
                using (var msOriginal = new MemoryStream(originalBytes))
                using (var sharpDocOriginal = PdfReader.Open(msOriginal, PdfDocumentOpenMode.Import))
                {
                    for (int i = 0; i < cleanReader.GetPageCount(); i++)
                    {
                        double rawWidth = 0, rawHeight = 0;
                        lock (PdfiumLock)
                        {
                            using (var pr = cleanReader.GetPageReader(i))
                            {
                                rawWidth = pr.GetPageWidth();
                                rawHeight = pr.GetPageHeight();
                            }
                        }

                        double uiWidth = rawWidth / renderScale;
                        double uiHeight = rawHeight / renderScale;

                        var pageVM = new PdfPageViewModel
                        {
                            PageIndex = i,
                            Width = uiWidth,
                            Height = uiHeight,
                            ImageSource = null // [핵심] 여기서 이미지를 만들지 않음!
                        };

                        // 주석 및 추가 정보 추출
                        if (i < sharpDocOriginal.PageCount)
                        {
                            var p = sharpDocOriginal.Pages[i];
                            pageVM.PdfPageWidthPoint = p.Width.Point;
                            pageVM.PdfPageHeightPoint = p.Height.Point;
                            pageVM.CropX = p.CropBox.X1;
                            pageVM.CropY = p.CropBox.Y1;
                            pageVM.CropWidthPoint = p.CropBox.Width;
                            pageVM.CropHeightPoint = p.CropBox.Height;

                            // [필수 추가] CropBox가 비어있으면(0,0,0,0) 전체 크기로 설정 (안전장치)
                            if (pageVM.CropWidthPoint <= 0 || pageVM.CropHeightPoint <= 0)
                            {
                                pageVM.CropX = 0;
                                pageVM.CropY = 0;
                                pageVM.CropWidthPoint = pageVM.PdfPageWidthPoint;
                                pageVM.CropHeightPoint = pageVM.PdfPageHeightPoint;
                            }

                            bool hasSignature = CheckIfPageHasSignature(p);
                            pageVM.HasSignature = hasSignature;

                            var extractedAnns = ExtractAnnotationsFromPage(p);

                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                foreach (var ann in extractedAnns)
                                {
                                    double finalX = ann.X * (uiWidth / p.Width.Point);
                                    double finalY = ann.Y * (uiHeight / p.Height.Point);
                                    double finalW = ann.Width * (uiWidth / p.Width.Point);
                                    double finalH = ann.Height * (uiHeight / p.Height.Point);

                                    ann.X = finalX;
                                    ann.Y = finalY;
                                    ann.Width = finalW;
                                    if (ann.Type != AnnotationType.Underline)
                                        ann.Height = finalH;

                                    pageVM.Annotations.Add(ann);
                                }
                            });

                            // [추가] 회전 및 미디어 박스 정보 저장
                            pageVM.Rotation = p.Rotate;
                            pageVM.MediaBoxInfo = $"{p.MediaBox.X1},{p.MediaBox.Y1},{p.MediaBox.Width},{p.MediaBox.Height}";
                        }

                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            model.Pages.Add(pageVM);
                        });
                    }
                }
            });
        }

        // [신규] 한 페이지만 렌더링하는 메서드 (스크롤 될 때 호출)
        public void RenderPageImage(PdfDocumentModel model, PdfPageViewModel pageVM)
        {
            if (pageVM.ImageSource != null)
                return; // 이미 있으면 패스
            if (model.CleanDocReader == null)
                return;

            lock (PdfiumLock)
            {
                using (var pageReader = model.CleanDocReader.GetPageReader(pageVM.PageIndex))
                {
                    var rawWidth = pageReader.GetPageWidth();
                    var rawHeight = pageReader.GetPageHeight();
                    var rawBytes = pageReader.GetImage(RenderFlags.RenderAnnotations);

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        var source = BitmapSource.Create(rawWidth, rawHeight, 96, 96,
                            PixelFormats.Bgra32, null, rawBytes, rawWidth * 4);
                        source.Freeze();
                        pageVM.ImageSource = source;
                    });
                }
            }
        }

        // ... (CheckIfPageHasSignature, ExtractAnnotationsFromPage, SavePdf 등 기존 메서드 그대로 유지)
        // (지면 관계상 생략하지만 기존 코드를 꼭 유지해주세요!)
        // (ExtractAnnotationsFromPage, ParseAnnotationColor, ParseAnnotationFont, SavePdf, CheckIfPageHasSignature)

        private bool CheckIfPageHasSignature(PdfPage page)
        {
            if (page.Annotations == null)
                return false;
            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k];
                if (annot == null)
                    continue;
                if (annot.Elements.GetString("/Subtype") == "/Widget" && annot.Elements.GetString("/FT") == "/Sig")
                    return true;
            }
            return false;
        }

        private List<PdfAnnotation> ExtractAnnotationsFromPage(PdfPage page)
        {
            var list = new List<PdfAnnotation>();
            if (page.Annotations == null)
                return list;
            for (int k = 0; k < page.Annotations.Count; k++)
            {
                var annot = page.Annotations[k];
                if (annot == null)
                    continue;
                var subtype = annot.Elements.GetString("/Subtype") ?? "";
                var rect = annot.Rectangle.ToXRect();
                double finalX = rect.X;
                double finalY = page.Height.Point - (rect.Y + rect.Height);
                double finalW = rect.Width;
                double finalH = rect.Height;
                PdfAnnotation? newAnnot = null;

                if (subtype == "/Widget" && annot.Elements.GetString("/FT") == "/Sig")
                {
                    string fieldName = annot.Elements.GetString("/T");
                    newAnnot = new PdfAnnotation { Type = AnnotationType.SignatureField, FieldName = fieldName };
                }
                else if (subtype == "/FreeText")
                {
                    string da = annot.Elements.GetString("/DA");
                    Brush textColor = Brushes.Black; // Simplified
                    try
                    { /*ParseColor logic*/
                    }
                    catch { }
                    newAnnot = new PdfAnnotation { Type = AnnotationType.FreeText, TextContent = annot.Contents, Foreground = textColor, Background = Brushes.Transparent };
                }
                // ... (Highlight, Underline 생략 - 기존 코드 유지)
                if (newAnnot != null)
                {
                    newAnnot.X = finalX;
                    newAnnot.Y = finalY;
                    newAnnot.Width = finalW;
                    newAnnot.Height = finalH;
                    list.Add(newAnnot);
                }
            }
            return list;
        }

        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            // 기존 SavePdf 코드 그대로 유지
            if (model == null || string.IsNullOrEmpty(model.FilePath))
                return;
            // ...
            try
            {
                File.Copy(model.FilePath, outputPath, true); // 임시 (기존 로직 사용하세요)
            }
            catch { }
        }
    }

     // [신규] PDF 파일에서 책갈피 읽어오기 (재귀 호출)
    public void LoadBookmarks(PdfDocumentModel model)
        {
            try
            {
                if (!File.Exists(model.FilePath))
                    return;

                // PdfSharp으로 문서 구조 읽기 (Import 모드)
                using (var doc = PdfReader.Open(model.FilePath, PdfDocumentOpenMode.Import))
                {
                    model.Bookmarks.Clear();
                    foreach (var outline in doc.Outlines)
                    {
                        var vm = ConvertOutlineToViewModel(outline, doc, null);
                        if (vm != null)
                            model.Bookmarks.Add(vm);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"책갈피 로드 실패: {ex.Message}");
            }
        }

        // [신규] 내부 헬퍼: PdfOutline -> ViewModel 변환
        private PdfBookmarkViewModel? ConvertOutlineToViewModel(PdfOutline outline, PdfDocument doc, PdfBookmarkViewModel? parent)
        {
            int pageIndex = -1;

            // 목적지 페이지 찾기
            if (outline.DestinationPage != null)
            {
                // PdfSharp의 Pages 컬렉션에서 인덱스 찾기
                for (int i = 0; i < doc.Pages.Count; i++)
                {
                    if (doc.Pages[i].Equals(outline.DestinationPage))
                    {
                        pageIndex = i;
                        break;
                    }
                }
            }

            // 페이지를 못 찾았으면 책갈피 무시 (혹은 0페이지로)
            if (pageIndex == -1)
                pageIndex = 0;

            var vm = new PdfBookmarkViewModel
            {
                Title = outline.Title,
                PageIndex = pageIndex,
                Parent = parent
            };

            // 자식 책갈피 재귀 호출
            foreach (var child in outline.Outlines)
            {
                var childVm = ConvertOutlineToViewModel(child, doc, vm);
                if (childVm != null)
                    vm.Children.Add(childVm);
            }

            return vm;
        }

        // [수정] SavePdf 메서드에서 책갈피 저장 로직 추가
        public void SavePdf(PdfDocumentModel model, string outputPath)
        {
            if (model == null || string.IsNullOrEmpty(model.FilePath))
                return;

            // 1. 임시 파일로 복사 (기존 로직 유지)
            string tempPath = Path.GetTempFileName();
            File.Copy(model.FilePath, tempPath, true);

            try
            {
                // 2. PdfSharp으로 열어서 수정 (Modify 모드)
                using (var doc = PdfReader.Open(tempPath, PdfDocumentOpenMode.Modify))
                {
                    // 2-1. 기존 주석/서명 저장 (기존 로직 유지 - 안 건드림)
                    // (이 부분은 이전 대화의 SavePdf 내용을 그대로 두시면 됩니다.)
                    // ... (생략된 기존 주석 저장 코드) ...

                    // 2-2. [추가] 책갈피(Outlines) 덮어쓰기
                    doc.Outlines.Clear(); // 기존 책갈피 삭제
                    foreach (var bm in model.Bookmarks)
                    {
                        SaveBookmarkToDocument(bm, doc.Outlines, doc);
                    }

                    doc.Save(outputPath);
                }
            }
            catch
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
                throw;
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        // [신규] ViewModel -> PdfOutline 변환 및 저장 (재귀)
        private void SaveBookmarkToDocument(PdfBookmarkViewModel vm, PdfOutlineCollection parentCollection, PdfDocument doc)
        {
            if (vm.PageIndex >= 0 && vm.PageIndex < doc.PageCount)
            {
                var page = doc.Pages[vm.PageIndex];
                var outline = parentCollection.Add(vm.Title, page, true); // true = 펼침 상태

                foreach (var child in vm.Children)
                {
                    SaveBookmarkToDocument(child, outline.Outlines, doc);
                }
            }
        }
    }