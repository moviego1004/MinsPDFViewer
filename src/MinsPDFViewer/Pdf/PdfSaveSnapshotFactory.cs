using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;

namespace MinsPDFViewer
{
    internal static class PdfSaveSnapshotFactory
    {
        public static async Task<(List<PageSaveData> Pages, List<BookmarkSaveData> Bookmarks)> CreateAsync(
            PdfDocumentModel model,
            Action<string>? log = null)
        {
            var pagesSnapshot = new List<PageSaveData>();
            var bookmarksSnapshot = new List<BookmarkSaveData>();

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var p in model.Pages)
                {
                    var pageData = new PageSaveData
                    {
                        OriginalPageIndex = p.OriginalPageIndex,
                        IsBlankPage = p.IsBlankPage,
                        Width = p.Width,
                        Height = p.Height,
                        PdfPageWidthPoint = p.PdfPageWidthPoint,
                        PdfPageHeightPoint = p.PdfPageHeightPoint,
                        Rotation = NormalizeRotation(p.Rotation),
                        OcrWords = p.OcrWords != null ? new List<OcrWordInfo>(p.OcrWords) : new List<OcrWordInfo>()
                    };

                    foreach (var ann in p.Annotations)
                    {
                        if (ann.Type == AnnotationType.SearchHighlight ||
                            ann.Type == AnnotationType.SignaturePlaceholder)
                            continue;

                        if (ann.Type == AnnotationType.FreeText && LooksLikeMojibake(ann.TextContent))
                        {
                            log?.Invoke($"Skipping mojibake FreeText annotation during save. Length={ann.TextContent?.Length ?? 0}, X={ann.X}, Y={ann.Y}");
                            continue;
                        }

                        pageData.Annotations.Add(new AnnotationSaveData
                        {
                            Type = ann.Type,
                            X = ann.X,
                            Y = ann.Y,
                            Width = ann.Width,
                            Height = ann.Height,
                            TextContent = ann.TextContent ?? "",
                            FontSize = ann.FontSize,
                            FontFamily = ann.FontFamily ?? "Malgun Gothic",
                            IsBold = ann.IsBold,
                            ForegroundColor = (ann.Foreground as SolidColorBrush)?.Color ?? Colors.Black,
                            BackgroundColor = (ann.Background as SolidColorBrush)?.Color ?? Colors.Transparent,
                            ImageBytes = ann.ImageBytes
                        });
                    }

                    pagesSnapshot.Add(pageData);
                }

                foreach (var bm in model.Bookmarks)
                {
                    var bookmark = MapBookmark(bm, model.Pages);
                    if (bookmark != null)
                        bookmarksSnapshot.Add(bookmark);
                }
            });

            return (pagesSnapshot, bookmarksSnapshot);
        }

        private static int NormalizeRotation(int rotation)
        {
            rotation %= 360;
            if (rotation < 0)
                rotation += 360;
            return rotation is 90 or 180 or 270 ? rotation : 0;
        }

        public static async Task<List<BookmarkSaveData>> CreateBookmarkSnapshotAsync(PdfDocumentModel model)
        {
            var bookmarksSnapshot = new List<BookmarkSaveData>();
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                foreach (var bm in model.Bookmarks)
                {
                    var bookmark = MapBookmark(bm, model.Pages);
                    if (bookmark != null)
                        bookmarksSnapshot.Add(bookmark);
                }
            });
            return bookmarksSnapshot;
        }

        private static BookmarkSaveData? MapBookmark(
            PdfBookmarkViewModel vm,
            System.Collections.ObjectModel.ObservableCollection<PdfPageViewModel> pages)
        {
            if (vm.PageIndex < 0 || vm.PageIndex >= pages.Count)
                return null;

            var data = new BookmarkSaveData
            {
                Title = vm.Title,
                OriginalPageIndex = pages[vm.PageIndex].OriginalPageIndex,
                PageIndex = vm.PageIndex
            };

            foreach (var child in vm.Children)
            {
                var childBookmark = MapBookmark(child, pages);
                if (childBookmark != null)
                    data.Children.Add(childBookmark);
            }

            return data;
        }

        private static bool LooksLikeMojibake(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            int suspicious = 0;
            foreach (char c in text)
            {
                if (c == 'í' || c == 'ê' || c == 'ë' || c == 'ì' || c == 'Ł' || c == 'œ' || c == 'š' || c == 'ﬂ' || c == '−')
                    suspicious++;
            }

            return suspicious >= 3;
        }
    }
}
