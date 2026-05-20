using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using DrawingBitmap = System.Drawing.Bitmap;
using DrawingGraphics = System.Drawing.Graphics;
using DrawingImageLockMode = System.Drawing.Imaging.ImageLockMode;
using DrawingImageFormat = System.Drawing.Imaging.ImageFormat;
using DrawingPixelFormat = System.Drawing.Imaging.PixelFormat;
using DrawingRectangle = System.Drawing.Rectangle;

namespace MinsPDFViewer
{
    public sealed class PdfiumEditService
    {
        private const int FPDF_ANNOT_FREETEXT = 3;
        private const int FPDF_ANNOT_HIGHLIGHT = 9;
        private const int FPDF_ANNOT_UNDERLINE = 10;
        private const int FPDF_ANNOT_STAMP = 13;
        private const int FPDF_ANNOT_WIDGET = 12;
        private const int FPDF_PAGEOBJ_IMAGE = 3;
        private const int FPDFANNOT_COLORTYPE_Color = 0;
        internal const string FreeTextMetadataPrefix = "MINS_FREETEXT_V1:";
        internal const string FreeTextMetadataV2Prefix = "MINS_FREETEXT_V2:";
        internal const string ImageStampMetadataPrefix = "MINS_IMAGESTAMP_V1:";
        internal const string ImageStampMetadataV2Prefix = "MINS_IMAGESTAMP_V2:";

        public async Task<bool> TrySaveAnnotationsAsync(PdfDocumentModel model, string outputPath)
        {
            if (!CanUsePdfiumAnnotationSave(model, out _))
                return false;

            var snapshot = await CreateSnapshotAsync(model);
            if (snapshot == null)
                return false;

            return await Task.Run(() => SaveSnapshot(snapshot, outputPath));
        }

        internal static string GetUnsupportedReason(PdfDocumentModel model)
        {
            CanUsePdfiumAnnotationSave(model, out string reason);
            return reason;
        }

        private static bool CanUsePdfiumAnnotationSave(PdfDocumentModel model, out string reason)
        {
            if (model.PdfDocument == null || string.IsNullOrWhiteSpace(model.FilePath) || !File.Exists(model.FilePath))
            {
                reason = "document is not loaded or source file is missing";
                return false;
            }

            int sourcePageCount;
            lock (PdfService.PdfiumLock)
            {
                sourcePageCount = model.PdfDocument?.PageCount ?? model.Pages.Count;
            }

            if (model.Pages.Count(p => !p.IsBlankPage) > sourcePageCount)
            {
                reason = "document contains inserted pages";
                return false;
            }

            int previousOriginalPageIndex = -1;
            var seenOriginalPages = new HashSet<int>();
            foreach (var page in model.Pages)
            {
                if (page.IsBlankPage)
                    continue;

                if (page.OriginalPageIndex < 0 || page.OriginalPageIndex >= sourcePageCount)
                {
                    reason = "page references an invalid source page";
                    return false;
                }

                if (!seenOriginalPages.Add(page.OriginalPageIndex))
                {
                    reason = "page set contains duplicated source pages";
                    return false;
                }

                if (page.OriginalPageIndex <= previousOriginalPageIndex)
                {
                    reason = "page order has changed";
                    return false;
                }

                previousOriginalPageIndex = page.OriginalPageIndex;
            }

            reason = "";
            return true;
        }

        private static async Task<DocumentSnapshot?> CreateSnapshotAsync(PdfDocumentModel model)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null || dispatcher.CheckAccess())
                return CreateSnapshot(model);

            var operation = dispatcher.InvokeAsync(() => CreateSnapshot(model));
            if (await Task.WhenAny(operation.Task, Task.Delay(TimeSpan.FromSeconds(5))) == operation.Task)
                return await operation.Task;

            return CreateSnapshot(model);
        }

        private static DocumentSnapshot CreateSnapshot(PdfDocumentModel model)
        {
            int sourcePageCount = 0;
            lock (PdfService.PdfiumLock)
            {
                sourcePageCount = model.PdfDocument?.PageCount ?? model.Pages.Count;
            }

            var pages = new List<PageSnapshot>();
            foreach (var page in model.Pages)
            {
                var pageSnapshot = new PageSnapshot
                {
                    OriginalPageIndex = page.OriginalPageIndex,
                    IsBlankPage = page.IsBlankPage,
                    Width = page.Width,
                    Height = page.Height,
                    PdfWidth = page.PdfPageWidthPoint,
                    PdfHeight = page.PdfPageHeightPoint,
                    Rotation = NormalizeRotation(page.Rotation),
                    ShouldRewriteAnnotations = page.IsBlankPage ||
                                               page.AnnotationsLoaded ||
                                               page.Annotations.Any(IsPersistentAnnotation) ||
                                               (page.OcrWords != null && page.OcrWords.Count > 0) ||
                                               NormalizeRotation(page.Rotation) != 0
                };

                foreach (var annotation in page.Annotations)
                {
                    if (!IsPersistentAnnotation(annotation))
                        continue;

                    if (annotation.Type == AnnotationType.FreeText && LooksLikeMojibake(annotation.TextContent))
                        continue;

                    pageSnapshot.Annotations.Add(new AnnotationSnapshot
                    {
                        Type = annotation.Type,
                        X = annotation.X,
                        Y = annotation.Y,
                        Width = annotation.Width,
                        Height = annotation.Height,
                        Text = annotation.TextContent ?? string.Empty,
                        FontSize = annotation.FontSize,
                        FontFamily = annotation.FontFamily ?? "Malgun Gothic",
                        IsBold = annotation.IsBold,
                        Foreground = (annotation.Foreground as SolidColorBrush)?.Color ?? Colors.Black,
                        Background = (annotation.Background as SolidColorBrush)?.Color ?? annotation.AnnotationColor,
                        ImageBytes = annotation.ImageBytes
                    });
                }

                if (page.OcrWords != null)
                {
                    foreach (var word in page.OcrWords)
                    {
                        if (string.IsNullOrWhiteSpace(word.Text) ||
                            word.BoundingBox.Width <= 0 ||
                            word.BoundingBox.Height <= 0)
                            continue;

                        pageSnapshot.OcrWords.Add(new OcrWordSnapshot
                        {
                            Text = word.Text,
                            X = word.BoundingBox.X,
                            Y = word.BoundingBox.Y,
                            Width = word.BoundingBox.Width,
                            Height = word.BoundingBox.Height
                        });
                    }
                }

                pages.Add(pageSnapshot);
            }

            return new DocumentSnapshot(model.FilePath, sourcePageCount, pages);
        }

        private static bool IsPersistentAnnotation(PdfAnnotation annotation)
        {
            return annotation.Type == AnnotationType.FreeText ||
                   annotation.Type == AnnotationType.Highlight ||
                   annotation.Type == AnnotationType.Underline ||
                   annotation.Type == AnnotationType.ImageStamp;
        }

        private static bool SaveSnapshot(DocumentSnapshot snapshot, string outputPath)
        {
            lock (PdfService.PdfiumLock)
            {
                IntPtr document = NativeMethods.FPDF_LoadDocument(snapshot.SourcePath, null);
                if (document == IntPtr.Zero)
                    return false;

                IntPtr ocrFont = IntPtr.Zero;
                try
                {
                    foreach (var pageSnapshot in snapshot.Pages)
                    {
                        if (pageSnapshot.IsBlankPage)
                            continue;

                        if (pageSnapshot.Rotation == 0 &&
                            (!pageSnapshot.ShouldRewriteAnnotations ||
                             (pageSnapshot.Annotations.Count == 0 && pageSnapshot.OcrWords.Count == 0)))
                            continue;

                        IntPtr page = NativeMethods.FPDF_LoadPage(document, pageSnapshot.OriginalPageIndex);
                        if (page == IntPtr.Zero)
                            return false;

                        try
                        {
                            if (pageSnapshot.Rotation != 0)
                                NativeMethods.FPDFPage_SetRotation(page, pageSnapshot.Rotation / 90);

                            if (pageSnapshot.ShouldRewriteAnnotations &&
                                (pageSnapshot.Annotations.Count > 0 || pageSnapshot.OcrWords.Count > 0))
                            {
                                RemoveEditableAnnotations(page);

                                foreach (var annotation in pageSnapshot.Annotations)
                                {
                                    AddPersistentObject(document, page, pageSnapshot, annotation);
                                }

                                AddOcrTextLayer(document, page, pageSnapshot, ref ocrFont);
                                NativeMethods.FPDFPage_GenerateContent(page);
                            }
                        }
                        finally
                        {
                            NativeMethods.FPDF_ClosePage(page);
                        }
                    }

                    DeleteRemovedPages(document, snapshot);
                    InsertBlankPages(document, snapshot, ref ocrFont);

                    using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
                    var writer = new NativeFileWriter(stream);
                    return NativeMethods.FPDF_SaveAsCopy(document, ref writer.FileWrite, 0);
                }
                finally
                {
                    if (ocrFont != IntPtr.Zero)
                        NativeMethods.FPDFFont_Close(ocrFont);

                    NativeMethods.FPDF_CloseDocument(document);
                }
            }
        }

        private static void DeleteRemovedPages(IntPtr document, DocumentSnapshot snapshot)
        {
            var retainedPages = new HashSet<int>(
                snapshot.Pages
                    .Where(p => !p.IsBlankPage)
                    .Select(p => p.OriginalPageIndex));
            for (int sourceIndex = snapshot.SourcePageCount - 1; sourceIndex >= 0; sourceIndex--)
            {
                if (!retainedPages.Contains(sourceIndex))
                    NativeMethods.FPDFPage_Delete(document, sourceIndex);
            }
        }

        private static void InsertBlankPages(IntPtr document, DocumentSnapshot snapshot, ref IntPtr ocrFont)
        {
            for (int targetIndex = 0; targetIndex < snapshot.Pages.Count; targetIndex++)
            {
                var pageSnapshot = snapshot.Pages[targetIndex];
                if (!pageSnapshot.IsBlankPage)
                    continue;

                double width = pageSnapshot.PdfWidth > 0 ? pageSnapshot.PdfWidth : 595;
                double height = pageSnapshot.PdfHeight > 0 ? pageSnapshot.PdfHeight : 842;
                IntPtr page = NativeMethods.FPDFPage_New(document, targetIndex, width, height);
                if (page == IntPtr.Zero)
                    continue;

                try
                {
                    if (pageSnapshot.Rotation != 0)
                        NativeMethods.FPDFPage_SetRotation(page, pageSnapshot.Rotation / 90);

                    foreach (var annotation in pageSnapshot.Annotations)
                        AddPersistentObject(document, page, pageSnapshot, annotation);

                    AddOcrTextLayer(document, page, pageSnapshot, ref ocrFont);
                    NativeMethods.FPDFPage_GenerateContent(page);
                }
                finally
                {
                    NativeMethods.FPDF_ClosePage(page);
                }
            }
        }

        private static void RemoveEditableAnnotations(IntPtr page)
        {
            for (int i = NativeMethods.FPDFPage_GetAnnotCount(page) - 1; i >= 0; i--)
            {
                IntPtr annotation = NativeMethods.FPDFPage_GetAnnot(page, i);
                if (annotation == IntPtr.Zero)
                    continue;

                bool shouldRemove = false;
                try
                {
                    int subtype = NativeMethods.FPDFAnnot_GetSubtype(annotation);
                    shouldRemove = subtype == FPDF_ANNOT_FREETEXT ||
                                   subtype == FPDF_ANNOT_HIGHLIGHT ||
                                   subtype == FPDF_ANNOT_UNDERLINE ||
                                   (subtype == FPDF_ANNOT_STAMP &&
                                    (IsManagedFreeTextStamp(annotation) || IsManagedImageStamp(annotation)));
                }
                finally
                {
                    NativeMethods.FPDFPage_CloseAnnot(annotation);
                }

                if (shouldRemove)
                    NativeMethods.FPDFPage_RemoveAnnot(page, i);
            }
        }

        private static void AddPersistentObject(IntPtr document, IntPtr page, PageSnapshot pageSnapshot, AnnotationSnapshot annotation)
        {
            if (annotation.Type == AnnotationType.ImageStamp)
            {
                AddImageAnnotation(document, page, pageSnapshot, annotation);
                return;
            }

            AddAnnotation(document, page, pageSnapshot, annotation);
        }

        private static void AddAnnotation(IntPtr document, IntPtr page, PageSnapshot pageSnapshot, AnnotationSnapshot annotation)
        {
            int subtype = annotation.Type switch
            {
                AnnotationType.FreeText => FPDF_ANNOT_STAMP,
                AnnotationType.Highlight => FPDF_ANNOT_HIGHLIGHT,
                AnnotationType.Underline => FPDF_ANNOT_UNDERLINE,
                _ => 0
            };

            if (subtype == 0 || annotation.Width <= 0 || annotation.Height <= 0)
                return;

            IntPtr pdfAnnotation = NativeMethods.FPDFPage_CreateAnnot(page, subtype);
            if (pdfAnnotation == IntPtr.Zero)
                return;

            try
            {
                var rect = ToPdfRect(pageSnapshot, annotation);
                NativeMethods.FPDFAnnot_SetRect(pdfAnnotation, ref rect);
                NativeMethods.FPDFAnnot_SetFlags(pdfAnnotation, 4);

                if (annotation.Type == AnnotationType.FreeText)
                {
                    NativeMethods.FPDFAnnot_SetStringValue(pdfAnnotation, "Contents", EncodeManagedFreeText(annotation));
                    AddFreeTextAppearance(document, pdfAnnotation, rect, annotation);
                }
                else
                {
                    Color color = annotation.Type == AnnotationType.Underline
                        ? Colors.Black
                        : annotation.Background;

                    NativeMethods.FPDFAnnot_SetColor(
                        pdfAnnotation,
                        FPDFANNOT_COLORTYPE_Color,
                        color.R,
                        color.G,
                        color.B,
                        color.A == 0 ? (byte)80 : color.A);

                    var quadPoints = ToQuadPoints(rect);
                    NativeMethods.FPDFAnnot_SetAttachmentPoints(pdfAnnotation, 0, ref quadPoints);
                }
            }
            finally
            {
                NativeMethods.FPDFPage_CloseAnnot(pdfAnnotation);
            }
        }

        private static void AddImageAnnotation(IntPtr document, IntPtr page, PageSnapshot pageSnapshot, AnnotationSnapshot annotation)
        {
            if (annotation.ImageBytes == null || annotation.ImageBytes.Length == 0 ||
                annotation.Width <= 0 || annotation.Height <= 0)
                return;

            IntPtr pdfAnnotation = NativeMethods.FPDFPage_CreateAnnot(page, FPDF_ANNOT_STAMP);
            if (pdfAnnotation == IntPtr.Zero)
                return;

            try
            {
                var rect = ToPdfRect(pageSnapshot, annotation);
                NativeMethods.FPDFAnnot_SetRect(pdfAnnotation, ref rect);
                NativeMethods.FPDFAnnot_SetFlags(pdfAnnotation, 4);
                NativeMethods.FPDFAnnot_SetStringValue(pdfAnnotation, "Contents", EncodeManagedImageStamp(annotation));
                AddImageAppearance(document, page, pdfAnnotation, rect, annotation.ImageBytes);
            }
            finally
            {
                NativeMethods.FPDFPage_CloseAnnot(pdfAnnotation);
            }
        }

        private static void AddImageAppearance(IntPtr document, IntPtr page, IntPtr annotation, NativeMethods.FS_RECTF rect, byte[] imageBytes)
        {
            IntPtr imageObject = NativeMethods.FPDFPageObj_NewImageObj(document);
            if (imageObject == IntPtr.Zero)
                return;

            bool appended = false;
            IntPtr pdfBitmap = IntPtr.Zero;
            try
            {
                IntPtr[] pages = { page };
                bool imageLoaded = TryLoadJpegImageObject(pages, imageObject, imageBytes);
                if (!imageLoaded)
                {
                    pdfBitmap = CreatePdfBitmap(imageBytes);
                    if (pdfBitmap == IntPtr.Zero)
                    {
                        LogToFile($"ImageStamp skipped: bitmap creation failed. Bytes={imageBytes.Length}");
                        return;
                    }

                    imageLoaded = NativeMethods.FPDFImageObj_SetBitmap(pages, pages.Length, imageObject, pdfBitmap);
                    if (!imageLoaded)
                    {
                        LogToFile($"ImageStamp skipped: FPDFImageObj_SetBitmap failed. Bytes={imageBytes.Length}");
                        return;
                    }
                }

                double width = Math.Max(0.1, rect.right - rect.left);
                double height = Math.Max(0.1, rect.top - rect.bottom);
                NativeMethods.FPDFPageObj_Transform(imageObject, width, 0, 0, height, rect.left, rect.bottom);
                appended = NativeMethods.FPDFAnnot_AppendObject(annotation, imageObject);
                LogToFile($"ImageStamp annotation inserted. Rect=({rect.left},{rect.bottom},{width},{height}), Bytes={imageBytes.Length}, Appended={appended}");
            }
            finally
            {
                if (!appended)
                    NativeMethods.FPDFPageObj_Destroy(imageObject);
                if (pdfBitmap != IntPtr.Zero)
                    NativeMethods.FPDFBitmap_Destroy(pdfBitmap);
            }
        }

        private static bool TryLoadJpegImageObject(IntPtr[] pages, IntPtr imageObject, byte[] imageBytes)
        {
            if (!IsJpeg(imageBytes))
                return false;

            using var access = new NativeImageFileAccess(imageBytes);
            return NativeMethods.FPDFImageObj_LoadJpegFileInline(
                pages,
                pages.Length,
                imageObject,
                ref access.FileAccess);
        }

        private static bool IsJpeg(byte[] imageBytes)
        {
            return imageBytes.Length > 3 &&
                   imageBytes[0] == 0xFF &&
                   imageBytes[1] == 0xD8 &&
                   imageBytes[^2] == 0xFF &&
                   imageBytes[^1] == 0xD9;
        }

        private static string EncodeManagedImageStamp(AnnotationSnapshot snapshot)
        {
            var metadata = new ManagedImageStampMetadata
            {
                Storage = "appearance"
            };
            string json = JsonSerializer.Serialize(metadata);
            return ImageStampMetadataV2Prefix + Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        }

        private static IntPtr CreatePdfBitmap(byte[] imageBytes)
        {
            using var input = new MemoryStream(imageBytes);
            using var source = new DrawingBitmap(input);
            using var bitmap = new DrawingBitmap(source.Width, source.Height, DrawingPixelFormat.Format32bppArgb);
            using (var graphics = DrawingGraphics.FromImage(bitmap))
            {
                graphics.Clear(System.Drawing.Color.White);
                graphics.DrawImage(source, 0, 0, bitmap.Width, bitmap.Height);
            }

            IntPtr pdfBitmap = NativeMethods.FPDFBitmap_CreateEx(bitmap.Width, bitmap.Height, 4, IntPtr.Zero, 0);
            if (pdfBitmap == IntPtr.Zero)
                return IntPtr.Zero;

            try
            {
                IntPtr targetBuffer = NativeMethods.FPDFBitmap_GetBuffer(pdfBitmap);
                int targetStride = NativeMethods.FPDFBitmap_GetStride(pdfBitmap);
                if (targetBuffer == IntPtr.Zero || targetStride <= 0)
                    return IntPtr.Zero;

                var rect = new DrawingRectangle(0, 0, bitmap.Width, bitmap.Height);
                var data = bitmap.LockBits(rect, DrawingImageLockMode.ReadOnly, DrawingPixelFormat.Format32bppArgb);
                try
                {
                    int rowBytes = bitmap.Width * 4;
                    byte[] row = new byte[rowBytes];
                    for (int y = 0; y < bitmap.Height; y++)
                    {
                        Marshal.Copy(IntPtr.Add(data.Scan0, y * data.Stride), row, 0, rowBytes);
                        Marshal.Copy(row, 0, IntPtr.Add(targetBuffer, y * targetStride), rowBytes);
                    }
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }

                IntPtr result = pdfBitmap;
                pdfBitmap = IntPtr.Zero;
                return result;
            }
            finally
            {
                if (pdfBitmap != IntPtr.Zero)
                    NativeMethods.FPDFBitmap_Destroy(pdfBitmap);
            }
        }

        private static void AddOcrTextLayer(IntPtr document, IntPtr page, PageSnapshot pageSnapshot, ref IntPtr font)
        {
            if (pageSnapshot.OcrWords.Count == 0)
                return;

            EnsureOcrFont(document, ref font);
            if (font == IntPtr.Zero)
                return;

            foreach (var word in pageSnapshot.OcrWords)
            {
                AddOcrWord(document, page, font, pageSnapshot, word);
            }
        }

        private static void EnsureOcrFont(IntPtr document, ref IntPtr font)
        {
            if (font != IntPtr.Zero)
                return;

            string fontPath = GetKoreanFontPath();
            if (string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
                return;

            byte[] fontBytes = File.ReadAllBytes(fontPath);
            font = NativeMethods.FPDFText_LoadFont(document, fontBytes, (uint)fontBytes.Length, 2, true);
        }

        private static void AddOcrWord(IntPtr document, IntPtr page, IntPtr font, PageSnapshot pageSnapshot, OcrWordSnapshot word)
        {
            var rect = ToPdfRect(pageSnapshot, word);
            float targetWidth = Math.Max(0.1f, rect.right - rect.left);
            float targetHeight = Math.Max(0.1f, rect.top - rect.bottom);
            float fontSize = Math.Max(1.0f, Math.Min(targetHeight, 72.0f));

            IntPtr textObject = NativeMethods.FPDFPageObj_CreateTextObj(document, font, fontSize);
            if (textObject == IntPtr.Zero)
                return;

            bool inserted = false;
            try
            {
                if (!NativeMethods.FPDFText_SetText(textObject, word.Text))
                    return;

                // Alpha 0 keeps the OCR layer invisible while preserving text for search/extraction.
                NativeMethods.FPDFText_SetFillColor(textObject, 0, 0, 0, 0);

                double horizontalScale = 1.0;
                if (NativeMethods.FPDFPageObj_GetBounds(textObject, out float left, out _, out float right, out _))
                {
                    float naturalWidth = right - left;
                    if (naturalWidth > 0.1f)
                    {
                        horizontalScale = targetWidth / naturalWidth;
                        horizontalScale = Math.Max(0.25, Math.Min(4.0, horizontalScale));
                    }
                }

                double x = rect.left;
                double y = rect.bottom + (fontSize * 0.12);
                NativeMethods.FPDFPageObj_Transform(textObject, horizontalScale, 0, 0, 1, x, y);
                NativeMethods.FPDFPage_InsertObject(page, textObject);
                inserted = true;
            }
            finally
            {
                if (!inserted)
                    NativeMethods.FPDFPageObj_Destroy(textObject);
            }
        }

        private static bool IsManagedFreeTextStamp(IntPtr annotation)
        {
            string contents = GetAnnotationStringValue(annotation, "Contents");
            return contents.StartsWith(FreeTextMetadataPrefix, StringComparison.Ordinal) ||
                   contents.StartsWith(FreeTextMetadataV2Prefix, StringComparison.Ordinal);
        }

        private static bool IsManagedImageStamp(IntPtr annotation)
        {
            string contents = GetAnnotationStringValue(annotation, "Contents");
            return IsManagedImageStampContents(contents);
        }

        internal static bool TryDecodeManagedFreeText(string contents, out string text)
        {
            text = string.Empty;
            if (TryDecodeManagedFreeText(contents, out ManagedFreeTextMetadata metadata))
            {
                text = metadata.Text;
                return true;
            }

            if (!contents.StartsWith(FreeTextMetadataPrefix, StringComparison.Ordinal))
                return false;

            try
            {
                string payload = contents.Substring(FreeTextMetadataPrefix.Length);
                text = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal static bool TryDecodeManagedFreeText(string contents, out ManagedFreeTextMetadata metadata)
        {
            metadata = new ManagedFreeTextMetadata();
            if (!contents.StartsWith(FreeTextMetadataV2Prefix, StringComparison.Ordinal))
                return false;

            try
            {
                string payload = contents.Substring(FreeTextMetadataV2Prefix.Length);
                string json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                metadata = JsonSerializer.Deserialize<ManagedFreeTextMetadata>(json) ?? new ManagedFreeTextMetadata();
                return true;
            }
            catch
            {
                metadata = new ManagedFreeTextMetadata();
                return false;
            }
        }

        internal static bool TryDecodeManagedImageStamp(string contents, out byte[] imageBytes)
        {
            imageBytes = Array.Empty<byte>();
            if (!contents.StartsWith(ImageStampMetadataPrefix, StringComparison.Ordinal))
                return false;

            try
            {
                string payload = contents.Substring(ImageStampMetadataPrefix.Length);
                string json = Encoding.UTF8.GetString(Convert.FromBase64String(payload));
                var metadata = JsonSerializer.Deserialize<ManagedImageStampMetadata>(json);
                if (string.IsNullOrWhiteSpace(metadata?.ImageBase64))
                    return false;

                imageBytes = Convert.FromBase64String(metadata.ImageBase64);
                return imageBytes.Length > 0;
            }
            catch
            {
                imageBytes = Array.Empty<byte>();
                return false;
            }
        }

        internal static bool IsManagedImageStampContents(string contents)
        {
            return contents.StartsWith(ImageStampMetadataPrefix, StringComparison.Ordinal) ||
                   contents.StartsWith(ImageStampMetadataV2Prefix, StringComparison.Ordinal);
        }

        internal static bool TryExtractManagedImageStampBytes(IntPtr annotation, string contents, out byte[] imageBytes)
        {
            imageBytes = Array.Empty<byte>();

            if (TryDecodeManagedImageStamp(contents, out imageBytes))
                return true;

            if (!contents.StartsWith(ImageStampMetadataV2Prefix, StringComparison.Ordinal))
                return false;

            return TryExtractFirstAnnotationImage(annotation, out imageBytes);
        }

        private static bool TryExtractFirstAnnotationImage(IntPtr annotation, out byte[] imageBytes)
        {
            imageBytes = Array.Empty<byte>();

            int objectCount = NativeMethods.FPDFAnnot_GetObjectCount(annotation);
            for (int i = 0; i < objectCount; i++)
            {
                IntPtr pageObject = NativeMethods.FPDFAnnot_GetObject(annotation, i);
                if (pageObject == IntPtr.Zero ||
                    NativeMethods.FPDFPageObj_GetType(pageObject) != FPDF_PAGEOBJ_IMAGE)
                {
                    continue;
                }

                if (TryGetDecodableRawImageBytes(pageObject, out imageBytes))
                    return true;

                if (TryRenderImageObjectToPng(pageObject, out imageBytes))
                    return true;
            }

            return false;
        }

        private static bool TryGetDecodableRawImageBytes(IntPtr imageObject, out byte[] imageBytes)
        {
            imageBytes = Array.Empty<byte>();

            ulong size = NativeMethods.FPDFImageObj_GetImageDataRaw(imageObject, IntPtr.Zero, 0);
            if (size == 0 || size > int.MaxValue)
                return false;

            IntPtr buffer = Marshal.AllocHGlobal(checked((int)size));
            try
            {
                ulong written = NativeMethods.FPDFImageObj_GetImageDataRaw(imageObject, buffer, size);
                if (written == 0 || written > size || written > int.MaxValue)
                    return false;

                byte[] raw = new byte[written];
                Marshal.Copy(buffer, raw, 0, checked((int)written));
                if (!CanDecodeImage(raw))
                    return false;

                imageBytes = raw;
                return true;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }

        private static bool TryRenderImageObjectToPng(IntPtr imageObject, out byte[] imageBytes)
        {
            imageBytes = Array.Empty<byte>();

            IntPtr pdfBitmap = NativeMethods.FPDFImageObj_GetBitmap(imageObject);
            if (pdfBitmap == IntPtr.Zero)
                return false;

            try
            {
                int width = NativeMethods.FPDFBitmap_GetWidth(pdfBitmap);
                int height = NativeMethods.FPDFBitmap_GetHeight(pdfBitmap);
                int stride = NativeMethods.FPDFBitmap_GetStride(pdfBitmap);
                IntPtr sourceBuffer = NativeMethods.FPDFBitmap_GetBuffer(pdfBitmap);
                if (width <= 0 || height <= 0 || stride <= 0 || sourceBuffer == IntPtr.Zero)
                    return false;

                using var bitmap = new DrawingBitmap(width, height, DrawingPixelFormat.Format32bppArgb);
                var rect = new DrawingRectangle(0, 0, width, height);
                var data = bitmap.LockBits(rect, DrawingImageLockMode.WriteOnly, DrawingPixelFormat.Format32bppArgb);
                try
                {
                    int rowBytes = width * 4;
                    byte[] row = new byte[rowBytes];
                    for (int y = 0; y < height; y++)
                    {
                        Marshal.Copy(IntPtr.Add(sourceBuffer, y * stride), row, 0, rowBytes);
                        Marshal.Copy(row, 0, IntPtr.Add(data.Scan0, y * data.Stride), rowBytes);
                    }
                }
                finally
                {
                    bitmap.UnlockBits(data);
                }

                using var output = new MemoryStream();
                bitmap.Save(output, DrawingImageFormat.Png);
                imageBytes = output.ToArray();
                return imageBytes.Length > 0;
            }
            finally
            {
                NativeMethods.FPDFBitmap_Destroy(pdfBitmap);
            }
        }

        private static bool CanDecodeImage(byte[] imageBytes)
        {
            try
            {
                using var stream = new MemoryStream(imageBytes);
                using var _ = new DrawingBitmap(stream);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string EncodeManagedFreeText(AnnotationSnapshot snapshot)
        {
            var metadata = new ManagedFreeTextMetadata
            {
                Text = snapshot.Text ?? string.Empty,
                FontFamily = snapshot.FontFamily ?? "Malgun Gothic",
                FontSize = snapshot.FontSize,
                IsBold = snapshot.IsBold,
                ForegroundA = snapshot.Foreground.A,
                ForegroundR = snapshot.Foreground.R,
                ForegroundG = snapshot.Foreground.G,
                ForegroundB = snapshot.Foreground.B
            };
            string json = JsonSerializer.Serialize(metadata);
            return FreeTextMetadataV2Prefix + Convert.ToBase64String(Encoding.UTF8.GetBytes(json));
        }

        private static void AddFreeTextAppearance(IntPtr document, IntPtr annotation, NativeMethods.FS_RECTF rect, AnnotationSnapshot snapshot)
        {
            string fontPath = GetFontPath(snapshot.FontFamily, snapshot.IsBold);
            if (string.IsNullOrEmpty(fontPath) || !File.Exists(fontPath))
                return;

            byte[] fontBytes = File.ReadAllBytes(fontPath);
            IntPtr font = NativeMethods.FPDFText_LoadFont(document, fontBytes, (uint)fontBytes.Length, 2, true);
            if (font == IntPtr.Zero)
                return;

            try
            {
                float fontSize = (float)Math.Max(1.0, Math.Min(snapshot.FontSize, snapshot.Height));
                IntPtr textObject = NativeMethods.FPDFPageObj_CreateTextObj(document, font, fontSize);
                if (textObject == IntPtr.Zero)
                    return;

                bool appended = false;
                try
                {
                    if (!NativeMethods.FPDFText_SetText(textObject, snapshot.Text))
                        return;

                    NativeMethods.FPDFText_SetFillColor(
                        textObject,
                        snapshot.Foreground.R,
                        snapshot.Foreground.G,
                        snapshot.Foreground.B,
                        snapshot.Foreground.A == 0 ? 255u : snapshot.Foreground.A);

                    double x = rect.left + 2;
                    double y = rect.top - fontSize - 2;
                    NativeMethods.FPDFPageObj_Transform(textObject, 1, 0, 0, 1, x, y);

                    appended = NativeMethods.FPDFAnnot_AppendObject(annotation, textObject);
                }
                finally
                {
                    if (!appended)
                        NativeMethods.FPDFPageObj_Destroy(textObject);
                }
            }
            finally
            {
                NativeMethods.FPDFFont_Close(font);
            }
        }

        private static string GetFontPath(string? fontFamily, bool isBold)
        {
            string fontFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            string family = (fontFamily ?? string.Empty).Replace(" ", "", StringComparison.Ordinal).ToLowerInvariant();

            var candidates = new List<string>();
            if (family.Contains("malgun") || family.Contains("맑은"))
                candidates.Add(Path.Combine(fontFolder, isBold ? "malgunbd.ttf" : "malgun.ttf"));
            else if (family.Contains("gulim") || family.Contains("굴림"))
                candidates.Add(Path.Combine(fontFolder, "gulim.ttc"));
            else if (family.Contains("batang") || family.Contains("바탕"))
                candidates.Add(Path.Combine(fontFolder, "batang.ttc"));

            candidates.Add(Path.Combine(fontFolder, isBold ? "malgunbd.ttf" : "malgun.ttf"));
            candidates.Add(Path.Combine(fontFolder, "malgun.ttf"));
            candidates.Add(Path.Combine(fontFolder, "gulim.ttc"));
            candidates.Add(Path.Combine(fontFolder, "batang.ttc"));

            return candidates.FirstOrDefault(File.Exists) ?? string.Empty;
        }

        private static string GetKoreanFontPath()
        {
            return GetFontPath("Malgun Gothic", false);
        }

        private static string GetAnnotationStringValue(IntPtr annot, string key)
        {
            ulong len = NativeMethods.FPDFAnnot_GetStringValue(annot, key, IntPtr.Zero, 0);
            if (len == 0) return string.Empty;
            IntPtr buffer = Marshal.AllocHGlobal((int)len);
            try
            {
                NativeMethods.FPDFAnnot_GetStringValue(annot, key, buffer, len);
                return Marshal.PtrToStringUni(buffer) ?? string.Empty;
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
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

        private static int NormalizeRotation(int rotation)
        {
            rotation %= 360;
            if (rotation < 0)
                rotation += 360;
            return rotation is 90 or 180 or 270 ? rotation : 0;
        }

        private static void LogToFile(string message)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:HH:mm:ss.fff}] [PdfiumEditService] {message}{Environment.NewLine}");
            }
            catch { }
        }

        private static NativeMethods.FS_RECTF ToPdfRect(PageSnapshot page, AnnotationSnapshot annotation)
        {
            double scaleX = page.PdfWidth / page.Width;
            double scaleY = page.PdfHeight / page.Height;
            float left = (float)(annotation.X * scaleX);
            float right = (float)((annotation.X + annotation.Width) * scaleX);
            float top = (float)(page.PdfHeight - (annotation.Y * scaleY));
            float bottom = (float)(page.PdfHeight - ((annotation.Y + annotation.Height) * scaleY));

            return new NativeMethods.FS_RECTF
            {
                left = Math.Min(left, right),
                right = Math.Max(left, right),
                top = Math.Max(top, bottom),
                bottom = Math.Min(top, bottom)
            };
        }

        private static NativeMethods.FS_RECTF ToPdfRect(PageSnapshot page, OcrWordSnapshot word)
        {
            double scaleX = page.PdfWidth / page.Width;
            double scaleY = page.PdfHeight / page.Height;
            float left = (float)(word.X * scaleX);
            float right = (float)((word.X + word.Width) * scaleX);
            float top = (float)(page.PdfHeight - (word.Y * scaleY));
            float bottom = (float)(page.PdfHeight - ((word.Y + word.Height) * scaleY));

            return new NativeMethods.FS_RECTF
            {
                left = Math.Min(left, right),
                right = Math.Max(left, right),
                top = Math.Max(top, bottom),
                bottom = Math.Min(top, bottom)
            };
        }

        private static NativeMethods.FS_QUADPOINTSF ToQuadPoints(NativeMethods.FS_RECTF rect)
        {
            return new NativeMethods.FS_QUADPOINTSF
            {
                x1 = rect.left,
                y1 = rect.top,
                x2 = rect.right,
                y2 = rect.top,
                x3 = rect.left,
                y3 = rect.bottom,
                x4 = rect.right,
                y4 = rect.bottom
            };
        }

        private sealed record DocumentSnapshot(string SourcePath, int SourcePageCount, List<PageSnapshot> Pages);

        private sealed class PageSnapshot
        {
            public int OriginalPageIndex { get; init; }
            public bool IsBlankPage { get; init; }
            public double Width { get; init; }
            public double Height { get; init; }
            public double PdfWidth { get; init; }
            public double PdfHeight { get; init; }
            public int Rotation { get; init; }
            public bool ShouldRewriteAnnotations { get; init; }
            public List<AnnotationSnapshot> Annotations { get; } = new();
            public List<OcrWordSnapshot> OcrWords { get; } = new();
        }

        private sealed class AnnotationSnapshot
        {
            public AnnotationType Type { get; init; }
            public double X { get; init; }
            public double Y { get; init; }
            public double Width { get; init; }
            public double Height { get; init; }
            public string Text { get; init; } = string.Empty;
            public double FontSize { get; init; } = 12;
            public string FontFamily { get; init; } = "Malgun Gothic";
            public bool IsBold { get; init; }
            public Color Foreground { get; init; }
            public Color Background { get; init; }
            public byte[]? ImageBytes { get; init; }
        }

        internal sealed class ManagedFreeTextMetadata
        {
            public string Text { get; set; } = string.Empty;
            public string FontFamily { get; set; } = "Malgun Gothic";
            public double FontSize { get; set; } = 12;
            public bool IsBold { get; set; }
            public byte ForegroundA { get; set; } = 255;
            public byte ForegroundR { get; set; }
            public byte ForegroundG { get; set; }
            public byte ForegroundB { get; set; }

            public Color ForegroundColor => Color.FromArgb(
                ForegroundA == 0 ? (byte)255 : ForegroundA,
                ForegroundR,
                ForegroundG,
                ForegroundB);
        }

        private sealed class ManagedImageStampMetadata
        {
            public string Storage { get; set; } = string.Empty;
            public string ImageBase64 { get; set; } = string.Empty;
        }

        private sealed class OcrWordSnapshot
        {
            public string Text { get; init; } = string.Empty;
            public double X { get; init; }
            public double Y { get; init; }
            public double Width { get; init; }
            public double Height { get; init; }
        }

        private sealed class NativeFileWriter
        {
            private readonly Stream _stream;
            private readonly NativeMethods.WriteBlockCallback _callback;

            public NativeFileWriter(Stream stream)
            {
                _stream = stream;
                _callback = WriteBlock;
                FileWrite = new NativeMethods.FPDF_FILEWRITE
                {
                    Version = 1,
                    WriteBlock = _callback
                };
            }

            public NativeMethods.FPDF_FILEWRITE FileWrite;

            private int WriteBlock(IntPtr fileWrite, IntPtr data, uint size)
            {
                byte[] buffer = new byte[size];
                Marshal.Copy(data, buffer, 0, checked((int)size));
                _stream.Write(buffer, 0, buffer.Length);
                return 1;
            }
        }

        private sealed class NativeImageFileAccess : IDisposable
        {
            private readonly byte[] _data;
            private readonly GCHandle _handle;
            private readonly NativeMethods.GetBlockCallback _callback;

            public NativeImageFileAccess(byte[] data)
            {
                _data = data;
                _callback = ReadBlock;
                _handle = GCHandle.Alloc(this);
                FileAccess = new NativeMethods.FPDF_FILEACCESS
                {
                    FileLen = checked((uint)data.Length),
                    GetBlock = _callback,
                    Param = GCHandle.ToIntPtr(_handle)
                };
            }

            public NativeMethods.FPDF_FILEACCESS FileAccess;

            private static int ReadBlock(IntPtr param, uint position, IntPtr buffer, uint size)
            {
                if (param == IntPtr.Zero || buffer == IntPtr.Zero)
                    return 0;

                var handle = GCHandle.FromIntPtr(param);
                if (handle.Target is not NativeImageFileAccess access)
                    return 0;

                if (position > access._data.Length || size > access._data.Length - position)
                    return 0;

                Marshal.Copy(access._data, checked((int)position), buffer, checked((int)size));
                return 1;
            }

            public void Dispose()
            {
                if (_handle.IsAllocated)
                    _handle.Free();
            }
        }

        private static class NativeMethods
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct FS_RECTF
            {
                public float left;
                public float top;
                public float right;
                public float bottom;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct FS_QUADPOINTSF
            {
                public float x1;
                public float y1;
                public float x2;
                public float y2;
                public float x3;
                public float y3;
                public float x4;
                public float y4;
            }

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            public delegate int WriteBlockCallback(IntPtr fileWrite, IntPtr data, uint size);

            [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
            public delegate int GetBlockCallback(IntPtr param, uint position, IntPtr buffer, uint size);

            [StructLayout(LayoutKind.Sequential)]
            public struct FPDF_FILEWRITE
            {
                public int Version;
                public WriteBlockCallback WriteBlock;
            }

            [StructLayout(LayoutKind.Sequential)]
            public struct FPDF_FILEACCESS
            {
                public uint FileLen;
                public GetBlockCallback GetBlock;
                public IntPtr Param;
            }

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDF_LoadDocument(
                [MarshalAs(UnmanagedType.LPStr)] string path,
                [MarshalAs(UnmanagedType.LPStr)] string? password);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDF_CloseDocument(IntPtr document);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDF_LoadPage(IntPtr document, int pageIndex);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDF_ClosePage(IntPtr page);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFPage_GetAnnotCount(IntPtr page);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPage_GetAnnot(IntPtr page, int index);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPage_CloseAnnot(IntPtr annotation);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFAnnot_GetSubtype(IntPtr annotation);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFPage_RemoveAnnot(IntPtr page, int index);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPage_CreateAnnot(IntPtr page, int subtype);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_SetRect(IntPtr annotation, ref FS_RECTF rect);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_SetStringValue(
                IntPtr annotation,
                [MarshalAs(UnmanagedType.LPStr)] string key,
                [MarshalAs(UnmanagedType.LPWStr)] string value);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern ulong FPDFAnnot_GetStringValue(
                IntPtr annotation,
                [MarshalAs(UnmanagedType.LPStr)] string key,
                IntPtr buffer,
                ulong buflen);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_SetColor(IntPtr annotation, int colorType, uint r, uint g, uint b, uint a);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_SetAttachmentPoints(IntPtr annotation, nuint quadIndex, ref FS_QUADPOINTSF quadPoints);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_SetFlags(IntPtr annotation, int flags);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFPage_GenerateContent(IntPtr page);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDF_SaveAsCopy(IntPtr document, ref FPDF_FILEWRITE fileWrite, uint flags);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPage_Delete(IntPtr document, int pageIndex);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPage_New(IntPtr document, int pageIndex, double width, double height);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPage_SetRotation(IntPtr page, int rotate);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFText_LoadFont(IntPtr document, byte[] data, uint size, int fontType, bool cid);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFFont_Close(IntPtr font);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPageObj_CreateTextObj(IntPtr document, IntPtr font, float fontSize);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPageObj_NewImageObj(IntPtr document);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFBitmap_CreateEx(int width, int height, int format, IntPtr firstScan, int stride);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFBitmap_GetBuffer(IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFBitmap_GetStride(IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFBitmap_Destroy(IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFImageObj_SetBitmap(
                [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] IntPtr[] pages,
                int count,
                IntPtr imageObject,
                IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFImageObj_LoadJpegFileInline(
                [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] IntPtr[] pages,
                int count,
                IntPtr imageObject,
                ref FPDF_FILEACCESS fileAccess);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFText_SetText(IntPtr textObject, [MarshalAs(UnmanagedType.LPWStr)] string text);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFText_SetFillColor(IntPtr textObject, uint r, uint g, uint b, uint a);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFPageObj_GetBounds(
                IntPtr pageObject,
                out float left,
                out float bottom,
                out float right,
                out float top);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPageObj_Transform(IntPtr pageObject, double a, double b, double c, double d, double e, double f);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPage_InsertObject(IntPtr page, IntPtr pageObject);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern bool FPDFAnnot_AppendObject(IntPtr annotation, IntPtr pageObject);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFAnnot_GetObjectCount(IntPtr annotation);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFAnnot_GetObject(IntPtr annotation, int index);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFPageObj_GetType(IntPtr pageObject);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern ulong FPDFImageObj_GetImageDataRaw(IntPtr imageObject, IntPtr buffer, ulong buflen);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFImageObj_GetBitmap(IntPtr imageObject);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFBitmap_GetWidth(IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern int FPDFBitmap_GetHeight(IntPtr bitmap);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFPageObj_Destroy(IntPtr pageObject);
        }
    }
}
