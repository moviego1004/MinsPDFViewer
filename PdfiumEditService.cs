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

namespace MinsPDFViewer
{
    public sealed class PdfiumEditService
    {
        private const int FPDF_ANNOT_FREETEXT = 3;
        private const int FPDF_ANNOT_HIGHLIGHT = 9;
        private const int FPDF_ANNOT_UNDERLINE = 10;
        private const int FPDF_ANNOT_STAMP = 13;
        private const int FPDF_ANNOT_WIDGET = 12;
        private const int FPDFANNOT_COLORTYPE_Color = 0;
        internal const string FreeTextMetadataPrefix = "MINS_FREETEXT_V1:";
        internal const string FreeTextMetadataV2Prefix = "MINS_FREETEXT_V2:";

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

            if (model.Pages.Any(p => p.IsBlankPage))
            {
                reason = "document contains blank pages";
                return false;
            }

            int sourcePageCount;
            lock (PdfService.PdfiumLock)
            {
                sourcePageCount = model.PdfDocument?.PageCount ?? model.Pages.Count;
            }

            if (model.Pages.Count != sourcePageCount)
            {
                reason = "page set has changed";
                return false;
            }

            int previousOriginalPageIndex = -1;
            foreach (var page in model.Pages)
            {
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
            if (Application.Current == null || Application.Current.Dispatcher.CheckAccess())
                return CreateSnapshot(model);

            return await Application.Current.Dispatcher.InvokeAsync(() => CreateSnapshot(model));
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
                    Width = page.Width,
                    Height = page.Height,
                    PdfWidth = page.PdfPageWidthPoint,
                    PdfHeight = page.PdfPageHeightPoint,
                    ShouldRewriteAnnotations = page.AnnotationsLoaded ||
                                               page.Annotations.Any(IsPersistentAnnotation) ||
                                               (page.OcrWords != null && page.OcrWords.Count > 0)
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
                        Background = (annotation.Background as SolidColorBrush)?.Color ?? annotation.AnnotationColor
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
                   annotation.Type == AnnotationType.Underline;
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
                        if (!pageSnapshot.ShouldRewriteAnnotations)
                            continue;

                        if (pageSnapshot.Annotations.Count == 0 && pageSnapshot.OcrWords.Count == 0)
                            continue;

                        IntPtr page = NativeMethods.FPDF_LoadPage(document, pageSnapshot.OriginalPageIndex);
                        if (page == IntPtr.Zero)
                            return false;

                        try
                        {
                            RemoveEditableAnnotations(page);

                            foreach (var annotation in pageSnapshot.Annotations)
                            {
                                AddAnnotation(document, page, pageSnapshot, annotation);
                            }

                            AddOcrTextLayer(document, page, pageSnapshot, ref ocrFont);
                            NativeMethods.FPDFPage_GenerateContent(page);
                        }
                        finally
                        {
                            NativeMethods.FPDF_ClosePage(page);
                        }
                    }

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
                                   (subtype == FPDF_ANNOT_STAMP && IsManagedFreeTextStamp(annotation));
                }
                finally
                {
                    NativeMethods.FPDFPage_CloseAnnot(annotation);
                }

                if (shouldRemove)
                    NativeMethods.FPDFPage_RemoveAnnot(page, i);
            }
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
            public double Width { get; init; }
            public double Height { get; init; }
            public double PdfWidth { get; init; }
            public double PdfHeight { get; init; }
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

            [StructLayout(LayoutKind.Sequential)]
            public struct FPDF_FILEWRITE
            {
                public int Version;
                public WriteBlockCallback WriteBlock;
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
            public static extern IntPtr FPDFText_LoadFont(IntPtr document, byte[] data, uint size, int fontType, bool cid);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern void FPDFFont_Close(IntPtr font);

            [DllImport("pdfium.dll", CallingConvention = CallingConvention.Cdecl)]
            public static extern IntPtr FPDFPageObj_CreateTextObj(IntPtr document, IntPtr font, float fontSize);

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
            public static extern void FPDFPageObj_Destroy(IntPtr pageObject);
        }
    }
}
