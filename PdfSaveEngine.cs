namespace MinsPDFViewer
{
    internal enum PdfSaveEngine
    {
        Pdfium,
        PdfiumWithPdfSharpBookmarkRewrite,
        PdfSharpLegacy,
        PdfSharpLegacyRasterFallback
    }
}
