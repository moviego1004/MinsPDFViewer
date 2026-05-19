using PdfSharpDoc = PdfSharp.Pdf.PdfDocument;

namespace MinsPDFViewer
{
    internal sealed class GenericPdfAnnotation : PdfSharp.Pdf.Annotations.PdfAnnotation
    {
        public GenericPdfAnnotation(PdfSharpDoc document) : base(document)
        {
        }
    }
}
