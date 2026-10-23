using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.X509.Store;

namespace MinsPDFViewer
{
    public class PdfSignatureService
    {
        private const int EstimatedSignatureSize = 8192;
        private const char PlaceholderChar = 'A';
        private const string ByteRangePlaceholder = "[0000000000 0000000000 0000000000 0000000000]";

        public void SignPdf(string sourcePath, string destPath, SignatureConfig config, int pageIndex, SignaturePdfRect? customRect)
        {
            if (config == null || config.PrivateKey == null || config.Certificate == null)
                throw new ArgumentException("서명 설정이 올바르지 않습니다.");

            byte[] originalBytes = File.ReadAllBytes(sourcePath);
            string originalText = Encoding.Latin1.GetString(originalBytes);
            var trailer = ReadTrailer(originalText);
            var objects = ReadObjects(originalText);

            if (!objects.TryGetValue((trailer.RootObjectNumber, trailer.RootGeneration), out var rootObject))
                throw new InvalidOperationException("PDF catalog object was not found.");

            var pages = ReadPageReferences(objects, rootObject.Content);
            if (pages.Count == 0)
                throw new InvalidOperationException("PDF page tree could not be parsed.");

            if (pageIndex < 0 || pageIndex >= pages.Count)
                pageIndex = pages.Count - 1;

            var pageRef = pages[pageIndex];
            if (!objects.TryGetValue((pageRef.ObjectNumber, pageRef.Generation), out var pageObject))
                throw new InvalidOperationException("PDF page object was not found.");

            var pageSize = ReadMediaBox(pageObject.Content);
            var rect = customRect.HasValue && customRect.Value.Width > 0 && customRect.Value.Height > 0
                ? customRect.Value
                : new SignaturePdfRect(Math.Max(0, pageSize.Width - 170), 50, 120, 60);

            int nextObjectNumber = Math.Max(trailer.Size, objects.Keys.Max(k => k.Number) + 1);
            int sigObjectNumber = nextObjectNumber++;
            int widgetObjectNumber = nextObjectNumber++;
            int appearanceObjectNumber = nextObjectNumber++;
            int appearanceImageObjectNumber = nextObjectNumber++;
            int acroFormObjectNumber = nextObjectNumber++;

            string fieldName = $"Signature{Guid.NewGuid():N}".Substring(0, 17);
            string now = DateTimeOffset.Now.ToString("yyyyMMddHHmmsszzz", CultureInfo.InvariantCulture).Replace(":", "'");
            now += "'";

            string signatureDictionary =
                "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached " +
                $"/ByteRange {ByteRangePlaceholder} " +
                $"/Contents <{new string(PlaceholderChar, EstimatedSignatureSize * 2)}> " +
                $"/M (D:{EscapeLiteral(now)}) " +
                $"/Name {EncodePdfTextString(config.Certificate.Subject)} " +
                $"/Reason {EncodePdfTextString(config.Reason ?? string.Empty)} " +
                $"/Location {EncodePdfTextString(config.Location ?? string.Empty)} >>";

            string appearanceImageObject = BuildAppearanceImageObject(rect, config);
            string appearanceStream = $"q\n{Format(rect.Width)} 0 0 {Format(rect.Height)} 0 0 cm\n/Im1 Do\nQ";
            string appearanceObject =
                $"<< /Type /XObject /Subtype /Form /FormType 1 /BBox [0 0 {Format(rect.Width)} {Format(rect.Height)}] " +
                "/Matrix [1 0 0 1 0 0] " +
                $"/Resources << /XObject << /Im1 {appearanceImageObjectNumber} 0 R >> >> " +
                $"/Length {Encoding.ASCII.GetByteCount(appearanceStream)} >>\nstream\n{appearanceStream}\nendstream";

            string widget =
                "<< /Type /Annot /Subtype /Widget /FT /Sig " +
                $"/Rect [{Format(rect.X)} {Format(rect.Y)} {Format(rect.Right)} {Format(rect.Top)}] " +
                $"/T ({EscapeLiteral(fieldName)}) /F 132 /P {pageRef.ObjectNumber} {pageRef.Generation} R " +
                $"/V {sigObjectNumber} 0 R /AP << /N {appearanceObjectNumber} 0 R >> /AS /N >>";

            string rewrittenCatalog = RewriteCatalog(
                rootObject.Content,
                acroFormObjectNumber,
                widgetObjectNumber,
                objects,
                out var rewrittenAcroFormObject);
            string rewrittenPage = RewritePageAnnots(
                pageObject.Content,
                widgetObjectNumber,
                objects,
                out var rewrittenAnnotsObject);

            var appendedObjects = new List<(int Number, int Generation, string Content)>
            {
                (trailer.RootObjectNumber, trailer.RootGeneration, rewrittenCatalog),
                (pageRef.ObjectNumber, pageRef.Generation, rewrittenPage),
                (sigObjectNumber, 0, signatureDictionary),
                (widgetObjectNumber, 0, widget),
                (appearanceObjectNumber, 0, appearanceObject),
                (appearanceImageObjectNumber, 0, appearanceImageObject)
            };

            appendedObjects.Add(rewrittenAcroFormObject);

            if (rewrittenAnnotsObject != null)
                appendedObjects.Add(rewrittenAnnotsObject.Value);

            byte[] unsignedPdf = BuildIncrementalPdf(originalBytes, appendedObjects, trailer);
            byte[] signedPdf = SignPreparedPdf(unsignedPdf, config);
            File.WriteAllBytes(destPath, signedPdf);
        }

        private static byte[] BuildIncrementalPdf(
            byte[] originalBytes,
            List<(int Number, int Generation, string Content)> objects,
            TrailerInfo trailer)
        {
            var sb = new StringBuilder();
            var offsets = new List<(int Number, int Generation, long Offset)>();
            long currentOffset = originalBytes.Length;

            foreach (var obj in objects)
            {
                string wrapped = $"{obj.Number} {obj.Generation} obj\n{obj.Content}\nendobj\n";
                offsets.Add((obj.Number, obj.Generation, currentOffset));
                sb.Append(wrapped);
                currentOffset += Encoding.Latin1.GetByteCount(wrapped);
            }

            long xrefOffset = currentOffset;
            sb.Append("xref\n");
            foreach (var group in GroupConsecutive(offsets.OrderBy(o => o.Number).ToList()))
            {
                sb.Append($"{group.First().Number} {group.Count}\n");
                foreach (var entry in group)
                    sb.Append($"{entry.Offset:0000000000} {entry.Generation:00000} n \n");
            }

            int newSize = Math.Max(trailer.Size, objects.Max(o => o.Number) + 1);
            sb.Append("trailer\n<< ");
            sb.Append($"/Size {newSize} /Root {trailer.RootObjectNumber} {trailer.RootGeneration} R ");
            if (!string.IsNullOrWhiteSpace(trailer.IdArray))
                sb.Append($"/ID {trailer.IdArray} ");
            sb.Append($"/Prev {trailer.PreviousXrefOffset} >>\n");
            sb.Append("startxref\n");
            sb.Append($"{xrefOffset}\n%%EOF\n");

            byte[] appendBytes = Encoding.Latin1.GetBytes(sb.ToString());
            var result = new byte[originalBytes.Length + appendBytes.Length];
            Buffer.BlockCopy(originalBytes, 0, result, 0, originalBytes.Length);
            Buffer.BlockCopy(appendBytes, 0, result, originalBytes.Length, appendBytes.Length);
            return result;
        }

        private byte[] SignPreparedPdf(byte[] pdfBytes, SignatureConfig config)
        {
            byte[] placeholder = Encoding.ASCII.GetBytes(new string(PlaceholderChar, 80));
            int patternPos = FindBytes(pdfBytes, placeholder, 0);
            if (patternPos == -1)
                throw new Exception("서명 공간(Placeholder)을 찾을 수 없습니다.");

            int contentsStart = patternPos;
            int contentsOpen = contentsStart - 1;
            int contentsClose = contentsStart + (EstimatedSignatureSize * 2);
            if (contentsOpen < 0 || contentsClose >= pdfBytes.Length || pdfBytes[contentsOpen] != '<' || pdfBytes[contentsClose] != '>')
                throw new Exception("서명 Placeholder 범위를 찾을 수 없습니다.");

            byte[] byteRangePlaceholder = Encoding.ASCII.GetBytes(ByteRangePlaceholder);
            int byteRangeStart = FindBytes(pdfBytes, byteRangePlaceholder, 0);
            if (byteRangeStart == -1)
                throw new Exception("/ByteRange 예약 공간을 찾을 수 없습니다.");

            int excludeStart = contentsOpen;
            int excludeEnd = contentsClose + 1;
            string byteRange = string.Format(
                CultureInfo.InvariantCulture,
                "[{0,10} {1,10} {2,10} {3,10}]",
                0,
                excludeStart,
                excludeEnd,
                pdfBytes.Length - excludeEnd);

            byte[] byteRangeBytes = Encoding.ASCII.GetBytes(byteRange);
            if (byteRangeBytes.Length != byteRangePlaceholder.Length)
                throw new Exception("ByteRange 예약 길이가 맞지 않습니다.");

            Buffer.BlockCopy(byteRangeBytes, 0, pdfBytes, byteRangeStart, byteRangeBytes.Length);

            byte[] signedContent = new byte[excludeStart + (pdfBytes.Length - excludeEnd)];
            Buffer.BlockCopy(pdfBytes, 0, signedContent, 0, excludeStart);
            Buffer.BlockCopy(pdfBytes, excludeEnd, signedContent, excludeStart, pdfBytes.Length - excludeEnd);

            byte[] cmsSignature = GenerateCmsSignature(signedContent, config);
            string hexSig = Convert.ToHexString(cmsSignature);
            if (hexSig.Length > EstimatedSignatureSize * 2)
                throw new Exception("서명 공간 부족");

            byte[] hexBytes = Encoding.ASCII.GetBytes(hexSig);
            Buffer.BlockCopy(hexBytes, 0, pdfBytes, contentsStart, hexBytes.Length);
            for (int i = contentsStart + hexBytes.Length; i < contentsClose; i++)
                pdfBytes[i] = (byte)'0';

            return pdfBytes;
        }

        private byte[] GenerateCmsSignature(byte[] data, SignatureConfig config)
        {
            if (config.Certificate?.Certificate == null)
                throw new ArgumentException("인증서가 없습니다.");

            var certList = new List<X509Certificate> { config.Certificate.Certificate };
            var store = X509StoreFactory.Create("Certificate/Collection", new X509CollectionStoreParameters(certList));
            var gen = new CmsSignedDataGenerator();
            gen.AddSigner(config.PrivateKey, config.Certificate.Certificate, CmsSignedDataGenerator.DigestSha256);
            gen.AddCertificates(store);
            var msg = new CmsProcessableByteArray(data);
            var signedData = gen.Generate(msg, false);
            return signedData.GetEncoded();
        }

        private static string BuildAppearanceImageObject(SignaturePdfRect rect, SignatureConfig config)
        {
            int width = Math.Max(24, (int)Math.Ceiling(rect.Width * 3));
            int height = Math.Max(12, (int)Math.Ceiling(rect.Height * 3));
            byte[] rgb = RenderSignatureAppearance(width, height, config);
            byte[] compressed = CompressZlib(rgb);
            string stream = Encoding.Latin1.GetString(compressed);
            return $"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} " +
                   "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode " +
                   $"/Length {compressed.Length} >>\nstream\n{stream}\nendstream";
        }

        private static byte[] RenderSignatureAppearance(int width, int height, SignatureConfig config)
        {
            using var bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                graphics.Clear(System.Drawing.Color.White);

                using var borderPen = new Pen(System.Drawing.Color.FromArgb(190, 180, 0, 0), Math.Max(2, width / 120));
                graphics.DrawRectangle(borderPen, 1, 1, width - 3, height - 3);

                string signer = ExtractCn(config.Certificate.Subject);
                string title = "전자서명 완료";
                string signerLine = string.IsNullOrWhiteSpace(signer) ? "서명자 정보 없음" : $"서명자: {signer}";
                string dateLine = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                using var titleFont = new Font("Malgun Gothic", Math.Max(7, height / 7.0f), FontStyle.Bold, GraphicsUnit.Pixel);
                using var bodyFont = new Font("Malgun Gothic", Math.Max(6, height / 8.5f), FontStyle.Regular, GraphicsUnit.Pixel);
                using var brush = new SolidBrush(System.Drawing.Color.FromArgb(190, 180, 0, 0));
                using var format = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter,
                    FormatFlags = StringFormatFlags.NoWrap
                };

                float lineHeight = height / 4.0f;
                graphics.DrawString(title, titleFont, brush, new RectangleF(4, height * 0.12f, width - 8, lineHeight), format);
                graphics.DrawString(signerLine, bodyFont, brush, new RectangleF(4, height * 0.40f, width - 8, lineHeight), format);
                graphics.DrawString(dateLine, bodyFont, brush, new RectangleF(4, height * 0.64f, width - 8, lineHeight), format);
            }

            var data = new byte[width * height * 3];
            int offset = 0;
            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    System.Drawing.Color pixel = bitmap.GetPixel(x, y);
                    data[offset++] = pixel.R;
                    data[offset++] = pixel.G;
                    data[offset++] = pixel.B;
                }
            }
            return data;
        }

        private static byte[] CompressZlib(byte[] data)
        {
            using var output = new MemoryStream();
            using (var zlib = new ZLibStream(output, CompressionLevel.Optimal, leaveOpen: true))
                zlib.Write(data, 0, data.Length);
            return output.ToArray();
        }

        private static string RewriteCatalog(
            string catalogContent,
            int acroFormObjectNumber,
            int widgetObjectNumber,
            Dictionary<(int Number, int Generation), PdfObjectInfo> objects,
            out (int Number, int Generation, string Content) rewrittenAcroFormObject)
        {
            string dictionary = ExtractDictionary(catalogContent, catalogContent.IndexOf("<<", StringComparison.Ordinal));
            var acroFormRefMatch = Regex.Match(dictionary, @"/AcroForm\s+(\d+)\s+(\d+)\s+R");
            if (acroFormRefMatch.Success)
            {
                int existingObjectNumber = int.Parse(acroFormRefMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                int existingGeneration = int.Parse(acroFormRefMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                if (objects.TryGetValue((existingObjectNumber, existingGeneration), out var existingAcroForm))
                {
                    string existingDictionary = ExtractDictionary(existingAcroForm.Content, existingAcroForm.Content.IndexOf("<<", StringComparison.Ordinal));
                    rewrittenAcroFormObject = (
                        existingObjectNumber,
                        existingGeneration,
                        MergeAcroFormFields(existingDictionary, widgetObjectNumber));
                    return dictionary;
                }
            }

            var acroFormDirectMatch = Regex.Match(dictionary, @"/AcroForm\s*(<<.*?>>)", RegexOptions.Singleline);
            string acroForm = "<< /Fields [] /SigFlags 3 >>";
            if (acroFormDirectMatch.Success)
            {
                acroForm = MergeAcroFormFields(acroFormDirectMatch.Groups[1].Value, widgetObjectNumber);
                dictionary = dictionary.Substring(0, acroFormDirectMatch.Index) +
                             dictionary.Substring(acroFormDirectMatch.Index + acroFormDirectMatch.Length);
            }
            else
            {
                acroForm = MergeAcroFormFields(acroForm, widgetObjectNumber);
            }

            int insert = dictionary.LastIndexOf(">>", StringComparison.Ordinal);
            if (insert < 0)
                throw new InvalidOperationException("PDF catalog dictionary could not be parsed.");
            rewrittenAcroFormObject = (acroFormObjectNumber, 0, acroForm);
            return dictionary.Insert(insert, $" /AcroForm {acroFormObjectNumber} 0 R ");
        }

        private static string MergeAcroFormFields(string acroFormDictionary, int widgetObjectNumber)
        {
            var fieldsMatch = Regex.Match(acroFormDictionary, @"/Fields\s*\[(.*?)\]", RegexOptions.Singleline);
            string widgetRef = $"{widgetObjectNumber} 0 R";
            if (fieldsMatch.Success)
            {
                string existing = fieldsMatch.Groups[1].Value.Trim();
                if (Regex.IsMatch(existing, $@"\b{widgetObjectNumber}\s+0\s+R\b"))
                    return acroFormDictionary;

                string replacement = string.IsNullOrWhiteSpace(existing)
                    ? $"/Fields [{widgetRef}]"
                    : $"/Fields [{existing} {widgetRef}]";
                acroFormDictionary = acroFormDictionary.Substring(0, fieldsMatch.Index) +
                                     replacement +
                                     acroFormDictionary.Substring(fieldsMatch.Index + fieldsMatch.Length);
            }
            else
            {
                int insert = acroFormDictionary.LastIndexOf(">>", StringComparison.Ordinal);
                if (insert < 0)
                    throw new InvalidOperationException("PDF AcroForm dictionary could not be parsed.");
                acroFormDictionary = acroFormDictionary.Insert(insert, $" /Fields [{widgetRef}] ");
            }

            if (Regex.IsMatch(acroFormDictionary, @"/SigFlags\s+\d+"))
                acroFormDictionary = Regex.Replace(acroFormDictionary, @"/SigFlags\s+\d+", "/SigFlags 3");
            else
            {
                int insert = acroFormDictionary.LastIndexOf(">>", StringComparison.Ordinal);
                if (insert < 0)
                    throw new InvalidOperationException("PDF AcroForm dictionary could not be parsed.");
                acroFormDictionary = acroFormDictionary.Insert(insert, " /SigFlags 3 ");
            }

            return acroFormDictionary;
        }

        private static string RewritePageAnnots(
            string pageContent,
            int widgetObjectNumber,
            Dictionary<(int Number, int Generation), PdfObjectInfo> objects,
            out (int Number, int Generation, string Content)? rewrittenAnnotsObject)
        {
            rewrittenAnnotsObject = null;
            string dictionary = ExtractDictionary(pageContent, pageContent.IndexOf("<<", StringComparison.Ordinal));
            int annotsKeyIndex = dictionary.IndexOf("/Annots", StringComparison.Ordinal);
            if (annotsKeyIndex >= 0)
            {
                int directArrayStart = dictionary.IndexOf('[', annotsKeyIndex);
                int nextReferenceStart = FindNextNonWhiteSpace(dictionary, annotsKeyIndex + "/Annots".Length);
                if (directArrayStart >= 0 && directArrayStart == nextReferenceStart)
                {
                    int directArrayEnd = FindMatchingArrayEnd(dictionary, directArrayStart);
                    string existing = dictionary.Substring(directArrayStart + 1, directArrayEnd - directArrayStart - 1).Trim();
                    string updated = string.IsNullOrWhiteSpace(existing)
                        ? $"[{widgetObjectNumber} 0 R]"
                        : $"[{existing} {widgetObjectNumber} 0 R]";
                    return dictionary.Substring(0, directArrayStart) + updated + dictionary.Substring(directArrayEnd + 1);
                }
            }

            var annotsRefMatch = Regex.Match(dictionary, @"/Annots\s+(\d+)\s+(\d+)\s+R");
            if (annotsRefMatch.Success)
            {
                int annotsObjectNumber = int.Parse(annotsRefMatch.Groups[1].Value, CultureInfo.InvariantCulture);
                int annotsGeneration = int.Parse(annotsRefMatch.Groups[2].Value, CultureInfo.InvariantCulture);
                if (objects.TryGetValue((annotsObjectNumber, annotsGeneration), out var annotsObject))
                {
                    string content = annotsObject.Content.Trim();
                    if (content.StartsWith("[", StringComparison.Ordinal) && content.EndsWith("]", StringComparison.Ordinal))
                    {
                        string existing = content.Substring(1, content.Length - 2).Trim();
                        string updated = string.IsNullOrWhiteSpace(existing)
                            ? $"[{widgetObjectNumber} 0 R]"
                            : $"[{existing} {widgetObjectNumber} 0 R]";
                        rewrittenAnnotsObject = (annotsObjectNumber, annotsGeneration, updated);
                        return dictionary;
                    }
                }
            }

            int insert = dictionary.LastIndexOf(">>", StringComparison.Ordinal);
            if (insert < 0)
                throw new InvalidOperationException("PDF page dictionary could not be parsed.");
            return dictionary.Insert(insert, $" /Annots [{widgetObjectNumber} 0 R] ");
        }

        private static int FindNextNonWhiteSpace(string text, int startIndex)
        {
            for (int i = Math.Max(0, startIndex); i < text.Length; i++)
            {
                if (!char.IsWhiteSpace(text[i]))
                    return i;
            }
            return -1;
        }

        private static int FindMatchingArrayEnd(string text, int arrayStart)
        {
            if (arrayStart < 0 || arrayStart >= text.Length || text[arrayStart] != '[')
                throw new InvalidOperationException("PDF array start was not found.");

            int depth = 0;
            for (int i = arrayStart; i < text.Length; i++)
            {
                if (text[i] == '(')
                {
                    i = SkipLiteralString(text, i);
                    continue;
                }

                if (text[i] == '[')
                {
                    depth++;
                    continue;
                }

                if (text[i] == ']')
                {
                    depth--;
                    if (depth == 0)
                        return i;
                }
            }

            throw new InvalidOperationException("PDF array end was not found.");
        }

        private static int SkipLiteralString(string text, int stringStart)
        {
            int depth = 0;
            for (int i = stringStart; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '\\')
                {
                    i++;
                    continue;
                }

                if (ch == '(')
                    depth++;
                else if (ch == ')')
                {
                    depth--;
                    if (depth == 0)
                        return i;
                }
            }

            return text.Length - 1;
        }

        private static TrailerInfo ReadTrailer(string text)
        {
            int startXref = text.LastIndexOf("startxref", StringComparison.Ordinal);
            if (startXref < 0)
                throw new InvalidOperationException("PDF startxref marker was not found.");

            var startXrefMatch = Regex.Match(text.Substring(startXref), @"startxref\s+(\d+)");
            if (!startXrefMatch.Success)
                throw new InvalidOperationException("PDF startxref offset could not be parsed.");

            long previousXrefOffset = long.Parse(startXrefMatch.Groups[1].Value, CultureInfo.InvariantCulture);
            int trailerIndex = text.IndexOf("trailer", (int)Math.Min(previousXrefOffset, int.MaxValue), StringComparison.Ordinal);
            if (trailerIndex < 0)
                trailerIndex = text.LastIndexOf("trailer", startXref, StringComparison.Ordinal);
            if (trailerIndex < 0)
                throw new InvalidOperationException("PDF trailer was not found.");

            string trailerDictionary = ExtractDictionary(text, text.IndexOf("<<", trailerIndex, StringComparison.Ordinal));
            var root = ReadReference(trailerDictionary, "Root");
            int size = ReadInteger(trailerDictionary, "Size");
            string? id = ReadRawArray(trailerDictionary, "ID");
            return new TrailerInfo(size, root.ObjectNumber, root.Generation, previousXrefOffset, id);
        }

        private static Dictionary<(int Number, int Generation), PdfObjectInfo> ReadObjects(string text)
        {
            var objects = new Dictionary<(int, int), PdfObjectInfo>();
            foreach (Match match in Regex.Matches(text, @"(?s)(\d+)\s+(\d+)\s+obj\s*(.*?)\s*endobj"))
            {
                int number = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                int generation = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                objects[(number, generation)] = new PdfObjectInfo(number, generation, match.Groups[3].Value.Trim());
            }
            return objects;
        }

        private static List<PageRef> ReadPageReferences(Dictionary<(int Number, int Generation), PdfObjectInfo> objects, string catalog)
        {
            var pagesRef = ReadReference(catalog, "Pages");
            var result = new List<PageRef>();
            VisitPageTree(pagesRef, objects, new HashSet<(int, int)>(), result);
            return result;
        }

        private static void VisitPageTree(PageRef node, Dictionary<(int Number, int Generation), PdfObjectInfo> objects, HashSet<(int, int)> visited, List<PageRef> pages)
        {
            if (!visited.Add((node.ObjectNumber, node.Generation)))
                return;
            if (!objects.TryGetValue((node.ObjectNumber, node.Generation), out var obj))
                return;
            if (Regex.IsMatch(obj.Content, @"/Type\s*/Page\b") && !Regex.IsMatch(obj.Content, @"/Type\s*/Pages\b"))
            {
                pages.Add(node);
                return;
            }
            foreach (var child in ReadReferencesFromArray(obj.Content, "Kids"))
                VisitPageTree(child, objects, visited, pages);
        }

        private static (double Width, double Height) ReadMediaBox(string pageContent)
        {
            var match = Regex.Match(pageContent, @"/MediaBox\s*\[\s*[-\d.]+\s+[-\d.]+\s+([-\d.]+)\s+([-\d.]+)\s*\]");
            if (!match.Success)
                return (595, 842);

            double width = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            double height = double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            return (width, height);
        }

        private static string ExtractDictionary(string text, int dictionaryStart)
        {
            if (dictionaryStart < 0 || dictionaryStart + 1 >= text.Length)
                throw new InvalidOperationException("PDF dictionary start was not found.");

            int depth = 0;
            for (int i = dictionaryStart; i < text.Length - 1; i++)
            {
                if (text[i] == '<' && text[i + 1] == '<')
                {
                    depth++;
                    i++;
                    continue;
                }
                if (text[i] == '>' && text[i + 1] == '>')
                {
                    depth--;
                    i++;
                    if (depth == 0)
                        return text.Substring(dictionaryStart, i - dictionaryStart + 1);
                }
            }
            throw new InvalidOperationException("PDF dictionary end was not found.");
        }

        private static PageRef ReadReference(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s+(\d+)\s+(\d+)\s+R");
            if (!match.Success)
                throw new InvalidOperationException($"PDF reference /{key} was not found.");
            return new PageRef(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
        }

        private static int ReadInteger(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s+(\d+)");
            if (!match.Success)
                throw new InvalidOperationException($"PDF integer /{key} was not found.");
            return int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        private static string? ReadRawArray(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s*(\[[^\]]+\])", RegexOptions.Singleline);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static List<PageRef> ReadReferencesFromArray(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s*\[(.*?)\]", RegexOptions.Singleline);
            if (!match.Success)
                return new List<PageRef>();
            return Regex.Matches(match.Groups[1].Value, @"(\d+)\s+(\d+)\s+R")
                .Select(m => new PageRef(int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture), int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)))
                .ToList();
        }

        private static List<List<(int Number, int Generation, long Offset)>> GroupConsecutive(List<(int Number, int Generation, long Offset)> objects)
        {
            var groups = new List<List<(int, int, long)>>();
            foreach (var obj in objects)
            {
                if (groups.Count == 0 || groups[^1][^1].Item1 + 1 != obj.Number)
                    groups.Add(new List<(int, int, long)>());
                groups[^1].Add(obj);
            }
            return groups;
        }

        private int FindBytes(byte[] src, byte[] pattern, int startIdx)
        {
            int max = src.Length - pattern.Length;
            for (int i = Math.Max(0, startIdx); i <= max; i++)
            {
                if (src[i] != pattern[0])
                    continue;
                bool match = true;
                for (int j = 1; j < pattern.Length; j++)
                {
                    if (src[i + j] != pattern[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match)
                    return i;
            }
            return -1;
        }

        private static string EncodePdfTextString(string value)
        {
            byte[] unicode = Encoding.BigEndianUnicode.GetBytes(value ?? string.Empty);
            var bytes = new byte[unicode.Length + 2];
            bytes[0] = 0xFE;
            bytes[1] = 0xFF;
            Buffer.BlockCopy(unicode, 0, bytes, 2, unicode.Length);
            return "<" + Convert.ToHexString(bytes) + ">";
        }

        private static string EscapeLiteral(string value)
        {
            return (value ?? string.Empty).Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        }

        private static string Format(double value)
        {
            return value.ToString("0.###", CultureInfo.InvariantCulture);
        }

        private static string ExtractCn(string dn)
        {
            var match = Regex.Match(dn ?? string.Empty, @"(?:^|,\s*)CN=([^,]+)");
            return match.Success ? match.Groups[1].Value : (dn ?? string.Empty);
        }

        private sealed record TrailerInfo(int Size, int RootObjectNumber, int RootGeneration, long PreviousXrefOffset, string? IdArray);
        private sealed record PdfObjectInfo(int Number, int Generation, string Content);
        private sealed record PageRef(int ObjectNumber, int Generation);
    }
}
