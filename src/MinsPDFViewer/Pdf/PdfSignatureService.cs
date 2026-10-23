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

        /// <summary>
        /// 원본 뒤에 증분 갱신으로 서명 필드·서명 사전을 붙이고 서명한다 (원본 바이트는 그대로 → 이전 서명 유지).
        /// customRect: PDF 사용자 공간 좌표. 페이지가 /Rotate로 돌아가 있으면 도장 그림도 화면에서 바로 서 보이게 돌린다.
        /// 파일 전체를 메모리에 올리지 않고, 해시도 파일에서 바로 읽어 계산한다.
        /// </summary>
        public void SignPdf(string sourcePath, string destPath, SignatureConfig config, int pageIndex, SignaturePdfRect? customRect)
        {
            if (config == null || config.PrivateKey == null || config.Certificate == null)
                throw new ArgumentException("서명 설정이 올바르지 않습니다.");

            byte[] append;
            long sigObjectOffset;
            byte[] sigObjectBytes;
            using (var store = PdfObjectStore.Open(sourcePath))
            {
                if (store.IsEncrypted)
                    throw new InvalidOperationException("암호화된 PDF에는 서명할 수 없습니다. 암호를 해제한 사본에 서명하세요.");

                var pages = store.GetPageRefs();
                if (pages.Count == 0)
                    throw new InvalidOperationException("PDF page tree could not be parsed.");
                if (pageIndex < 0 || pageIndex >= pages.Count)
                    pageIndex = pages.Count - 1;

                var pageRef = pages[pageIndex];
                if (!store.TryGetValue(pageRef, out var pageObject))
                    throw new InvalidOperationException("PDF page object was not found.");
                string pageDict = PdfSyntax.ExtractLeadingDictionary(pageObject.Content);

                var mediaBox = ReadBox(store.ResolveValue(store.GetInheritedValue(pageDict, "MediaBox"))) ?? new[] { 0.0, 0, 595, 842 };
                int rotation = 0;
                if (PdfSyntax.TryParseNumber(store.ResolveValue(store.GetInheritedValue(pageDict, "Rotate")) ?? "0", out double rot))
                    rotation = (((int)Math.Round(rot / 90.0) * 90) % 360 + 360) % 360;

                var rect = customRect.HasValue && customRect.Value.Width > 0 && customRect.Value.Height > 0
                    ? customRect.Value
                    : new SignaturePdfRect(Math.Max(mediaBox[0], mediaBox[2] - 170), mediaBox[1] + 50, 120, 60);

                int nextObjectNumber = store.Size;
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

                // 화면에서 보이는 도장 크기 (페이지가 90°/270° 돌아가 있으면 가로·세로가 바뀜)
                bool quarter = rotation % 180 == 90;
                double viewWidth = quarter ? rect.Height : rect.Width;
                double viewHeight = quarter ? rect.Width : rect.Height;
                string matrix = rotation switch
                {
                    90 => "[0 1 -1 0 0 0]",
                    180 => "[-1 0 0 -1 0 0]",
                    270 => "[0 -1 1 0 0 0]",
                    _ => "[1 0 0 1 0 0]"
                };

                byte[] imageObject = BuildAppearanceImageObject(appearanceImageObjectNumber, viewWidth, viewHeight, config);
                string appearanceStream = $"q\n{Format(viewWidth)} 0 0 {Format(viewHeight)} 0 0 cm\n/Im1 Do\nQ";
                string appearanceDict =
                    $"<< /Type /XObject /Subtype /Form /FormType 1 /BBox [0 0 {Format(viewWidth)} {Format(viewHeight)}] " +
                    $"/Matrix {matrix} " +
                    $"/Resources << /XObject << /Im1 {appearanceImageObjectNumber} 0 R >> >> " +
                    $"/Length {Encoding.ASCII.GetByteCount(appearanceStream)} >>";

                string widget =
                    "<< /Type /Annot /Subtype /Widget /FT /Sig " +
                    $"/Rect [{Format(rect.X)} {Format(rect.Y)} {Format(rect.Right)} {Format(rect.Top)}] " +
                    $"/T ({EscapeLiteral(fieldName)}) /F 132 /P {pageRef.Number} {pageRef.Generation} R " +
                    $"/V {sigObjectNumber} 0 R /AP << /N {appearanceObjectNumber} 0 R >> >>";

                var objects = new List<(int Number, int Generation, byte[] Data)>();
                string catalog = RewriteCatalog(store, acroFormObjectNumber, widgetObjectNumber, out var acroFormObject);
                objects.Add((store.RootNumber, store.RootGeneration, PdfIncrementalWriter.WrapObject(store.RootNumber, store.RootGeneration, catalog)));
                string updatedPage = RewritePageAnnots(store, pageDict, widgetObjectNumber, out var annotsObject);
                objects.Add((pageRef.Number, pageRef.Generation, PdfIncrementalWriter.WrapObject(pageRef.Number, pageRef.Generation, updatedPage)));
                sigObjectBytes = PdfIncrementalWriter.WrapObject(sigObjectNumber, 0, signatureDictionary);
                objects.Add((sigObjectNumber, 0, sigObjectBytes));
                objects.Add((widgetObjectNumber, 0, PdfIncrementalWriter.WrapObject(widgetObjectNumber, 0, widget)));
                objects.Add((appearanceObjectNumber, 0, PdfIncrementalWriter.WrapStream(appearanceObjectNumber, 0, appearanceDict, Encoding.ASCII.GetBytes(appearanceStream))));
                objects.Add((appearanceImageObjectNumber, 0, imageObject));
                if (acroFormObject != null)
                    objects.Add((acroFormObject.Value.Number, acroFormObject.Value.Generation,
                        PdfIncrementalWriter.WrapObject(acroFormObject.Value.Number, acroFormObject.Value.Generation, acroFormObject.Value.Content)));
                if (annotsObject != null)
                    objects.Add((annotsObject.Value.Number, annotsObject.Value.Generation,
                        PdfIncrementalWriter.WrapObject(annotsObject.Value.Number, annotsObject.Value.Generation, annotsObject.Value.Content)));

                append = PdfIncrementalWriter.BuildAppend(store, objects, out var offsets);
                sigObjectOffset = offsets[sigObjectNumber];
            }

            if (!string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(destPath), StringComparison.OrdinalIgnoreCase))
                File.Copy(sourcePath, destPath, overwrite: true);
            PdfIncrementalWriter.AppendToFile(destPath, append);
            SignInPlace(destPath, sigObjectOffset, sigObjectBytes, config);
        }

        /// <summary>파일 안의 서명 사전 자리(/ByteRange, /Contents)를 채운다.</summary>
        private void SignInPlace(string path, long sigObjectOffset, byte[] sigObjectBytes, SignatureConfig config)
        {
            string sigText = Encoding.Latin1.GetString(sigObjectBytes);
            int byteRangeRel = sigText.IndexOf(ByteRangePlaceholder, StringComparison.Ordinal);
            int contentsRel = sigText.IndexOf("/Contents <", StringComparison.Ordinal);
            if (byteRangeRel < 0 || contentsRel < 0)
                throw new Exception("서명 공간(Placeholder)을 찾을 수 없습니다.");

            long contentsOpen = sigObjectOffset + contentsRel + "/Contents ".Length;
            long contentsClose = contentsOpen + 1 + EstimatedSignatureSize * 2;

            using var file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            if (ReadByteAt(file, contentsOpen) != '<' || ReadByteAt(file, contentsClose) != '>')
                throw new Exception("서명 Placeholder 범위를 찾을 수 없습니다.");

            long excludeEnd = contentsClose + 1;
            long[] range = { 0, contentsOpen, excludeEnd, file.Length - excludeEnd };
            string byteRange = string.Format(CultureInfo.InvariantCulture, "[{0,10} {1,10} {2,10} {3,10}]", range[0], range[1], range[2], range[3]);
            if (byteRange.Length != ByteRangePlaceholder.Length)
                throw new Exception("ByteRange 예약 길이가 맞지 않습니다.");
            file.Position = sigObjectOffset + byteRangeRel;
            file.Write(Encoding.ASCII.GetBytes(byteRange));

            byte[] cmsSignature = GenerateCmsSignature(new FileRangeProcessable(file, range), config);
            string hexSig = Convert.ToHexString(cmsSignature);
            if (hexSig.Length > EstimatedSignatureSize * 2)
                throw new Exception("서명 공간 부족");

            var hex = new byte[EstimatedSignatureSize * 2];
            Array.Fill(hex, (byte)'0');
            Encoding.ASCII.GetBytes(hexSig, 0, hexSig.Length, hex, 0);
            file.Position = contentsOpen + 1;
            file.Write(hex);
            file.Flush(true);
        }

        private static int ReadByteAt(FileStream file, long position)
        {
            file.Position = position;
            return file.ReadByte();
        }

        /// <summary>서명 대상(ByteRange 두 구간)을 파일에서 바로 흘려 보낸다 (메모리에 모으지 않음).</summary>
        internal sealed class FileRangeProcessable : CmsProcessable
        {
            private readonly Stream _file;
            private readonly long[] _range;

            public FileRangeProcessable(Stream file, long[] range)
            {
                _file = file;
                _range = range;
            }

            public void Write(Stream outStream)
            {
                var buffer = new byte[1 << 20];
                for (int i = 0; i + 1 < _range.Length; i += 2)
                {
                    _file.Position = _range[i];
                    long remaining = _range[i + 1];
                    while (remaining > 0)
                    {
                        int read = _file.Read(buffer, 0, (int)Math.Min(buffer.Length, remaining));
                        if (read <= 0)
                            throw new EndOfStreamException("서명 범위가 파일보다 깁니다.");
                        outStream.Write(buffer, 0, read);
                        remaining -= read;
                    }
                }
            }

            [Obsolete]
            public object GetContent() => throw new NotSupportedException();
        }

        private byte[] GenerateCmsSignature(CmsProcessable content, SignatureConfig config)
        {
            if (config.Certificate?.Certificate == null)
                throw new ArgumentException("인증서가 없습니다.");

            var certList = new List<X509Certificate> { config.Certificate.Certificate };
            var store = X509StoreFactory.Create("Certificate/Collection", new X509CollectionStoreParameters(certList));
            var gen = new CmsSignedDataGenerator();
            gen.AddSigner(config.PrivateKey, config.Certificate.Certificate, CmsSignedDataGenerator.DigestSha256);
            gen.AddCertificates(store);
            var signedData = gen.Generate(content, false);
            return signedData.GetEncoded();
        }

        /// <summary>도장 그림(RGB 24비트, 위→아래) 그리기. 기본은 GDI+, 테스트에서 바꿀 수 있다.</summary>
        internal static Func<int, int, SignatureConfig, byte[]> AppearanceRenderer { get; set; } = RenderSignatureAppearance;

        private static byte[] BuildAppearanceImageObject(int number, double viewWidth, double viewHeight, SignatureConfig config)
        {
            int width = Math.Max(24, (int)Math.Ceiling(viewWidth * 3));
            int height = Math.Max(12, (int)Math.Ceiling(viewHeight * 3));
            byte[] rgb = AppearanceRenderer(width, height, config);
            byte[] compressed = CompressZlib(rgb);
            return PdfIncrementalWriter.WrapStream(number, 0,
                $"<< /Type /XObject /Subtype /Image /Width {width} /Height {height} " +
                "/ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode " +
                $"/Length {compressed.Length} >>", compressed);
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

        /// <summary>카탈로그 /AcroForm에 서명 필드를 더한다. 간접 AcroForm이면 그 객체를, 없거나 직접 사전이면 새 객체로.</summary>
        private static string RewriteCatalog(PdfObjectStore store, int acroFormObjectNumber, int widgetObjectNumber,
            out (int Number, int Generation, string Content)? acroFormObject)
        {
            string catalog = store.RootDictionary;
            string? acroForm = PdfSyntax.GetValue(catalog, "AcroForm");
            if (PdfSyntax.TryParseReference(acroForm, out int num, out int gen) && store.TryGetValue((num, gen), out var existing))
            {
                string dict = PdfSyntax.ExtractLeadingDictionary(existing.Content);
                acroFormObject = (num, gen, MergeAcroFormFields(store, dict, widgetObjectNumber));
                return catalog;
            }

            string baseDict = acroForm != null && acroForm.StartsWith("<<", StringComparison.Ordinal) ? acroForm : "<< >>";
            acroFormObject = (acroFormObjectNumber, 0, MergeAcroFormFields(store, baseDict, widgetObjectNumber));
            return PdfSyntax.SetValue(catalog, "AcroForm", $"{acroFormObjectNumber} 0 R");
        }

        private static string MergeAcroFormFields(PdfObjectStore store, string acroForm, int widgetObjectNumber)
        {
            string widgetRef = $"{widgetObjectNumber} 0 R";
            string? fields = store.ResolveValue(PdfSyntax.GetValue(acroForm, "Fields"));
            var items = fields != null && fields.StartsWith("[", StringComparison.Ordinal) ? PdfSyntax.ParseArrayItems(fields) : new List<string>();
            items.Add(widgetRef);
            acroForm = PdfSyntax.SetValue(acroForm, "Fields", "[" + string.Join(" ", items.Select(i => i.Trim())) + "]");
            return PdfSyntax.SetValue(acroForm, "SigFlags", "3");
        }

        /// <summary>페이지 /Annots에 위젯을 더한다. 간접 배열이면 그 배열 객체를 고친다.</summary>
        private static string RewritePageAnnots(PdfObjectStore store, string pageDict, int widgetObjectNumber,
            out (int Number, int Generation, string Content)? annotsObject)
        {
            annotsObject = null;
            string widgetRef = $"{widgetObjectNumber} 0 R";
            string? annots = PdfSyntax.GetValue(pageDict, "Annots");
            if (PdfSyntax.TryParseReference(annots, out int num, out int gen) && store.TryGetValue((num, gen), out var arrayObject) &&
                arrayObject.Content.TrimStart().StartsWith("[", StringComparison.Ordinal))
            {
                var refItems = PdfSyntax.ParseArrayItems(arrayObject.Content.Trim());
                refItems.Add(widgetRef);
                annotsObject = (num, gen, "[" + string.Join(" ", refItems.Select(i => i.Trim())) + "]");
                return pageDict;
            }

            var items = annots != null && annots.StartsWith("[", StringComparison.Ordinal) ? PdfSyntax.ParseArrayItems(annots) : new List<string>();
            items.Add(widgetRef);
            return PdfSyntax.SetValue(pageDict, "Annots", "[" + string.Join(" ", items.Select(i => i.Trim())) + "]");
        }

        private static double[]? ReadBox(string? value)
        {
            if (value == null)
                return null;
            var items = PdfSyntax.ParseArrayItems(value);
            if (items.Count != 4)
                return null;
            var n = new double[4];
            for (int i = 0; i < 4; i++)
            {
                if (!PdfSyntax.TryParseNumber(items[i], out n[i]))
                    return null;
            }
            return new[] { Math.Min(n[0], n[2]), Math.Min(n[1], n[3]), Math.Max(n[0], n[2]), Math.Max(n[1], n[3]) };
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

    }
}
