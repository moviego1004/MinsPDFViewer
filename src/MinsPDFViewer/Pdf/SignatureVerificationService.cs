using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509;

namespace MinsPDFViewer
{
    public class SignatureVerificationService
    {
        public SignatureValidationResult? VerifySignatureAtPoint(string filePath, int pageIndex, double pdfX, double pdfY)
        {
            var signatures = ReadSignatureFields(filePath);
            var match = signatures.FirstOrDefault(s =>
                s.PageIndex == pageIndex &&
                pdfX >= s.Rect.X &&
                pdfX <= s.Rect.Right &&
                pdfY >= s.Rect.Y &&
                pdfY <= s.Rect.Top);

            return match == null ? null : VerifySignature(filePath, match);
        }

        public SignatureValidationResult? VerifySignatureForAnnotation(
            string filePath,
            int pageIndex,
            string? fieldName,
            SignaturePdfRect annotationRect)
        {
            var signatures = ReadSignatureFields(filePath);
            var match = signatures.FirstOrDefault(s =>
                s.PageIndex == pageIndex &&
                !string.IsNullOrWhiteSpace(fieldName) &&
                s.FieldName == fieldName);

            match ??= signatures.FirstOrDefault(s =>
                s.PageIndex == pageIndex &&
                annotationRect.CenterX >= s.Rect.X &&
                annotationRect.CenterX <= s.Rect.Right &&
                annotationRect.CenterY >= s.Rect.Y &&
                annotationRect.CenterY <= s.Rect.Top);

            return match == null ? null : VerifySignature(filePath, match);
        }

        private SignatureValidationResult VerifySignature(string filePath, SignatureFieldInfo field)
        {
            var result = new SignatureValidationResult();
            try
            {
                using var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                long[] range = field.ByteRange;
                if (range.Length != 4 || range[0] != 0 || range[1] < 0 || range[2] < range[1] || range[3] < 0 || range[2] + range[3] > file.Length)
                    throw new Exception("서명 범위(/ByteRange)가 잘못되었습니다.");

                // 서명이 빠진 구간은 정확히 /Contents <...> 여야 한다 (다른 내용을 서명 밖으로 빼는 위조 방지)
                file.Position = range[1];
                int open = file.ReadByte();
                file.Position = range[2] - 1;
                int close = file.ReadByte();
                if (open != '<' || close != '>')
                    throw new Exception("서명 범위가 서명 값(/Contents)과 맞지 않습니다.");

                result.CoversWholeDocument = range[2] + range[3] == file.Length;

                var processable = new PdfSignatureService.FileRangeProcessable(file, range);
                var cmsMsg = new CmsSignedData(processable, field.Contents);
                var signerStore = cmsMsg.GetSignerInfos();
                var signers = signerStore.GetSigners();

                if (signers.Count == 0)
                    throw new Exception("서명자 정보가 없습니다.");

                foreach (SignerInformation signer in signers)
                {
                    var certStore = cmsMsg.GetCertificates("Collection");
                    var matches = certStore.GetMatches(signer.SignerID);
                    if (matches.Count == 0)
                        throw new Exception("서명자 인증서를 찾을 수 없습니다.");

                    X509Certificate cert = matches.Cast<X509Certificate>().First();
                    if (signer.Verify(cert))
                    {
                        result.IsValid = true;
                        result.IsDocumentModified = !result.CoversWholeDocument;
                        result.Message = result.CoversWholeDocument
                            ? "서명이 유효하며, 문서가 수정되지 않았습니다."
                            : "서명한 판은 그대로입니다. 서명 이후 덧붙여 저장된 변경(주석·다른 서명 등)이 있습니다.";
                    }
                    else
                    {
                        result.IsValid = false;
                        result.IsDocumentModified = true;
                        result.Message = "문서 데이터와 서명 해시값이 일치하지 않습니다. (변조됨)";
                    }

                    result.SignerName = ParseCnFromDn(cert.SubjectDN.ToString());

                    if (signer.SignedAttributes != null)
                    {
                        var signingTimeAttr = signer.SignedAttributes[PkcsObjectIdentifiers.Pkcs9AtSigningTime];
                        if (signingTimeAttr != null && signingTimeAttr.AttrValues.Count > 0)
                        {
                            try
                            {
                                string? timeStr = signingTimeAttr.AttrValues[0]?.ToString();
                                if (!string.IsNullOrEmpty(timeStr))
                                {
                                    result.SigningTime = DateTime.ParseExact(
                                        timeStr,
                                        new[] { "yyMMddHHmmss'Z'", "yyyyMMddHHmmss'Z'" },
                                        null,
                                        System.Globalization.DateTimeStyles.AssumeUniversal);
                                }
                            }
                            catch { }
                        }
                    }
                }

                if (result.SigningTime == default)
                    result.SigningTime = field.SigningTime ?? default;
                result.Reason = field.Reason ?? string.Empty;
                result.Location = field.Location ?? string.Empty;
                if (string.IsNullOrEmpty(result.SignerName))
                    result.SignerName = "서명자 정보 없음";
            }
            catch (CmsException ex) when (ex.Message.Contains("digest", StringComparison.OrdinalIgnoreCase))
            {
                // 서명 범위의 바이트가 서명할 때와 다름 = 변조
                result.IsValid = false;
                result.IsDocumentModified = true;
                result.Message = "문서 데이터와 서명 해시값이 일치하지 않습니다. (변조됨)";
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.Message = $"검증 오류: {ex.Message}";
            }

            return result;
        }

        /// <summary>서명 필드(위젯 + 서명 사전). xref 스트림·객체 스트림 문서도 읽는다.</summary>
        private static List<SignatureFieldInfo> ReadSignatureFields(string filePath)
        {
            var result = new List<SignatureFieldInfo>();
            using var store = PdfObjectStore.Open(filePath);

            var pages = store.GetPageRefs();
            var pageIndexByObject = new Dictionary<(int, int), int>();
            var pageByWidget = new Dictionary<(int, int), int>();
            var directWidgets = new List<(string Dict, int PageIndex)>();
            for (int i = 0; i < pages.Count; i++)
            {
                pageIndexByObject[pages[i]] = i;
                if (!store.TryGetValue(pages[i], out var pageObj))
                    continue;
                string? annots = store.ResolveValue(PdfSyntax.GetValue(PdfSyntax.ExtractLeadingDictionary(pageObj.Content), "Annots"));
                if (annots == null || !annots.StartsWith("[", StringComparison.Ordinal))
                    continue;
                foreach (var item in PdfSyntax.ParseArrayItems(annots))
                {
                    if (PdfSyntax.TryParseReference(item, out int n, out int g))
                        pageByWidget[(n, g)] = i;
                    else if (item.TrimStart().StartsWith("<<", StringComparison.Ordinal))
                        directWidgets.Add((item.Trim(), i)); // /Annots 안에 직접 쓴 주석 사전
                }
            }

            var candidates = new List<(string Dict, int Number, int PageIndex)>();
            foreach (int number in store.ObjectNumbers)
            {
                var obj = store.GetObject(number, 0);
                if (obj == null || !obj.Content.StartsWith("<<", StringComparison.Ordinal))
                    continue;
                try
                {
                    candidates.Add((PdfSyntax.ExtractLeadingDictionary(obj.Content), number, pageByWidget.TryGetValue((number, 0), out int wp) ? wp : -1));
                }
                catch (FormatException)
                {
                }
            }
            candidates.AddRange(directWidgets.Select(d => (d.Dict, -1, d.PageIndex)));

            foreach (var (widget, number, annotsPage) in candidates)
            {
                if (PdfSyntax.GetValue(widget, "Subtype") != "/Widget")
                    continue;

                // /FT·/V·/T는 위젯에 없으면 부모 필드에서
                string? ft = InheritedFieldValue(store, widget, "FT");
                if (ft != "/Sig")
                    continue;
                // 서명 값은 보통 간접 객체지만 직접 사전으로 넣는 도구도 있다
                string? sig = store.ResolveValue(InheritedFieldValue(store, widget, "V"));
                if (sig == null || !sig.StartsWith("<<", StringComparison.Ordinal))
                    continue;
                sig = PdfSyntax.ExtractLeadingDictionary(sig);

                int pageIndex = -1;
                if (PdfSyntax.TryParseReference(PdfSyntax.GetValue(widget, "P"), out int pn, out int pg) &&
                    pageIndexByObject.TryGetValue((pn, pg), out int idx))
                    pageIndex = idx;
                else
                    pageIndex = annotsPage;

                if (!TryReadRect(store.ResolveValue(PdfSyntax.GetValue(widget, "Rect")) ?? "", out var rect) ||
                    !TryReadByteRange(store.ResolveValue(PdfSyntax.GetValue(sig, "ByteRange")) ?? "", out var byteRange) ||
                    !TryReadContents(PdfSyntax.GetValue(sig, "Contents") ?? "", out var contents))
                    continue;

                result.Add(new SignatureFieldInfo
                {
                    FieldName = DecodeTextString(InheritedFieldValue(store, widget, "T")),
                    PageIndex = pageIndex,
                    Rect = rect,
                    ByteRange = byteRange,
                    Contents = contents,
                    Reason = DecodeTextString(PdfSyntax.GetValue(sig, "Reason")),
                    Location = DecodeTextString(PdfSyntax.GetValue(sig, "Location")),
                    SigningTime = ReadPdfDate(DecodeTextString(PdfSyntax.GetValue(sig, "M")))
                });
            }

            return result;
        }

        private static string? InheritedFieldValue(PdfObjectStore store, string dict, string key)
        {
            var visited = new HashSet<(int, int)>();
            for (int depth = 0; depth < 32; depth++)
            {
                string? value = PdfSyntax.GetValue(dict, key);
                if (value != null)
                    return value;
                if (!PdfSyntax.TryParseReference(PdfSyntax.GetValue(dict, "Parent"), out int n, out int g) || !visited.Add((n, g)) ||
                    !store.TryGetValue((n, g), out var parent))
                    return null;
                dict = PdfSyntax.ExtractLeadingDictionary(parent.Content);
            }
            return null;
        }

        /// <summary>PDF 문자열 값 (리터럴 또는 16진, UTF-16BE BOM 지원).</summary>
        internal static string? DecodeTextString(string? value)
        {
            if (value == null)
                return null;
            value = value.Trim();
            byte[] bytes;
            if (value.StartsWith("<", StringComparison.Ordinal) && value.EndsWith(">", StringComparison.Ordinal))
            {
                string hex = new string(value.Substring(1, value.Length - 2).Where(Uri.IsHexDigit).ToArray());
                if (hex.Length % 2 == 1)
                    hex += "0";
                bytes = Convert.FromHexString(hex);
            }
            else if (value.StartsWith("(", StringComparison.Ordinal) && value.EndsWith(")", StringComparison.Ordinal))
            {
                bytes = UnescapeLiteralBytes(value.Substring(1, value.Length - 2));
            }
            else
            {
                return value;
            }

            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
            return Encoding.Latin1.GetString(bytes);
        }

        private static byte[] UnescapeLiteralBytes(string s)
        {
            var output = new List<byte>(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c != '\\' || i + 1 >= s.Length)
                {
                    output.Add((byte)c);
                    continue;
                }
                char n = s[++i];
                switch (n)
                {
                    case 'n': output.Add((byte)'\n'); break;
                    case 'r': output.Add((byte)'\r'); break;
                    case 't': output.Add((byte)'\t'); break;
                    case 'b': output.Add((byte)'\b'); break;
                    case 'f': output.Add((byte)'\f'); break;
                    case '\r':
                        if (i + 1 < s.Length && s[i + 1] == '\n') i++;
                        break;
                    case '\n': break;
                    default:
                        if (n >= '0' && n <= '7')
                        {
                            int code = n - '0';
                            for (int k = 0; k < 2 && i + 1 < s.Length && s[i + 1] >= '0' && s[i + 1] <= '7'; k++)
                                code = code * 8 + (s[++i] - '0');
                            output.Add((byte)code);
                        }
                        else
                        {
                            output.Add((byte)n);
                        }
                        break;
                }
            }
            return output.ToArray();
        }

        private static bool TryReadRect(string text, out SignaturePdfRect rect)
        {
            rect = default;
            var items = PdfSyntax.ParseArrayItems(text);
            if (items.Count != 4)
                return false;
            var n = new double[4];
            for (int i = 0; i < 4; i++)
            {
                if (!PdfSyntax.TryParseNumber(items[i], out n[i]))
                    return false;
            }
            double left = Math.Min(n[0], n[2]), right = Math.Max(n[0], n[2]);
            double bottom = Math.Min(n[1], n[3]), top = Math.Max(n[1], n[3]);
            rect = new SignaturePdfRect(left, bottom, right - left, top - bottom);
            return true;
        }

        private static bool TryReadByteRange(string text, out long[] byteRange)
        {
            byteRange = Array.Empty<long>();
            var items = PdfSyntax.ParseArrayItems(text);
            if (items.Count != 4)
                return false;
            var values = new long[4];
            for (int i = 0; i < 4; i++)
            {
                if (!long.TryParse(items[i].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out values[i]))
                    return false;
            }
            byteRange = values;
            return true;
        }

        private static bool TryReadContents(string text, out byte[] contents)
        {
            contents = Array.Empty<byte>();
            text = text.Trim();
            if (!text.StartsWith("<", StringComparison.Ordinal) || !text.EndsWith(">", StringComparison.Ordinal))
                return false;
            string hex = new string(text.Substring(1, text.Length - 2).Where(Uri.IsHexDigit).ToArray());
            if (hex.Length % 2 == 1)
                hex += "0";
            byte[] paddedContents = Convert.FromHexString(hex);
            contents = TrimPdfSignaturePadding(paddedContents);
            return true;
        }

        private static byte[] TrimPdfSignaturePadding(byte[] paddedContents)
        {
            if (paddedContents.Length == 0)
                return paddedContents;

            try
            {
                using var ms = new MemoryStream(paddedContents, writable: false);
                using var asn1 = new Asn1InputStream(ms);
                asn1.ReadObject();
                if (ms.Position > 0 && ms.Position <= paddedContents.Length)
                    return paddedContents.Take((int)ms.Position).ToArray();
            }
            catch
            {
                // Fall through to the conservative legacy trim below.
            }

            int length = paddedContents.Length;
            while (length > 2 &&
                   paddedContents[length - 1] == 0 &&
                   paddedContents[length - 2] == 0 &&
                   paddedContents[length - 3] == 0)
            {
                length--;
            }

            return paddedContents.Take(length).ToArray();
        }

        private static DateTime? ReadPdfDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;
            value = value.StartsWith("D:", StringComparison.Ordinal) ? value.Substring(2) : value;
            if (value.Length >= 14 &&
                DateTime.TryParseExact(value.Substring(0, 14), "yyyyMMddHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                return parsed;
            return null;
        }

        private string ParseCnFromDn(string dn)
        {
            try
            {
                var parts = dn.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    var kv = part.Trim().Split('=');
                    if (kv.Length == 2 && kv[0].Trim().ToUpper() == "CN")
                        return kv[1].Trim();
                }
                return dn;
            }
            catch { return dn; }
        }


        private sealed class SignatureFieldInfo
        {
            public string? FieldName { get; set; }
            public int PageIndex { get; set; }
            public SignaturePdfRect Rect { get; set; }
            public long[] ByteRange { get; set; } = Array.Empty<long>();
            public byte[] Contents { get; set; } = Array.Empty<byte>();
            public string? Reason { get; set; }
            public string? Location { get; set; }
            public DateTime? SigningTime { get; set; }
        }
    }
}
