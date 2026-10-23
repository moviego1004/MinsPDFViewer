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
                byte[] fileBytes = File.ReadAllBytes(filePath);
                int totalLen = field.ByteRange[1] + field.ByteRange[3];
                byte[] signedContent = new byte[totalLen];
                Buffer.BlockCopy(fileBytes, field.ByteRange[0], signedContent, 0, field.ByteRange[1]);
                Buffer.BlockCopy(fileBytes, field.ByteRange[2], signedContent, field.ByteRange[1], field.ByteRange[3]);

                var processable = new CmsProcessableByteArray(signedContent);
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
                        result.IsDocumentModified = false;
                        result.Message = "서명이 유효하며, 문서가 수정되지 않았습니다.";
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
            catch (Exception ex)
            {
                result.IsValid = false;
                result.Message = $"검증 오류: {ex.Message}";
            }

            return result;
        }

        private static List<SignatureFieldInfo> ReadSignatureFields(string filePath)
        {
            string text = Encoding.Latin1.GetString(File.ReadAllBytes(filePath));
            var objects = ReadObjects(text);
            var pageIndexByObject = BuildPageIndexMap(objects);
            var result = new List<SignatureFieldInfo>();

            foreach (var field in objects.Values)
            {
                if (!Regex.IsMatch(field.Content, @"/Subtype\s*/Widget") ||
                    !Regex.IsMatch(field.Content, @"/FT\s*/Sig"))
                    continue;

                var valueRef = TryReadReference(field.Content, "V");
                if (valueRef == null || !objects.TryGetValue((valueRef.ObjectNumber, valueRef.Generation), out var sigObject))
                    continue;

                var pageRef = TryReadReference(field.Content, "P");
                int pageIndex = pageRef != null && pageIndexByObject.TryGetValue((pageRef.ObjectNumber, pageRef.Generation), out int idx)
                    ? idx
                    : -1;

                if (!TryReadRect(field.Content, out var rect) ||
                    !TryReadByteRange(sigObject.Content, out var byteRange) ||
                    !TryReadContents(sigObject.Content, out var contents))
                    continue;

                result.Add(new SignatureFieldInfo
                {
                    FieldName = ReadLiteral(field.Content, "T"),
                    PageIndex = pageIndex,
                    Rect = rect,
                    ByteRange = byteRange,
                    Contents = contents,
                    Reason = ReadTextString(sigObject.Content, "Reason"),
                    Location = ReadTextString(sigObject.Content, "Location"),
                    SigningTime = ReadPdfDate(sigObject.Content)
                });
            }

            return result;
        }

        private static Dictionary<(int Number, int Generation), int> BuildPageIndexMap(Dictionary<(int Number, int Generation), PdfObjectInfo> objects)
        {
            var map = new Dictionary<(int, int), int>();
            var catalog = objects.Values.LastOrDefault(o => Regex.IsMatch(o.Content, @"/Type\s*/Catalog"));
            if (catalog == null)
                return map;

            var pagesRef = TryReadReference(catalog.Content, "Pages");
            if (pagesRef == null)
                return map;

            var pages = new List<PageRef>();
            VisitPageTree(pagesRef, objects, new HashSet<(int, int)>(), pages);
            for (int i = 0; i < pages.Count; i++)
                map[(pages[i].ObjectNumber, pages[i].Generation)] = i;
            return map;
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

        private static PageRef? TryReadReference(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s+(\d+)\s+(\d+)\s+R");
            if (!match.Success)
                return null;
            return new PageRef(int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture), int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture));
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

        private static bool TryReadRect(string text, out SignaturePdfRect rect)
        {
            rect = default;
            var match = Regex.Match(text, @"/Rect\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]");
            if (!match.Success)
                return false;
            double left = double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
            double bottom = double.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
            double right = double.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture);
            double top = double.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture);
            rect = new SignaturePdfRect(left, bottom, right - left, top - bottom);
            return true;
        }

        private static bool TryReadByteRange(string text, out int[] byteRange)
        {
            byteRange = Array.Empty<int>();
            var match = Regex.Match(text, @"/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]");
            if (!match.Success)
                return false;
            byteRange = new[]
            {
                int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture),
                int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture),
                int.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture),
                int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture)
            };
            return true;
        }

        private static bool TryReadContents(string text, out byte[] contents)
        {
            contents = Array.Empty<byte>();
            var match = Regex.Match(text, @"/Contents\s*<([0-9A-Fa-f\s]+)>", RegexOptions.Singleline);
            if (!match.Success)
                return false;
            string hex = Regex.Replace(match.Groups[1].Value, @"\s+", "");
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

        private static string? ReadLiteral(string text, string key)
        {
            var match = Regex.Match(text, @$"/{Regex.Escape(key)}\s*\((.*?)\)", RegexOptions.Singleline);
            return match.Success ? UnescapeLiteral(match.Groups[1].Value) : null;
        }

        private static string? ReadTextString(string text, string key)
        {
            var hex = Regex.Match(text, @$"/{Regex.Escape(key)}\s*<([0-9A-Fa-f]+)>");
            if (hex.Success)
            {
                byte[] bytes = Convert.FromHexString(hex.Groups[1].Value);
                if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                    return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
                return Encoding.Latin1.GetString(bytes);
            }

            return ReadLiteral(text, key);
        }

        private static DateTime? ReadPdfDate(string text)
        {
            string? value = ReadLiteral(text, "M");
            if (string.IsNullOrWhiteSpace(value))
                return null;
            value = value.StartsWith("D:", StringComparison.Ordinal) ? value.Substring(2) : value;
            if (value.Length >= 14 &&
                DateTime.TryParseExact(value.Substring(0, 14), "yyyyMMddHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                return parsed;
            return null;
        }

        private static string UnescapeLiteral(string value)
        {
            return value.Replace("\\(", "(").Replace("\\)", ")").Replace("\\\\", "\\");
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

        private sealed record PdfObjectInfo(int Number, int Generation, string Content);
        private sealed record PageRef(int ObjectNumber, int Generation);

        private sealed class SignatureFieldInfo
        {
            public string? FieldName { get; set; }
            public int PageIndex { get; set; }
            public SignaturePdfRect Rect { get; set; }
            public int[] ByteRange { get; set; } = Array.Empty<int>();
            public byte[] Contents { get; set; } = Array.Empty<byte>();
            public string? Reason { get; set; }
            public string? Location { get; set; }
            public DateTime? SigningTime { get; set; }
        }
    }
}
