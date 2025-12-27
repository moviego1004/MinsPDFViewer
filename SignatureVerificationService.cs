using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Advanced;

namespace MinsPDFViewer
{
    public class SignatureVerificationService
    {
        public SignatureValidationResult VerifySignature(string filePath, PdfDictionary sigDict)
        {
            var result = new SignatureValidationResult();
            try
            {
                if (sigDict.Elements["/Contents"] is not PdfString contents)
                    throw new Exception("서명 데이터(/Contents)가 없습니다.");

                if (sigDict.Elements["/ByteRange"] is not PdfArray byteRangeArray)
                    throw new Exception("서명 범위(/ByteRange)가 없습니다.");

                byte[] sigBytes = contents.Value.Select(c => (byte)c).ToArray();

                int[] ranges = new int[4];
                for (int i = 0; i < 4; i++)
                    ranges[i] = ((PdfInteger)byteRangeArray.Elements[i]).Value;

                byte[] fileBytes = File.ReadAllBytes(filePath);

                int len1 = ranges[1];
                int len2 = ranges[3];
                int totalLen = len1 + len2;

                byte[] signedContent = new byte[totalLen];
                Array.Copy(fileBytes, ranges[0], signedContent, 0, len1);
                Array.Copy(fileBytes, ranges[2], signedContent, len1, len2);

                var processable = new CmsProcessableByteArray(signedContent);
                var cmsMsg = new CmsSignedData(processable, sigBytes);

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
                                var timeObj = (Org.BouncyCastle.Asn1.Asn1Object)signingTimeAttr.AttrValues[0];
                                // [수정] Null 안전 처리
                                string? timeStr = timeObj?.ToString();
                                if (!string.IsNullOrEmpty(timeStr))
                                {
                                    result.SigningTime = DateTime.ParseExact(timeStr,
                                        new[] { "yyMMddHHmmss'Z'", "yyyyMMddHHmmss'Z'" },
                                        null, System.Globalization.DateTimeStyles.AssumeUniversal);
                                }
                            }
                            catch { }
                        }
                    }
                }

                if (result.SigningTime == default)
                {
                    if (sigDict.Elements["/M"] is PdfDate date)
                        result.SigningTime = date.Value;
                }
                if (sigDict.Elements["/Reason"] is PdfString reason)
                    result.Reason = reason.Value;
                if (sigDict.Elements["/Location"] is PdfString location)
                    result.Location = location.Value;

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

        private string ParseCnFromDn(string dn)
        {
            try
            {
                var parts = dn.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var part in parts)
                {
                    var kv = part.Trim().Split('=');
                    if (kv.Length == 2 && kv[0].Trim().ToUpper() == "CN")
                    {
                        return kv[1].Trim();
                    }
                }
                return dn;
            }
            catch { return dn; }
        }
    }
}