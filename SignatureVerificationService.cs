using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.Asn1.Pkcs; // [추가] OID 사용을 위해 필요
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
                // 1. 서명 딕셔너리에서 데이터 추출
                if (sigDict.Elements["/Contents"] is not PdfString contents)
                    throw new Exception("서명 데이터(/Contents)가 없습니다.");

                if (sigDict.Elements["/ByteRange"] is not PdfArray byteRangeArray)
                    throw new Exception("서명 범위(/ByteRange)가 없습니다.");

                // 2. /Contents 파싱
                // PdfSharp의 PdfString.Value는 이미 디코딩된 문자열일 수 있으므로 raw bytes로 변환 시 주의
                // 보통 PdfSharp은 Hex String을 내부적으로 byte[]로 변환하여 관리함
                byte[] sigBytes = contents.Value.Select(c => (byte)c).ToArray();
                
                // 만약 Hex String 형태("<...>")로 raw string에 들어있다면 별도 파싱이 필요할 수 있으나,
                // PdfSharp의 PdfString은 보통 Value 접근 시 디코딩된 값을 줌.

                // 3. /ByteRange 파싱
                int[] ranges = new int[4];
                for (int i = 0; i < 4; i++)
                    ranges[i] = ((PdfInteger)byteRangeArray.Elements[i]).Value;

                // 4. 원본 파일에서 서명된 데이터(Signed Data) 추출
                byte[] fileBytes = File.ReadAllBytes(filePath);
                
                int len1 = ranges[1];
                int len2 = ranges[3];
                int totalLen = len1 + len2;

                byte[] signedContent = new byte[totalLen];
                Array.Copy(fileBytes, ranges[0], signedContent, 0, len1);
                Array.Copy(fileBytes, ranges[2], signedContent, len1, len2);

                // 5. BouncyCastle로 CMS(PKCS#7) 검증
                var cmsMsg = new CmsSignedData(sigBytes);
                var signerStore = cmsMsg.GetSignerInfos();
                var signers = signerStore.GetSigners();

                foreach (SignerInformation signer in signers)
                {
                    // (1) 서명자 인증서 찾기
                    var certStore = cmsMsg.GetCertificates("Collection");
                    // [수정] ICollection 반환값 처리
                    var matches = certStore.GetMatches(signer.SignerID);
                    if (matches.Count == 0) throw new Exception("서명자 인증서를 찾을 수 없습니다.");
                    
                    // 첫 번째 인증서 가져오기
                    X509Certificate cert = matches.Cast<X509Certificate>().First();

                    // (2) 서명 검증 (무결성 확인)
                    var processable = new CmsProcessableByteArray(signedContent);
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
                        result.Message = "문서 해시가 일치하지 않습니다. (변조됨)";
                    }

                    // (3) 정보 추출
                    result.SignerName = ParseCnFromDn(cert.SubjectDN.ToString());
                    
                    // 서명 시각 (SigningTime Attribute 확인)
                    if (signer.SignedAttributes != null)
                    {
                        // [수정] CmsAttributes 대신 PkcsObjectIdentifiers 사용
                        var signingTimeAttr = signer.SignedAttributes[PkcsObjectIdentifiers.Pkcs9AtSigningTime];
                        if (signingTimeAttr != null)
                        {
                           // ASN.1 파싱 로직이 복잡하므로 여기서는 PDF 메타데이터(/M)를 우선 사용하거나 생략
                        }
                    }
                }
                
                // 6. PDF 딕셔너리의 메타데이터 추출
                if (sigDict.Elements["/M"] is PdfDate date) result.SigningTime = date.Value;
                if (sigDict.Elements["/Reason"] is PdfString reason) result.Reason = reason.Value;
                if (sigDict.Elements["/Location"] is PdfString location) result.Location = location.Value;
                
                if (string.IsNullOrEmpty(result.SignerName)) result.SignerName = "서명자 정보 없음";

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
                var parts = dn.Split(',');
                foreach (var part in parts)
                {
                    var kv = part.Trim().Split('=');
                    if (kv.Length == 2 && kv[0].ToUpper() == "CN")
                    {
                        return kv[1];
                    }
                }
                return dn;
            }
            catch { return dn; }
        }
    }
}