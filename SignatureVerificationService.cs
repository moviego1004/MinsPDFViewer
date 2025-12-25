using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.Asn1.Pkcs; 
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
                // PdfSharp이 제공하는 값을 바이트 배열로 변환
                byte[] sigBytes = contents.Value.Select(c => (byte)c).ToArray();

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
                // [핵심 수정] Detached 서명이므로, 원본 데이터(processable)를 생성자에 같이 넘겨줘야 함!
                var processable = new CmsProcessableByteArray(signedContent);
                var cmsMsg = new CmsSignedData(processable, sigBytes); // <-- 서명과 원본을 연결

                var signerStore = cmsMsg.GetSignerInfos();
                var signers = signerStore.GetSigners();

                if (signers.Count == 0) throw new Exception("서명자 정보가 없습니다.");

                foreach (SignerInformation signer in signers)
                {
                    // (1) 서명자 인증서 찾기
                    var certStore = cmsMsg.GetCertificates("Collection");
                    var matches = certStore.GetMatches(signer.SignerID);
                    if (matches.Count == 0) throw new Exception("서명자 인증서를 찾을 수 없습니다.");
                    
                    X509Certificate cert = matches.Cast<X509Certificate>().First();

                    // (2) 서명 검증 수행
                    // 내부적으로 processable의 해시를 계산하여 서명 해시와 비교함
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

                    // (3) 정보 추출 (검증 성공/실패와 무관하게 정보는 표시)
                    result.SignerName = ParseCnFromDn(cert.SubjectDN.ToString());
                    
                    // 서명 시각 (SigningTime Attribute)
                    if (signer.SignedAttributes != null)
                    {
                        var signingTimeAttr = signer.SignedAttributes[PkcsObjectIdentifiers.Pkcs9AtSigningTime];
                        if (signingTimeAttr != null && signingTimeAttr.AttrValues.Count > 0)
                        {
                            // ASN.1 Time 파싱 (간략화)
                            try 
                            {
                                var timeObj = (Org.BouncyCastle.Asn1.Asn1Object)signingTimeAttr.AttrValues[0];
                                // Time 객체 처리 (UtcTime or GeneralizedTime)
                                // BouncyCastle 버전에 따라 다를 수 있으나 ToString()으로 시도
                                // 정확히 하려면 Org.BouncyCastle.Asn1.Cms.Time.GetInstance(timeObj).Date 사용
                                result.SigningTime = DateTime.ParseExact(timeObj.ToString(), 
                                    new[] { "yyMMddHHmmss'Z'", "yyyyMMddHHmmss'Z'" }, 
                                    null, System.Globalization.DateTimeStyles.AssumeUniversal);
                            }
                            catch 
                            {
                                // 파싱 실패 시 PDF 메타데이터 시간 사용
                            }
                        }
                    }
                }
                
                // 6. PDF 딕셔너리의 메타데이터 추출 (BouncyCastle에서 못 가져왔을 경우 대비)
                if (result.SigningTime == default)
                {
                    if (sigDict.Elements["/M"] is PdfDate date) result.SigningTime = date.Value;
                }
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
            // 예: "C=KR,O=yessign,CN=홍길동" 또는 "CN=홍길동,O=..."
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
                return dn; // CN을 못 찾으면 전체 DN 반환
            }
            catch { return dn; }
        }
    }
}