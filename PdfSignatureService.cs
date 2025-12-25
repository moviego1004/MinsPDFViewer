using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.Advanced; 
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509.Store;

namespace MinsPDFViewer
{
    public class PdfSignatureService
    {
        // 서명 공간 크기 (충분히 크게 설정)
        private const int EstimatedSignatureSize = 8192; 

        public void SignPdf(string sourcePath, string destPath, SignatureConfig config)
        {
            if (config == null || config.PrivateKey == null || config.Certificate == null)
                throw new ArgumentException("서명 설정이 올바르지 않습니다.");

            // 1. PdfSharp으로 문서 구조 잡기 (서명 필드 예약)
            using (var doc = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Modify))
            {
                // 서명 딕셔너리 생성
                var sigDict = new PdfDictionary(doc);
                sigDict.Elements["/Type"] = new PdfName("/Sig");
                sigDict.Elements["/Filter"] = new PdfName("/Adobe.PPKLite");
                sigDict.Elements["/SubFilter"] = new PdfName("/adbe.pkcs7.detached");
                sigDict.Elements["/M"] = new PdfDate(DateTime.Now);
                sigDict.Elements["/Name"] = new PdfString(config.Certificate.Subject);
                
                if (!string.IsNullOrEmpty(config.Reason))
                    sigDict.Elements["/Reason"] = new PdfString(config.Reason);
                if (!string.IsNullOrEmpty(config.Location))
                    sigDict.Elements["/Location"] = new PdfString(config.Location);

                // [수정] Contents 예약: 문자열 "000..."으로 생성 
                // PDFSharp 6.x에서는 byte[] 생성자가 모호할 수 있으므로 문자열로 처리
                // 나중에 이 부분(괄호 포함)을 <Hex...>로 덮어씀
                string placeholder = new string('0', EstimatedSignatureSize * 2);
                sigDict.Elements["/Contents"] = new PdfString(placeholder);

                // ByteRange: [0 0 0 0] 형태로 예약
                var byteRangePlaceholder = new PdfArray(doc);
                byteRangePlaceholder.Elements.Add(new PdfInteger(0));
                byteRangePlaceholder.Elements.Add(new PdfInteger(0));
                byteRangePlaceholder.Elements.Add(new PdfInteger(0));
                byteRangePlaceholder.Elements.Add(new PdfInteger(0));
                sigDict.Elements["/ByteRange"] = byteRangePlaceholder;

                // [수정] AcroForm 필드 접근 (읽기 전용 속성 우회)
                var catalog = doc.Internals.Catalog;
                var acroForm = catalog.Elements.GetDictionary("/AcroForm");
                if (acroForm == null)
                {
                    acroForm = new PdfDictionary(doc);
                    catalog.Elements["/AcroForm"] = acroForm;
                }

                // [수정] GenericPdfAnnotation 사용 (타입 호환성 해결)
                var sigField = new GenericPdfAnnotation(doc);
                sigField.Elements["/Type"] = new PdfName("/Annot");
                sigField.Elements["/Subtype"] = new PdfName("/Widget");
                sigField.Elements["/FT"] = new PdfName("/Sig");
                sigField.Elements["/Rect"] = new PdfRectangle(new PdfSharp.Drawing.XRect(0, 0, 0, 0)); // 보이지 않음
                sigField.Elements["/T"] = new PdfString($"Signature{Guid.NewGuid().ToString().Substring(0, 8)}");
                sigField.Elements["/V"] = sigDict;
                sigField.Elements["/P"] = doc.Pages[0];
                sigField.Elements["/F"] = new PdfInteger(132); // Locked

                doc.Pages[0].Annotations.Add(sigField);
                
                // [수정] Fields 배열에 추가
                if (acroForm.Elements["/Fields"] is PdfArray fields)
                {
                    fields.Elements.Add(sigField.Reference);
                }
                else 
                { 
                    var newFields = new PdfArray(doc); 
                    newFields.Elements.Add(sigField.Reference); 
                    acroForm.Elements["/Fields"] = newFields; 
                }

                // 2. 임시 저장
                using (var ms = new MemoryStream())
                {
                    doc.Save(ms);
                    byte[] pdfBytes = ms.ToArray();

                    // 3. 바이너리 패치
                    byte[] signedPdf = SignWithBouncyCastle(pdfBytes, config, placeholder.Length);
                    File.WriteAllBytes(destPath, signedPdf);
                }
            }
        }

        private byte[] SignWithBouncyCastle(byte[] pdfBytes, SignatureConfig config, int contentLen)
        {
            // 1. /Contents (000...) 위치 찾기
            // PdfSharp이 문자열을 (...) 형태로 저장했을 가능성이 높음
            string searchKeyHex = "/Contents<";
            string searchKeyStr = "/Contents(";
            
            int contentStart = FindBytes(pdfBytes, Encoding.ASCII.GetBytes(searchKeyHex));
            bool isHex = true;

            if (contentStart == -1)
            {
                // 공백 포함 시도
                contentStart = FindBytes(pdfBytes, Encoding.ASCII.GetBytes("/Contents <"));
            }

            if (contentStart == -1)
            {
                // 문자열 형태 시도
                contentStart = FindBytes(pdfBytes, Encoding.ASCII.GetBytes(searchKeyStr));
                isHex = false;
                if (contentStart == -1) contentStart = FindBytes(pdfBytes, Encoding.ASCII.GetBytes("/Contents ("));
            }

            if (contentStart == -1) throw new Exception("PDF 서명 필드(/Contents)를 찾을 수 없습니다.");

            // 시작점: 괄호나 꺾쇠 다음 문자
            // 괄호 '(' 또는 '<'를 찾아야 함
            int delimiterPos = -1;
            for(int i = contentStart; i < contentStart + 20; i++)
            {
                if (pdfBytes[i] == '(' || pdfBytes[i] == '<')
                {
                    delimiterPos = i;
                    break;
                }
            }
            if (delimiterPos == -1) throw new Exception("Contents 시작 괄호를 찾을 수 없습니다.");

            int startOffset = delimiterPos + 1; 
            int endOffset = startOffset + contentLen;          

            // 2. /ByteRange [...] 위치 찾기
            int byteRangeKeyPos = FindBytes(pdfBytes, Encoding.ASCII.GetBytes("/ByteRange"));
            if (byteRangeKeyPos == -1) throw new Exception("PDF 서명 필드(/ByteRange)를 찾을 수 없습니다.");
            
            int byteRangeStart = -1;
            for(int i = byteRangeKeyPos; i < byteRangeKeyPos + 100; i++) 
            {
                if(pdfBytes[i] == '[') { byteRangeStart = i; break; }
            }
            if (byteRangeStart == -1) throw new Exception("ByteRange 배열 시작을 찾을 수 없습니다.");

            int byteRangeEnd = -1;
            for(int i = byteRangeStart; i < byteRangeStart + 100; i++)
            {
                if(pdfBytes[i] == ']') { byteRangeEnd = i; break; }
            }

            int fileLen = pdfBytes.Length;
            
            // 3. 서명할 데이터 범위 지정 (Contents 값 제외한 전체)
            // 주의: Contents의 괄호/꺾쇠까지 포함해서 제외해야 표준에 맞음 (또는 값만 제외)
            // 여기서는 '값' 부분만 제외하고 서명 (startOffset ~ endOffset)
            byte[] range1 = new byte[startOffset];
            Array.Copy(pdfBytes, 0, range1, 0, startOffset);
            
            byte[] range2 = new byte[fileLen - endOffset];
            Array.Copy(pdfBytes, endOffset, range2, 0, range2.Length);

            byte[] dataToSign = new byte[range1.Length + range2.Length];
            range1.CopyTo(dataToSign, 0);
            range2.CopyTo(dataToSign, range1.Length);

            // 4. 서명 생성
            byte[] cmsSignature = GenerateCmsSignature(dataToSign, config);
            string hexSig = BitConverter.ToString(cmsSignature).Replace("-", "");
            byte[] sigBytes = Encoding.ASCII.GetBytes(hexSig);

            // 공간 확인 (Hex 문자열이 예약된 공간보다 작거나 같아야 함)
            if (sigBytes.Length > contentLen) 
                throw new Exception($"서명 공간 부족 (할당: {contentLen}, 필요: {sigBytes.Length})");

            // 5. 결과 파일 작성
            byte[] finalPdf = new byte[fileLen];
            Array.Copy(pdfBytes, finalPdf, fileLen);

            // (1) 서명 덮어쓰기
            // 원본이 (...) 형태였어도 서명은 <...> 형태(Hex)가 표준이므로
            // 시작 괄호 '('를 '<'로, 끝 괄호 ')'를 '>'로 바꾸고 내용을 Hex로 채움
            finalPdf[delimiterPos] = (byte)'<'; 
            Array.Copy(sigBytes, 0, finalPdf, startOffset, sigBytes.Length);
            // 남는 공간은 0(null)으로 채움
            for (int k = startOffset + sigBytes.Length; k < endOffset; k++) finalPdf[k] = 0;
            finalPdf[endOffset] = (byte)'>';

            // (2) ByteRange 업데이트
            // [0, startOffset-1, endOffset+1, length-endOffset-1] (괄호 포함 제외 시)
            // 여기서는 단순하게 값 제외 기준으로 설정
            string newByteRangeStr = $"[0 {startOffset} {endOffset} {range2.Length}]";
            byte[] newByteRangeBytes = Encoding.ASCII.GetBytes(newByteRangeStr);
            
            if (newByteRangeBytes.Length <= (byteRangeEnd - byteRangeStart + 1))
            {
                Array.Copy(newByteRangeBytes, 0, finalPdf, byteRangeStart, newByteRangeBytes.Length);
                for(int k = byteRangeStart + newByteRangeBytes.Length; k <= byteRangeEnd; k++) finalPdf[k] = 32; // Space
            }

            return finalPdf;
        }

        private int FindBytes(byte[] src, byte[] pattern)
        {
            int max = src.Length - pattern.Length;
            for (int i = 0; i <= max; i++)
            {
                if (src[i] != pattern[0]) continue;
                bool match = true;
                for (int j = 1; j < pattern.Length; j++) { if (src[i + j] != pattern[j]) { match = false; break; } }
                if (match) return i;
            }
            return -1;
        }

        private byte[] GenerateCmsSignature(byte[] data, SignatureConfig config)
        {
            if (config.Certificate?.Certificate == null) throw new ArgumentException("인증서가 없습니다.");

            var certList = new List<X509Certificate> { config.Certificate.Certificate };
            var store = X509StoreFactory.Create("Certificate/Collection", new X509CollectionStoreParameters(certList));
            var gen = new CmsSignedDataGenerator();
            gen.AddSigner(config.PrivateKey, config.Certificate.Certificate, CmsSignedDataGenerator.DigestSha256);
            gen.AddCertificates(store);
            var msg = new CmsProcessableByteArray(data);
            var signedData = gen.Generate(msg, false); // Detached
            return signedData.GetEncoded();
        }
    }
}