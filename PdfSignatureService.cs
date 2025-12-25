using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.Advanced; 
using PdfSharp.Drawing; 
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509.Store;

namespace MinsPDFViewer
{
    public class PdfSignatureService
    {
        private const int EstimatedSignatureSize = 8192; 
        private const char PlaceholderChar = 'A'; 

        public void SignPdf(string sourcePath, string destPath, SignatureConfig config)
        {
            if (config == null || config.PrivateKey == null || config.Certificate == null)
                throw new ArgumentException("서명 설정이 올바르지 않습니다.");

            using (var doc = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Modify))
            {
                var sigRect = new PdfRectangle(new XRect(0, 0, 0, 0));
                var page = doc.Pages[doc.PageCount - 1]; 

                if (config.UseVisualStamp)
                {
                    double margin = 50;
                    double stampWidth = 120;
                    double stampHeight = 60;
                    double drawX = page.Width.Point - stampWidth - margin;
                    double drawY = page.Height.Point - stampHeight - margin;

                    double pdfY = page.Height.Point - drawY - stampHeight;
                    sigRect = new PdfRectangle(new XRect(drawX, pdfY, stampWidth, stampHeight));

                    using (var gfx = XGraphics.FromPdfPage(page))
                    {
                        if (!string.IsNullOrEmpty(config.VisualStampPath) && File.Exists(config.VisualStampPath))
                        {
                            try
                            {
                                using (var xImage = XImage.FromFile(config.VisualStampPath))
                                {
                                    double ratio = (double)xImage.PixelWidth / xImage.PixelHeight;
                                    double imgW = stampWidth;
                                    double imgH = imgW / ratio;
                                    if (imgH > stampHeight) { imgH = stampHeight; imgW = imgH * ratio; }
                                    double offsetX = (stampWidth - imgW) / 2;
                                    double offsetY = (stampHeight - imgH) / 2;
                                    gfx.DrawImage(xImage, drawX + offsetX, drawY + offsetY, imgW, imgH);
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            var pen = new XPen(XColors.Red, 2);
                            var rect = new XRect(drawX, drawY, stampWidth, stampHeight);
                            gfx.DrawRectangle(XBrushes.White, rect);
                            gfx.DrawRectangle(pen, rect);

                            var fontTitle = new XFont("Malgun Gothic", 10, XFontStyleEx.Bold);
                            var fontBody = new XFont("Malgun Gothic", 8, XFontStyleEx.Regular);
                            var brush = XBrushes.Red;
                            double currentY = drawY + 15;
                            var format = new XStringFormat { Alignment = XStringAlignment.Center };
                            double centerX = drawX + (stampWidth / 2);

                            gfx.DrawString("[ 전자서명 완료 ]", fontTitle, brush, new XPoint(centerX, currentY), format);
                            currentY += 15;
                            string name = config.Certificate.Subject;
                            if (name.Length > 10) name = name.Substring(0, 10) + "...";
                            gfx.DrawString($"서명자 : {name}", fontBody, brush, new XPoint(centerX, currentY), format);
                            currentY += 15;
                            gfx.DrawString($"일자 : {DateTime.Now:yyyy-MM-dd}", fontBody, brush, new XPoint(centerX, currentY), format);
                        }
                    }
                }

                var sigDict = new PdfDictionary(doc);
                sigDict.Elements["/Type"] = new PdfName("/Sig");
                sigDict.Elements["/Filter"] = new PdfName("/Adobe.PPKLite");
                sigDict.Elements["/SubFilter"] = new PdfName("/adbe.pkcs7.detached");
                sigDict.Elements["/M"] = new PdfDate(DateTime.Now);
                sigDict.Elements["/Name"] = new PdfString(config.Certificate.Subject);
                
                if (!string.IsNullOrEmpty(config.Reason)) sigDict.Elements["/Reason"] = new PdfString(config.Reason);
                if (!string.IsNullOrEmpty(config.Location)) sigDict.Elements["/Location"] = new PdfString(config.Location);

                string placeholder = new string(PlaceholderChar, EstimatedSignatureSize * 2);
                sigDict.Elements["/Contents"] = new PdfString(placeholder);

                // ByteRange 공간 확보 (충분히 길게 0으로 채움)
                var byteRangePlaceholder = new PdfArray(doc);
                byteRangePlaceholder.Elements.Add(new PdfInteger(0));
                byteRangePlaceholder.Elements.Add(new PdfInteger(int.MaxValue));
                byteRangePlaceholder.Elements.Add(new PdfInteger(int.MaxValue));
                byteRangePlaceholder.Elements.Add(new PdfInteger(int.MaxValue));
                sigDict.Elements["/ByteRange"] = byteRangePlaceholder;

                var catalog = doc.Internals.Catalog;
                var acroForm = catalog.Elements.GetDictionary("/AcroForm");
                if (acroForm == null)
                {
                    acroForm = new PdfDictionary(doc);
                    catalog.Elements["/AcroForm"] = acroForm;
                }

                var sigField = new GenericPdfAnnotation(doc);
                sigField.Elements["/Type"] = new PdfName("/Annot");
                sigField.Elements["/Subtype"] = new PdfName("/Widget");
                sigField.Elements["/FT"] = new PdfName("/Sig");
                sigField.Elements["/Rect"] = sigRect;
                sigField.Elements["/T"] = new PdfString($"Signature{Guid.NewGuid().ToString().Substring(0, 8)}");
                sigField.Elements["/V"] = sigDict;
                sigField.Elements["/P"] = page;
                sigField.Elements["/F"] = new PdfInteger(132); 

                page.Annotations.Add(sigField);
                
                if (acroForm.Elements["/Fields"] is PdfArray fields) fields.Elements.Add(sigField.Reference!);
                else { var newFields = new PdfArray(doc); newFields.Elements.Add(sigField.Reference!); acroForm.Elements["/Fields"] = newFields; }

                using (var ms = new MemoryStream())
                {
                    doc.Save(ms);
                    byte[] pdfBytes = ms.ToArray();
                    byte[] signedPdf = SignWithBouncyCastle(pdfBytes, config);
                    File.WriteAllBytes(destPath, signedPdf);
                }
            }
        }

        private byte[] SignWithBouncyCastle(byte[] pdfBytes, SignatureConfig config)
        {
            // 1. Placeholder 위치 찾기
            byte[] searchPattern = Encoding.ASCII.GetBytes(new string(PlaceholderChar, 50));
            int patternPos = FindBytes(pdfBytes, searchPattern, 0);
            if (patternPos == -1) throw new Exception("서명 공간(Placeholder)을 찾을 수 없습니다.");

            // 2. Contents 괄호 범위 찾기
            int startOffset = -1;
            for (int i = patternPos; i >= Math.Max(0, patternPos - 20); i--)
            {
                if (pdfBytes[i] == '(' || pdfBytes[i] == '<') { startOffset = i + 1; break; }
            }
            if (startOffset == -1) throw new Exception("서명 Placeholder의 시작을 찾을 수 없습니다.");

            int endOffset = -1;
            int searchEndStart = patternPos + searchPattern.Length;
            for (int i = searchEndStart; i < pdfBytes.Length; i++)
            {
                if (pdfBytes[i] == ')' || pdfBytes[i] == '>') { endOffset = i; break; }
            }
            if (endOffset == -1) throw new Exception("서명 Placeholder의 끝을 찾을 수 없습니다.");
            
            int availableSpace = endOffset - startOffset;

            // 3. ByteRange 위치 찾기
            byte[] byteRangeKey = Encoding.ASCII.GetBytes("/ByteRange");
            int byteRangePos = LastIndexOfBytes(pdfBytes, byteRangeKey, startOffset);
            
            if (byteRangePos == -1 || (startOffset - byteRangePos > 2000)) 
            {
                int forwardSearch = FindBytes(pdfBytes, byteRangeKey, endOffset);
                if (forwardSearch != -1 && (forwardSearch - endOffset < 2000)) byteRangePos = forwardSearch;
            }
            if (byteRangePos == -1) throw new Exception("/ByteRange를 찾을 수 없습니다.");

            int byteRangeStart = -1;
            for(int i = byteRangePos; i < byteRangePos + 200; i++) { if(pdfBytes[i] == '[') { byteRangeStart = i; break; } }
            int byteRangeEnd = -1;
            for(int i = byteRangeStart; i < byteRangeStart + 200; i++) { if(pdfBytes[i] == ']') { byteRangeEnd = i; break; } }

            // 4. [핵심 수정] ByteRange 값을 먼저 계산하고, 원본 pdfBytes에 '미리' 덮어씌움
            //    그래야 이 값을 포함해서 해시를 계산할 때, 최종 파일 내용과 일치하게 됨.
            int fileLen = pdfBytes.Length;
            int excludeStart = startOffset - 1; 
            int excludeEnd = endOffset + 1;    

            string newByteRangeStr = $"[0 {excludeStart} {excludeEnd} {fileLen - excludeEnd}]";
            byte[] newByteRangeBytes = Encoding.ASCII.GetBytes(newByteRangeStr);
            int byteRangeAvailable = byteRangeEnd - byteRangeStart + 1;

            if (newByteRangeBytes.Length > byteRangeAvailable)
                throw new Exception($"ByteRange 예약 공간 부족 (필요: {newByteRangeBytes.Length}, 확보: {byteRangeAvailable})");

            // (1) pdfBytes에 ByteRange 업데이트 (패치)
            Array.Copy(newByteRangeBytes, 0, pdfBytes, byteRangeStart, newByteRangeBytes.Length);
            // 남는 공간은 공백(0x20)으로 채움
            for(int k = byteRangeStart + newByteRangeBytes.Length; k <= byteRangeEnd; k++) 
            {
                pdfBytes[k] = 32; 
            }

            // 5. 서명 데이터 생성 (이제 pdfBytes는 올바른 ByteRange 값을 포함하고 있음)
            //    서명 제외 범위: Contents의 괄호 포함 영역
            byte[] range1 = new byte[excludeStart];
            Array.Copy(pdfBytes, 0, range1, 0, excludeStart);
            
            byte[] range2 = new byte[fileLen - excludeEnd];
            Array.Copy(pdfBytes, excludeEnd, range2, 0, range2.Length);

            byte[] dataToSign = new byte[range1.Length + range2.Length];
            range1.CopyTo(dataToSign, 0);
            range2.CopyTo(dataToSign, range1.Length);

            byte[] cmsSignature = GenerateCmsSignature(dataToSign, config);
            string hexSig = BitConverter.ToString(cmsSignature).Replace("-", "");
            byte[] sigBytes = Encoding.ASCII.GetBytes(hexSig);

            if (sigBytes.Length > availableSpace) throw new Exception("서명 공간 부족");

            // 6. 서명 값 주입 (Contents 업데이트)
            //    참고: pdfBytes는 이미 위에서 ByteRange가 수정된 상태이므로 그대로 사용
            byte[] finalPdf = new byte[fileLen];
            Array.Copy(pdfBytes, finalPdf, fileLen); // 복사해서 최종본 생성

            finalPdf[excludeStart] = (byte)'<'; 
            Array.Copy(sigBytes, 0, finalPdf, startOffset, sigBytes.Length);
            for (int k = startOffset + sigBytes.Length; k < endOffset; k++) finalPdf[k] = 0;
            finalPdf[endOffset] = (byte)'>';

            // (ByteRange는 이미 4번 단계에서 올바르게 수정되어 있음)

            return finalPdf;
        }

        private int LastIndexOfBytes(byte[] src, byte[] pattern, int startIndex)
        {
            if (startIndex >= src.Length) startIndex = src.Length - 1;
            for (int i = startIndex; i >= 0; i--)
            {
                if (src[i] != pattern[0]) continue;
                bool match = true;
                for (int j = 1; j < pattern.Length; j++) { if (i + j >= src.Length || src[i + j] != pattern[j]) { match = false; break; } }
                if (match) return i;
            }
            return -1;
        }

        private int FindBytes(byte[] src, byte[] pattern, int startIdx)
        {
            int max = src.Length - pattern.Length;
            for (int i = startIdx; i <= max; i++)
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
            var signedData = gen.Generate(msg, false); 
            return signedData.GetEncoded();
        }
    }
}