using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using Org.BouncyCastle.X509.Store;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Advanced;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;

namespace MinsPDFViewer
{
    public class PdfSignatureService
    {
        private const int EstimatedSignatureSize = 8192;
        private const char PlaceholderChar = 'A';

        public void SignPdf(string sourcePath, string destPath, SignatureConfig config, int pageIndex, XRect? customRect)
        {
            if (config == null || config.PrivateKey == null || config.Certificate == null)
                throw new ArgumentException("서명 설정이 올바르지 않습니다.");

            using (var doc = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Modify))
            {
                if (pageIndex < 0 || pageIndex >= doc.PageCount)
                    pageIndex = doc.PageCount - 1;
                var page = doc.Pages[pageIndex];

                double stampWidth = 120;
                double stampHeight = 60;
                XRect targetRect;

                if (customRect.HasValue && customRect.Value.Width > 0 && customRect.Value.Height > 0)
                {
                    targetRect = customRect.Value;
                    stampWidth = targetRect.Width;
                    stampHeight = targetRect.Height;
                }
                else
                {
                    double margin = 50;
                    double drawX = page.Width.Point - stampWidth - margin;
                    double drawY = margin;
                    targetRect = new XRect(drawX, drawY, stampWidth, stampHeight);
                }

                var sigRect = new PdfRectangle(targetRect);

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

                string placeholder = new string(PlaceholderChar, EstimatedSignatureSize * 2);
                sigDict.Elements["/Contents"] = new PdfString(placeholder);

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

                // --- Appearance Stream 생성 ---
                if (config.UseVisualStamp)
                {
                    var form = new XForm(doc, new XRect(0, 0, stampWidth, stampHeight));
                    bool imageDrawn = false;

                    using (var gfx = XGraphics.FromForm(form))
                    {
                        // [수정] 배경 흰색 칠하기 제거 -> 투명 배경
                        // gfx.DrawRectangle(XBrushes.White, 0, 0, stampWidth, stampHeight);

                        // 1. 이미지 그리기 시도
                        if (!string.IsNullOrEmpty(config.VisualStampPath) && File.Exists(config.VisualStampPath))
                        {
                            try
                            {
                                using (var fs = new FileStream(config.VisualStampPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                using (var xImage = XImage.FromStream(fs))
                                {
                                    double ratio = (double)xImage.PixelWidth / xImage.PixelHeight;
                                    double imgW = stampWidth;
                                    double imgH = imgW / ratio;
                                    if (imgH > stampHeight)
                                    {
                                        imgH = stampHeight;
                                        imgW = imgH * ratio;
                                    }
                                    double offsetX = (stampWidth - imgW) / 2;
                                    double offsetY = (stampHeight - imgH) / 2;
                                    gfx.DrawImage(xImage, offsetX, offsetY, imgW, imgH);
                                    imageDrawn = true;
                                }
                            }
                            catch { imageDrawn = false; }
                        }

                        // 2. 텍스트 도장 그리기 (이미지 실패 시 혹은 텍스트 모드)
                        if (!imageDrawn)
                        {
                            var pen = new XPen(XColors.Red, 2);
                            // 테두리
                            gfx.DrawRectangle(pen, 1, 1, stampWidth - 2, stampHeight - 2);

                            var fontTitle = new XFont("Malgun Gothic", 10, XFontStyleEx.Bold);
                            var fontBody = new XFont("Malgun Gothic", 8, XFontStyleEx.Regular);
                            var brush = XBrushes.Red;

                            var format = new XStringFormat { Alignment = XStringAlignment.Center, LineAlignment = XLineAlignment.Center };

                            double centerY = stampHeight / 2;
                            double centerX = stampWidth / 2;
                            double lineHeight = 14;

                            gfx.DrawString("[ 전자서명 완료 ]", fontTitle, brush, new XPoint(centerX, centerY - lineHeight), format);

                            string name = config.Certificate.Subject;
                            if (name.Contains("CN="))
                            {
                                int start = name.IndexOf("CN=") + 3;
                                int end = name.IndexOf(",", start);
                                if (end == -1)
                                    end = name.Length;
                                name = name.Substring(start, end - start);
                            }
                            if (name.Length > 10)
                                name = name.Substring(0, 10) + "...";

                            gfx.DrawString($"서명자 : {name}", fontBody, brush, new XPoint(centerX, centerY), format);
                            gfx.DrawString($"일자 : {DateTime.Now:yyyy-MM-dd}", fontBody, brush, new XPoint(centerX, centerY + lineHeight), format);
                        }
                    }

                    var pdfFormProp = typeof(XForm).GetProperty("PdfForm", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                    if (pdfFormProp != null)
                    {
                        var obj = pdfFormProp.GetValue(form);
                        if (obj is PdfFormXObject pdfForm)
                        {
                            var apDict = new PdfDictionary(doc);
                            apDict.Elements["/N"] = pdfForm.Reference;
                            sigField.Elements["/AP"] = apDict;
                        }
                    }
                }

                page.Annotations.Add(sigField);

                if (acroForm.Elements["/Fields"] is PdfArray fields)
                    fields.Elements.Add(sigField.Reference!);
                else
                {
                    var newFields = new PdfArray(doc);
                    newFields.Elements.Add(sigField.Reference!);
                    acroForm.Elements["/Fields"] = newFields;
                }

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
            byte[] searchPattern = Encoding.ASCII.GetBytes(new string(PlaceholderChar, 50));
            int patternPos = FindBytes(pdfBytes, searchPattern, 0);
            if (patternPos == -1)
                throw new Exception("서명 공간(Placeholder)을 찾을 수 없습니다.");

            int startOffset = -1;
            for (int i = patternPos; i >= Math.Max(0, patternPos - 20); i--)
            {
                if (pdfBytes[i] == '(' || pdfBytes[i] == '<')
                {
                    startOffset = i + 1;
                    break;
                }
            }
            if (startOffset == -1)
                throw new Exception("서명 Placeholder의 시작을 찾을 수 없습니다.");

            int endOffset = -1;
            int searchEndStart = patternPos + searchPattern.Length;
            for (int i = searchEndStart; i < pdfBytes.Length; i++)
            {
                if (pdfBytes[i] == ')' || pdfBytes[i] == '>')
                {
                    endOffset = i;
                    break;
                }
            }
            if (endOffset == -1)
                throw new Exception("서명 Placeholder의 끝을 찾을 수 없습니다.");

            int availableSpace = endOffset - startOffset;

            byte[] byteRangeKey = Encoding.ASCII.GetBytes("/ByteRange");
            int byteRangePos = LastIndexOfBytes(pdfBytes, byteRangeKey, startOffset);

            if (byteRangePos == -1 || (startOffset - byteRangePos > 2000))
            {
                int forwardSearch = FindBytes(pdfBytes, byteRangeKey, endOffset);
                if (forwardSearch != -1 && (forwardSearch - endOffset < 2000))
                    byteRangePos = forwardSearch;
            }
            if (byteRangePos == -1)
                throw new Exception("/ByteRange를 찾을 수 없습니다.");

            int byteRangeStart = -1;
            for (int i = byteRangePos; i < byteRangePos + 200; i++)
            {
                if (pdfBytes[i] == '[')
                {
                    byteRangeStart = i;
                    break;
                }
            }
            int byteRangeEnd = -1;
            for (int i = byteRangeStart; i < byteRangeStart + 200; i++)
            {
                if (pdfBytes[i] == ']')
                {
                    byteRangeEnd = i;
                    break;
                }
            }

            int fileLen = pdfBytes.Length;
            int excludeStart = startOffset - 1;
            int excludeEnd = endOffset + 1;

            string newByteRangeStr = $"[0 {excludeStart} {excludeEnd} {fileLen - excludeEnd}]";
            byte[] newByteRangeBytes = Encoding.ASCII.GetBytes(newByteRangeStr);
            int byteRangeAvailable = byteRangeEnd - byteRangeStart + 1;

            if (newByteRangeBytes.Length > byteRangeAvailable)
                throw new Exception($"ByteRange 예약 공간 부족");

            Array.Copy(newByteRangeBytes, 0, pdfBytes, byteRangeStart, newByteRangeBytes.Length);
            for (int k = byteRangeStart + newByteRangeBytes.Length; k <= byteRangeEnd; k++)
                pdfBytes[k] = 32;

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

            if (sigBytes.Length > availableSpace)
                throw new Exception("서명 공간 부족");

            byte[] finalPdf = new byte[fileLen];
            Array.Copy(pdfBytes, finalPdf, fileLen);

            finalPdf[excludeStart] = (byte)'<';
            Array.Copy(sigBytes, 0, finalPdf, startOffset, sigBytes.Length);
            for (int k = startOffset + sigBytes.Length; k < endOffset; k++)
                finalPdf[k] = 0;
            finalPdf[endOffset] = (byte)'>';

            return finalPdf;
        }

        private int LastIndexOfBytes(byte[] src, byte[] pattern, int startIndex)
        {
            if (startIndex >= src.Length)
                startIndex = src.Length - 1;
            for (int i = startIndex; i >= 0; i--)
            {
                if (src[i] != pattern[0])
                    continue;
                bool match = true;
                for (int j = 1; j < pattern.Length; j++)
                {
                    if (i + j >= src.Length || src[i + j] != pattern[j])
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

        private int FindBytes(byte[] src, byte[] pattern, int startIdx)
        {
            int max = src.Length - pattern.Length;
            for (int i = startIdx; i <= max; i++)
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
    }
}