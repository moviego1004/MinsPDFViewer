using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.OpenSsl;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;

namespace MinsPDFViewer
{
    public class CertificateService
    {
        private readonly string _npkiBasePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            @"AppData\LocalLow\NPKI"
        );

        public List<NpkiCertificateModel> LoadUserCertificates()
        {
            var list = new List<NpkiCertificateModel>();
            if (!Directory.Exists(_npkiBasePath)) return list;

            var certFiles = Directory.GetFiles(_npkiBasePath, "signCert.der", SearchOption.AllDirectories);
            foreach (var certPath in certFiles)
            {
                var model = GetCertificateFromPath(Path.GetDirectoryName(certPath)!);
                if (model != null) list.Add(model);
            }
            return list;
        }

        // [신규] 특정 경로에서 인증서 로드 (USB 등 찾기용)
        public NpkiCertificateModel? GetCertificateFromPath(string directoryPath)
        {
            try
            {
                var certPath = Path.Combine(directoryPath, "signCert.der");
                var keyPath = Path.Combine(directoryPath, "signPri.key");

                if (!File.Exists(certPath) || !File.Exists(keyPath)) return null;

                var parser = new X509CertificateParser();
                X509Certificate cert;
                using (var stream = File.OpenRead(certPath))
                {
                    cert = parser.ReadCertificate(stream);
                }

                return new NpkiCertificateModel
                {
                    Subject = ParseCn(cert.SubjectDN.ToString()),
                    Issuer = ParseCn(cert.IssuerDN.ToString()),
                    NotBefore = cert.NotBefore,
                    NotAfter = cert.NotAfter,
                    DirectoryPath = directoryPath,
                    Certificate = cert
                };
            }
            catch
            {
                return null;
            }
        }

        public AsymmetricKeyParameter GetPrivateKey(NpkiCertificateModel model, string password)
        {
            var keyPath = Path.Combine(model.DirectoryPath, "signPri.key");
            if (!File.Exists(keyPath)) throw new FileNotFoundException("개인키 파일을 찾을 수 없습니다.");

            byte[] keyBytes = File.ReadAllBytes(keyPath);
            try
            {
                // NPKI 개인키 복호화 (SEED 알고리즘 이슈가 있을 수 있으나 우선 표준 시도)
                return PrivateKeyFactory.DecryptKey(password.ToCharArray(), keyBytes);
            }
            catch (Exception ex)
            {
                throw new Exception($"암호가 일치하지 않거나 지원되지 않는 형식입니다.\n{ex.Message}");
            }
        }

        private string ParseCn(string dn)
        {
            var parts = dn.Split(',');
            foreach (var part in parts)
            {
                var trim = part.Trim();
                if (trim.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
                    return trim.Substring(3);
            }
            return dn;
        }
    }
}