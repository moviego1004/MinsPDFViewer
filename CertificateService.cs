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
        // NPKI 기본 경로 (Windows 기준)
        private readonly string _npkiBasePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            @"AppData\LocalLow\NPKI"
        );

        public List<NpkiCertificateModel> LoadUserCertificates()
        {
            var list = new List<NpkiCertificateModel>();

            if (!Directory.Exists(_npkiBasePath)) return list;

            // NPKI 하위의 모든 폴더 검색 (예: yessign, crosscert 등)
            // 보통 구조: NPKI -> 기관폴더 -> USER -> 개별인증서폴더
            var certFiles = Directory.GetFiles(_npkiBasePath, "signCert.der", SearchOption.AllDirectories);

            foreach (var certPath in certFiles)
            {
                try
                {
                    var dir = Path.GetDirectoryName(certPath);
                    var keyPath = Path.Combine(dir!, "signPri.key");

                    // 개인키 파일이 없으면 무효
                    if (!File.Exists(keyPath)) continue;

                    // DER 인증서 파싱
                    var parser = new X509CertificateParser();
                    X509Certificate cert;
                    using (var stream = File.OpenRead(certPath))
                    {
                        cert = parser.ReadCertificate(stream);
                    }

                    // 모델 생성
                    list.Add(new NpkiCertificateModel
                    {
                        Subject = ParseCn(cert.SubjectDN.ToString()),
                        Issuer = ParseCn(cert.IssuerDN.ToString()),
                        NotBefore = cert.NotBefore,
                        NotAfter = cert.NotAfter,
                        DirectoryPath = dir!,
                        Certificate = cert
                    });
                }
                catch (Exception ex)
                {
                    // 손상된 인증서는 무시하고 계속 진행
                    System.Diagnostics.Debug.WriteLine($"인증서 로드 실패 ({certPath}): {ex.Message}");
                }
            }

            return list;
        }

        // 비밀번호로 암호화된 개인키(signPri.key) 추출
        public AsymmetricKeyParameter GetPrivateKey(NpkiCertificateModel model, string password)
        {
            var keyPath = Path.Combine(model.DirectoryPath, "signPri.key");
            if (!File.Exists(keyPath)) throw new FileNotFoundException("개인키 파일을 찾을 수 없습니다.");

            byte[] keyBytes = File.ReadAllBytes(keyPath);

            try
            {
                // BouncyCastle을 사용하여 암호화된 키 파싱
                // NPKI 키는 보통 PKCS#8 포맷이며, SEED 등의 알고리즘이 사용될 수 있음.
                // Portable.BouncyCastle은 일반적인 알고리즘은 지원하지만, 
                // 특정 한국형 알고리즘(SEED-CBC 등)이 걸려있으면 예외가 날 수 있음.
                // 일단 표준 방식으로 시도.
                
                var privKey = PrivateKeyFactory.DecryptKey(password.ToCharArray(), keyBytes);
                return privKey;
            }
            catch (Exception ex)
            {
                throw new Exception($"비밀번호가 틀리거나 지원하지 않는 암호화 방식입니다.\n({ex.Message})");
            }
        }

        // CN=홍길동,OU=... 형태에서 이름만 추출하는 헬퍼
        private string ParseCn(string dn)
        {
            var parts = dn.Split(',');
            foreach (var part in parts)
            {
                var trim = part.Trim();
                if (trim.StartsWith("CN=", StringComparison.OrdinalIgnoreCase))
                {
                    return trim.Substring(3);
                }
            }
            return dn; // CN이 없으면 전체 반환
        }
    }
}