using System;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.X509;

namespace MinsPDFViewer
{
    // 화면에 보여줄 인증서 정보 (UI용 모델)
    public class NpkiCertificateModel
    {
        public string Subject { get; set; } = "";      // 사용자 이름 (예: 홍길동)
        public string Issuer { get; set; } = "";       // 발급 기관 (예: yessign)
        public DateTime NotBefore { get; set; }        // 유효기간 시작
        public DateTime NotAfter { get; set; }         // 유효기간 만료
        public string DirectoryPath { get; set; } = "";// 파일 경로 (signCert.der, signPri.key 위치)
        public bool IsExpired => DateTime.Now > NotAfter; // 만료 여부

        // 내부 로직용 BouncyCastle 객체
        public X509Certificate? Certificate { get; set; }
        
        public override string ToString() => $"{Subject} (만료: {NotAfter:yyyy-MM-dd})";
    }

    // 서명 생성 시 필요한 정보 묶음
    public class SignatureConfig
    {
        // [수정] 필수 값으로 지정하여 Null 경고 제거
        public required NpkiCertificateModel Certificate { get; set; }
        public required AsymmetricKeyParameter PrivateKey { get; set; } // 비밀번호로 푼 개인키
        
        public string Reason { get; set; } = "I agree to the terms of this document.";
        public string Location { get; set; } = "Korea";
        public bool UseVisualStamp { get; set; } = true;
        public string? VisualStampPath { get; set; } // 도장 이미지 경로 (선택)
    }
}