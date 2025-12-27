using System;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.X509;

namespace MinsPDFViewer
{
    // 화면에 보여줄 인증서 정보 (UI용 모델)
    public class NpkiCertificateModel
    {
        public string Subject { get; set; } = "";
        public string Issuer { get; set; } = "";
        public DateTime NotBefore
        {
            get; set;
        }
        public DateTime NotAfter
        {
            get; set;
        }
        public string DirectoryPath { get; set; } = "";
        public bool IsExpired => DateTime.Now > NotAfter;

        public X509Certificate? Certificate
        {
            get; set;
        }

        public override string ToString() => $"{Subject} (만료: {NotAfter:yyyy-MM-dd})";
    }

    // 서명 생성 설정
    public class SignatureConfig
    {
        public required NpkiCertificateModel Certificate
        {
            get; set;
        }
        public required AsymmetricKeyParameter PrivateKey
        {
            get; set;
        }

        public string Reason { get; set; } = "I agree to the terms of this document.";
        public string Location { get; set; } = "Korea";
        public bool UseVisualStamp { get; set; } = true;
        public string? VisualStampPath
        {
            get; set;
        }
    }

    // [신규] 서명 검증 결과 모델
    public class SignatureValidationResult
    {
        public bool IsValid
        {
            get; set;
        }           // 무결성 검증 결과 (해시 일치 여부)
        public string SignerName { get; set; } = "";// 서명자 이름 (CN)
        public DateTime SigningTime
        {
            get; set;
        }   // 서명 시각
        public string Reason { get; set; } = "";    // 사유
        public string Location { get; set; } = "";  // 위치
        public string Message { get; set; } = "";   // 검증 메시지 (에러 내용 등)
        public bool IsDocumentModified
        {
            get; set;
        }// 문서 변조 여부
    }
}