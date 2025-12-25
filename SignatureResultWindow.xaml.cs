using System.Windows;
using System.Windows.Media;

namespace MinsPDFViewer
{
    public partial class SignatureResultWindow : Window
    {
        public SignatureResultWindow(SignatureValidationResult result)
        {
            InitializeComponent();
            DisplayResult(result);
        }

        private void DisplayResult(SignatureValidationResult result)
        {
            TxtSigner.Text = result.SignerName;
            TxtDate.Text = result.SigningTime.ToString("yyyy-MM-dd HH:mm:ss");
            TxtLocation.Text = string.IsNullOrEmpty(result.Location) ? "(없음)" : result.Location;
            TxtReason.Text = string.IsNullOrEmpty(result.Reason) ? "(없음)" : result.Reason;

            if (result.IsValid)
            {
                TxtIcon.Text = "✅";
                TxtTitle.Text = "서명이 유효합니다";
                TxtTitle.Foreground = Brushes.Green;
                TxtIntegrity.Text = "문서가 서명 이후 수정되지 않았습니다.";
                TxtIntegrity.Foreground = Brushes.Green;
            }
            else
            {
                TxtIcon.Text = "❌";
                TxtTitle.Text = "서명이 유효하지 않습니다";
                TxtTitle.Foreground = Brushes.Red;
                TxtIntegrity.Text = result.Message; // 에러 메시지 표시
                TxtIntegrity.Foreground = Brushes.Red;
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}