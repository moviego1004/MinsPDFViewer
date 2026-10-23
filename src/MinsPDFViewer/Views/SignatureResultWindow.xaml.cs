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
                if (result.CoversWholeDocument)
                {
                    TxtIntegrity.Text = "문서가 서명 이후 수정되지 않았습니다.";
                    TxtIntegrity.Foreground = Brushes.Green;
                }
                else
                {
                    TxtIntegrity.Text = "서명한 판은 그대로입니다. 서명 이후 추가된 변경(주석·다른 서명 등)이 있습니다.";
                    TxtIntegrity.Foreground = Brushes.DarkOrange;
                }
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