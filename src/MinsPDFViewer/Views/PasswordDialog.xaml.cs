using System.Windows;
using System.Windows.Input;

namespace MinsPDFViewer
{
    public partial class PasswordDialog : Window
    {
        public string PasswordText { get; private set; } = string.Empty;

        public PasswordDialog(string? prompt = null, Window? owner = null)
        {
            InitializeComponent();
            if (owner != null)
                Owner = owner;

            if (!string.IsNullOrEmpty(prompt))
                TxtPrompt.Text = prompt;

            Loaded += (s, e) => PbPassword.Focus();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            PasswordText = PbPassword.Password;
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void PbPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnOk_Click(sender, e);
            }
        }
    }
}
