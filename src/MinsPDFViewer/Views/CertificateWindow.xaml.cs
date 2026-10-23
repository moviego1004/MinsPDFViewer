using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace MinsPDFViewer
{
    public partial class CertificateWindow : Window
    {
        private readonly CertificateService _certService;
        public SignatureConfig? ResultConfig { get; private set; } // 결과 반환용

        public CertificateWindow()
        {
            InitializeComponent();
            _certService = new CertificateService();
            LoadCertificates();
        }

        private void LoadCertificates()
        {
            try
            {
                var certs = _certService.LoadUserCertificates();
                LbCertificates.ItemsSource = certs;
                if (certs.Count > 0) LbCertificates.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"인증서 로드 실패: {ex.Message}");
            }
        }

        // USB 등 외부 경로 찾기
        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "인증서 파일 (signCert.der)|signCert.der|모든 파일 (*.*)|*.*",
                Title = "인증서 파일(signCert.der)을 선택하세요"
            };

            if (dlg.ShowDialog() == true)
            {
                string dirPath = System.IO.Path.GetDirectoryName(dlg.FileName)!;
                var model = _certService.GetCertificateFromPath(dirPath);

                if (model != null)
                {
                    // 목록에 없으면 추가
                    var list = LbCertificates.ItemsSource as List<NpkiCertificateModel> ?? new List<NpkiCertificateModel>();
                    if (!list.Exists(c => c.DirectoryPath == model.DirectoryPath))
                    {
                        list.Add(model);
                        LbCertificates.ItemsSource = null; // 갱신 트리거
                        LbCertificates.ItemsSource = list;
                    }
                    
                    // 해당 아이템 선택
                    LbCertificates.SelectedItem = model;
                    foreach (var item in LbCertificates.Items)
                    {
                        if ((item as NpkiCertificateModel)?.DirectoryPath == model.DirectoryPath)
                        {
                            LbCertificates.SelectedItem = item;
                            break;
                        }
                    }
                    TxtPassword.Focus();
                }
                else
                {
                    MessageBox.Show("올바른 NPKI 인증서 폴더가 아닙니다.\n(signCert.der와 signPri.key가 모두 있어야 합니다.)");
                }
            }
        }

        private void LbCertificates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            BtnConfirm.IsEnabled = LbCertificates.SelectedItem != null;
        }

        private void BtnConfirm_Click(object sender, RoutedEventArgs e)
        {
            ProcessLogin();
        }

        private void TxtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) ProcessLogin();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ProcessLogin()
        {
            var selectedCert = LbCertificates.SelectedItem as NpkiCertificateModel;
            if (selectedCert == null) return;

            string password = TxtPassword.Password;
            if (string.IsNullOrEmpty(password))
            {
                MessageBox.Show("비밀번호를 입력하세요.");
                return;
            }

            try
            {
                // 비밀번호 검증 및 개인키 추출 시도
                var privKey = _certService.GetPrivateKey(selectedCert, password);

                // 성공 시 설정 저장
                ResultConfig = new SignatureConfig
                {
                    Certificate = selectedCert,
                    PrivateKey = privKey
                };

                MessageBox.Show($"로그인 성공!\n사용자: {selectedCert.Subject}");
                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"로그인 실패: {ex.Message}", "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtPassword.SelectAll();
                TxtPassword.Focus();
            }
        }
    }
}