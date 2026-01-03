using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json; // .NET 내장 JSON

namespace MinsPDFViewer
{
    public class HistoryService
    {
        private readonly string _historyPath;
        private Dictionary<string, int> _historyData;

        public HistoryService()
        {
            // C:\Users\사용자\AppData\Roaming\MinsPDFViewer\pdf_history.json 경로 생성
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var folder = Path.Combine(appData, "MinsPDFViewer");

            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            _historyPath = Path.Combine(folder, "pdf_history.json");
            _historyData = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase); // 대소문자 무시

            LoadHistory();
        }

        // 파일 경로에 해당하는 마지막 페이지 번호 가져오기
        public int GetLastPage(string filePath)
        {
            if (_historyData.TryGetValue(filePath, out int pageIndex))
            {
                return pageIndex;
            }
            return 0; // 기록 없으면 0페이지
        }

        // 마지막 페이지 번호 저장 (메모리 상)
        public void SetLastPage(string filePath, int pageIndex)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            // 이미 있으면 덮어쓰기, 없으면 추가
            _historyData[filePath] = pageIndex;
        }

        // 파일로 영구 저장
        public void SaveHistory()
        {
            try
            {
                var json = JsonSerializer.Serialize(_historyData);
                File.WriteAllText(_historyPath, json);
            }
            catch { /* 저장 실패는 조용히 무시 */ }
        }

        // 파일에서 불러오기
        private void LoadHistory()
        {
            try
            {
                if (File.Exists(_historyPath))
                {
                    var json = File.ReadAllText(_historyPath);
                    var data = JsonSerializer.Deserialize<Dictionary<string, int>>(json);
                    if (data != null)
                        _historyData = data;
                }
            }
            catch { /* 로드 실패 시 빈 상태로 시작 */ }
        }
    }
}