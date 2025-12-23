using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Media;
using Docnet.Core.Readers;

namespace MinsPDFViewer
{
    // 각 탭(문서)의 상태를 관리하는 클래스
    public class PdfDocumentModel : INotifyPropertyChanged
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "제목 없음";
        
        // 탭마다 독립적인 DocReader와 페이지 목록을 가짐 (잔상 방지 핵심)
        public IDocReader? DocReader { get; set; }
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        // 탭별 줌 상태 유지
        private double _zoom = 1.0;
        public double Zoom 
        { 
            get => _zoom; 
            set 
            { 
                _zoom = Math.Clamp(value, 0.2, 5.0); 
                OnPropertyChanged(nameof(Zoom)); 
                OnPropertyChanged(nameof(ZoomPercentText));
            } 
        }
        public string ZoomPercentText => $"{Math.Round(Zoom * 100)}%";

        // 탭 전환 시 스크롤 위치 기억
        public double SavedVerticalOffset { get; set; } = 0;
        public double SavedHorizontalOffset { get; set; } = 0;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}