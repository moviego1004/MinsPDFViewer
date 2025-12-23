using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using Docnet.Core.Readers;

namespace MinsPDFViewer
{
    // 탭별 데이터(문서) 상태 관리 모델
    public class PdfDocumentModel : INotifyPropertyChanged
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "제목 없음";
        
        // 문서 객체와 페이지 목록
        public IDocReader? DocReader { get; set; }
        public ObservableCollection<PdfPageViewModel> Pages { get; set; } = new ObservableCollection<PdfPageViewModel>();
        
        // 탭별 독립적인 줌 레벨
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

        // 탭 전환 시 복원할 스크롤 위치
        public double SavedVerticalOffset { get; set; } = 0;
        public double SavedHorizontalOffset { get; set; } = 0;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        
    }
}