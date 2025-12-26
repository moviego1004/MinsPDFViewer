using System.Windows;

namespace MinsPDFViewer
{
    // WPF의 리스트 가상화(Virtualization) 시 데이터 바인딩이 끊기는 문제를 해결하는 클래스
    public class BindingProxy : Freezable
    {
        protected override Freezable CreateInstanceCore()
        {
            return new BindingProxy();
        }

        public object Data
        {
            get { return (object)GetValue(DataProperty); }
            set { SetValue(DataProperty, value); }
        }

        // Using a DependencyProperty as the backing store for Data. This enables animation, styling, binding, etc...
        public static readonly DependencyProperty DataProperty =
            DependencyProperty.Register("Data", typeof(object), typeof(BindingProxy), new PropertyMetadata(null));
    }
}