using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MinsPDFViewer
{
    // [필수] XAML에서 사용하기 위한 컨버터 클래스
    public class BoolToFontWeightConverter : IValueConverter
    {
        // View에서 바인딩할 때 (bool -> FontWeight)
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isBold && isBold)
                return FontWeights.Bold;
            return FontWeights.Normal;
        }

        // View에서 값을 다시 저장할 때 (FontWeight -> bool)
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FontWeight weight)
                return weight == FontWeights.Bold;
            return false;
        }
    }
}