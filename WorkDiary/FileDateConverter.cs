using ClassLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WorkDiary
{
    public class FileDateConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                string filename = values[0].ToString();
                string filename1 = filename.Substring(0, filename.Length - 12);

                string date = values[1].ToString();

                var q = from n in date.Split('/')
                        select n.PadLeft(2, '0');

                string filename2 = string.Join("", q.ToArray());

                return filename1 + filename2 + ".xls";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            //throw new NotImplementedException();
            return null;
        }
    }
}
