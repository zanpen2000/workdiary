using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WorkDiary
{
    using System.Windows.Controls;

    /// <summary>
    /// 
    /// </summary>
    class FilePathValidRule:ValidationRule
    {
        public override ValidationResult Validate(object value, System.Globalization.CultureInfo cultureInfo)
        {
            string path = value.ToString();
            return new ValidationResult(System.IO.File.Exists(path), "");
        }
    }
}
