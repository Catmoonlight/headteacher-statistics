using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HogStatGenerator
{
    public static class StringExtensions
    {
        public static bool IsNumber(this string str)
        {
            return Int32.TryParse(str, out var result);
        }

        public static string ToLowerRus(this string str)
        {
            for(int i = 0; i < str.Length; i++)
            {
                str = str.Replace('А', 'а');
                str = str.Replace('Б', 'б');
                str = str.Replace('В', 'в');
                str = str.Replace('Г', 'г');
                str = str.Replace('Д', 'д');
                str = str.Replace('Е', 'е');
                str = str.Replace('Ё', 'ё');
                str = str.Replace('Ж', 'ж');
                str = str.Replace('З', 'з');
                str = str.Replace('И', 'и');
                str = str.Replace('Й', 'й');
                str = str.Replace('К', 'к');
                str = str.Replace('Л', 'л');
                str = str.Replace('М', 'м');
                str = str.Replace('Н', 'н');
                str = str.Replace('О', 'о');
                str = str.Replace('П', 'п');
                str = str.Replace('Р', 'р');
                str = str.Replace('С', 'с');
                str = str.Replace('Т', 'т');
                str = str.Replace('У', 'у');
                str = str.Replace('Ф', 'ф');
                str = str.Replace('Х', 'х');
                str = str.Replace('Ц', 'ц');
                str = str.Replace('Ч', 'ч');
                str = str.Replace('Ш', 'ш');
                str = str.Replace('Щ', 'щ');
                str = str.Replace('Ъ', 'ъ');
                str = str.Replace('Ы', 'ы');
                str = str.Replace('Ь', 'ь');
                str = str.Replace('Э', 'э');
                str = str.Replace('Ю', 'ю');
                str = str.Replace('Я', 'я');
            }
            return str;
        }
    }
}
