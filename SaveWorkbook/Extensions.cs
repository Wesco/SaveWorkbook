using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace SaveWorkbook
{
    static class Extensions
    {
        public static string Right(this string value, int length)
        {
            return value.Substring(value.Length - length);
        }

        public static string Left(this string value, int length)
        {
            return value.Substring(0, length);
        }

        public static string Find(this string value, string text)
        {
            int index = 0;

            index = value.IndexOf(text, 0);

            if (index >= 0)
                return value.Substring(index, text.Length);
            else
                return String.Empty;
        }

        public static Range GetValue(this Range value)
        {
            if (value.Value == null)
                value.Value = String.Empty;

            return value;
        }

        public static string RemoveWhiteSpace(this string value)
        {
            value = value.Replace((char)0xA0, ' ');
            value = value.Replace(" ", String.Empty);
            value = value.Replace("\t", String.Empty);

            return value;
        }

        public static string SingleSpace(this string value)
        {
            while (value.Contains("  "))
                value = value.Replace("  ", " ");

            return value;
        }
    }
}
