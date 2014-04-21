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

        public static bool HasDuplicates(this Range value)
        {
            int TotalRows = value.Rows.Count;
            int TotalCols = value.Columns.Count;
            List<Range> ItemList = new List<Range>();

            for (int x = 1; x <= TotalCols; x++)
            {
                for (int i = 1; i <= TotalRows; i++)
                {
                    ItemList.Add(value.Cells[i, x]);
                }
            }

            return ItemList.Count != ItemList.Distinct().Count();
        }

        /// <summary>
        /// Converts the data to a string.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Returns the value as a string.</returns>
        public static string Value3(this Range value)
        {
            string Result;

            if (value.Value == null)
                Result = String.Empty;
            else if (String.IsNullOrEmpty(value.Value.ToString()))
                Result = String.Empty;
            else
                Result = value.Value.ToString();
            return Result;
        }

        /// <summary>
        /// Converts the data to a string.
        /// </summary>
        /// <param name="value"></param>
        /// <returns>Returns the value as a string.</returns>
        public static string Value3(this Range value, string format)
        {
            string Result;

            if (value.Value == null)
                Result = String.Empty;
            else if (String.IsNullOrEmpty(value.Value.ToString()))
                Result = String.Empty;
            else
            {
                Result = value.Value.ToString(format);
            }
            return Result;
        }
    }
}
