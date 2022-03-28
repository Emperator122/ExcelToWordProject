using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Remoting.Messaging;

namespace ExcelToWordProject.Utils
{
    static class OtherUtils
    {
        public static bool TryParseInt(string s, out int result) => int.TryParse(s, NumberStyles.Any,
            (IFormatProvider) CultureInfo.InvariantCulture, out result);
        public static int StrToInt(string str)
        {
            int result;
            if (!int.TryParse(str, out result))
                return 0;
            return result;
        }

        public static List<int> StrToListInt(string str)
        {
            List<int> result = new List<int>();
            for (int i = 0; i < str.Length; i++)
                result.Add(StrToInt(str[i].ToString()));
            return result;
        }

        public static string ListToDelimiteredString(string delimiter, string endDelimiter, List<string> list)
        {
            string result = "";
            for (int i = 0; i < list.Count(); i++)
            {
                result += (i == list.Count() - 1) ? list[i] + endDelimiter : list[i] + delimiter;
            }
            return result;
        }

        public static string ListToDelimiteredString(string delimiter, string endDelimiter, List<int> list)
        {
            string result = "";
            for (int i = 0; i < list.Count(); i++)
            {
                result += (i == list.Count() - 1) ? list[i] + endDelimiter : list[i] + delimiter;
            }
            return result;
        }

        private static readonly string[][] EscapePattern =
        {
                new []{"\n", "\\n"},
                new []{"\r", "\\r"},
                new []{"\t", "\\t"},
        };
        public static string EscapeString(string str)
        {
            if (str == null)
                return null;
            var result = str;
            foreach (var pattern in EscapePattern)
                result = result.Replace(pattern[0], pattern[1]);
            return result;
        }

        public static string RestoreEscapedString(string str)
        {
            if (str == null)
                return null;
            var result = str;
            foreach (var pattern in EscapePattern)
                result = result.Replace(pattern[1], pattern[0]);
            return result;
        }
    }
}
