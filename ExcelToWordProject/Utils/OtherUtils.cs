using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject.Utils
{
    static class OtherUtils
    {
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
    }
}
