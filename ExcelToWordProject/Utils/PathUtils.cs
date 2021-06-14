using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToWordProject.Utils
{
    static class PathUtils
    {
        public static string RemoveIllegalFileNameCharacters(string fileName)
        {
            string regexSearch = new string(Path.GetInvalidFileNameChars());
            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            return r.Replace(fileName, "");
        } 
    }
}
