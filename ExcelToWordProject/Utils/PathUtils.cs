using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Words.NET;

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

        public static string FixFileNameLimit(string fileName)
        {
            string name = Path.GetFileName(fileName);
            string ext = Path.GetExtension(fileName);
            if (name.Length >= 255)
                return name.Substring(0, 254 - ext.Length) + ext;
            else
                return fileName;

        }

        public static DocX CopyFile(string baseDocumentPath, string copyPath, bool randomName = false)
        {
            try
            {
                DocX doc = DocX.Load(baseDocumentPath);

                if(randomName)
                {
                    string newName = Path.GetRandomFileName() + Path.GetExtension(copyPath);
                    string newPath = Path.Combine(Path.GetDirectoryName(copyPath), newName);
                    doc.SaveAs(newPath);
                    return DocX.Load(newPath);
                }
                else
                {
                    doc.SaveAs(copyPath);
                    return DocX.Load(copyPath);
                }
            }
            catch
            {
                if (!randomName)
                    return CopyFile(baseDocumentPath, copyPath, true);
                else
                    return null;
            }
        }
    }
}
