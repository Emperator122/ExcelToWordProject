using System;
using System.IO;
using System.Text.RegularExpressions;
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
                using (var doc = DocX.Load(baseDocumentPath))
                {
                    if (randomName)
                    {
                        var newName = Path.GetRandomFileName() + Path.GetExtension(copyPath);
                        var newPath = Path.Combine(Path.GetDirectoryName(copyPath), newName);
                        doc.SaveAs(newPath);
                        return DocX.Load(newPath);
                    }
                    doc.SaveAs(copyPath);
                    return DocX.Load(copyPath);
                }
            }
            catch
            {
                if (!randomName)
                    return CopyFile(baseDocumentPath, copyPath, true);
                return null;
            }
        }

        /// <summary>
        /// Creates a relative path from one file or folder to another.
        /// P.S. https://stackoverflow.com/questions/275689/how-to-get-relative-path-from-absolute-path
        /// </summary>
        /// <param name="fromPath">Contains the directory that defines the start of the relative path.</param>
        /// <param name="toPath">Contains the path that defines the endpoint of the relative path.</param>
        /// <returns>The relative path from the start directory to the end path or <c>toPath</c> if the paths are not related.</returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="UriFormatException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public static string MakeRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) throw new ArgumentNullException("fromPath");
            if (string.IsNullOrEmpty(toPath)) throw new ArgumentNullException("toPath");

            var fromUri = new Uri(fromPath);
            var toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme) { return toPath; } // path can't be made relative.

            var relativeUri = fromUri.MakeRelativeUri(toUri);
            var relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (toUri.Scheme.Equals("file", StringComparison.InvariantCultureIgnoreCase))
            {
                relativePath = relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            }

            return relativePath;
        }
    }
}
