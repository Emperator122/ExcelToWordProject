using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelToWordProject.Syllabus.Tags;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace ExcelToWordProject.Utils
{
    internal static class DocxExtensions
    {
        public static void InsertPureXmlAfterSelf(this Paragraph paragraph, XElement xmlElement, DocX document)
        {
            var p = paragraph.InsertParagraphAfterSelf("test");
            Type[] types = { typeof(DocX), typeof(XElement), typeof(int), typeof(ContainerType) };
            var p2 = typeof(Paragraph)
                .GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, types, null)
                ?.Invoke(new object[] { document, xmlElement, 0, ContainerType.None });
            p.InsertParagraphAfterSelf(p2 as Paragraph);
            paragraph.Remove(false);
        }

        public static void InsertLinesAfterSelf(this Paragraph paragraph, string[] lines, BaseSyllabusTag tag)
        {
            var p = paragraph.InsertParagraphAfterSelf(paragraph);
            for (var i = 0; i < lines.Length; i++)
            {
                p.ReplaceText(tag.Tag, lines[i]);
                if (i != lines.Length - 1)
                    p = p.InsertParagraphAfterSelf(paragraph);
            }

            paragraph.Remove(false);
        }


    }
}
