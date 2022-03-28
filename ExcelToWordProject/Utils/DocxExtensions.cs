using System;
using System.Collections.Generic;
using System.IO;
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
        private static XNamespace w = (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static void InsertPureXmlAfterSelf(this Paragraph paragraph, XElement xmlElement, Document document)
        {
            Type[] types = { typeof(DocX), typeof(XElement), typeof(int), typeof(ContainerType) };
            var p2 = typeof(Paragraph)
                .GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, types, null)
                ?.Invoke(new object[] { document, xmlElement, 0, ContainerType.None });
            paragraph.InsertParagraphAfterSelf(p2 as Paragraph);
            var sr = new StreamReader("extracted_style.xml");
            var data = sr.ReadToEnd();
            var styles = GetDocumentStyles(document);
            var style = XElement.Parse(data);
            styles?.Element(w + "styles")?.Add((object)style);

            paragraph.Remove(false);
        }

        public static void InsertDocumentAfterSelf(this Paragraph paragraph, Document sourceDocument)
        {
            DocX doc = DocX.Load("document.docx");
            var customDocument = new CustomDocument(sourceDocument);
            customDocument.InsertDocument(new CustomDocument(doc), paragraph);
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

        private static void InsertStyle()
        {

        }

        private static XDocument GetDocumentStyles(this Document document)
        {
            var styles = typeof(Document).GetField("_styles", BindingFlags.NonPublic | BindingFlags.Instance)?.GetValue(document);
            return styles as XDocument;
        }


    }
}
