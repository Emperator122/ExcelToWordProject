using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Xceed.Document.NET;

// ReSharper disable InconsistentNaming

namespace ExcelToWordProject.Utils
{
    /// <summary>
    /// Класс, призванный расприватить методы класса Document
    /// из библиотеки DocX, а также на их основе добавить новые
    /// методы.
    /// Создан ради реализации функцинала вставки документа после
    /// заданного параграфа
    /// </summary>
    internal class CustomDocument : IDisposable
    {
        internal static XNamespace w = (XNamespace) "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        internal static XNamespace rel = (XNamespace) "http://schemas.openxmlformats.org/package/2006/relationships";

        internal static XNamespace r =
            (XNamespace) "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal static XNamespace m = (XNamespace) "http://schemas.openxmlformats.org/officeDocument/2006/math";

        internal static XNamespace customPropertiesSchema =
            (XNamespace) "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

        internal static XNamespace customVTypesSchema =
            (XNamespace) "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

        internal static XNamespace wp =
            (XNamespace) "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        internal static XNamespace a = (XNamespace) "http://schemas.openxmlformats.org/drawingml/2006/main";
        internal static XNamespace c = (XNamespace) "http://schemas.openxmlformats.org/drawingml/2006/chart";
        internal static XNamespace pic = (XNamespace) "http://schemas.openxmlformats.org/drawingml/2006/picture";

        internal static XNamespace n =
            (XNamespace) "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";

        internal static XNamespace v = (XNamespace) "urn:schemas-microsoft-com:vml";
        internal static XNamespace mc = (XNamespace) "http://schemas.openxmlformats.org/markup-compatibility/2006";

        internal static XNamespace wps =
            (XNamespace) "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        internal static XNamespace w14 = (XNamespace) "http://schemas.microsoft.com/office/word/2010/wordml";
        internal static XNamespace w15 = (XNamespace) "http://schemas.microsoft.com/office/word/2012/wordml";
        internal static XNamespace o = (XNamespace) "urn:schemas-microsoft-com:office:office";
        internal static XNamespace d = (XNamespace) "http://www.w3.org/2000/09/xmldsig#";
        internal static XNamespace doffice = (XNamespace) "http://schemas.microsoft.com/office/2006/digsig";
        internal static XNamespace w10 = (XNamespace) "urn:schemas-microsoft-com:office:word";
        public Document Document { get; }

        private XDocument _mainDoc
        {
            get => typeof(Document).GetField("_mainDoc", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_mainDoc", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }


        private XDocument _styles
        {
            get =>
                typeof(Document).GetField("_styles", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_styles", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private XDocument _settings
        {
            get =>
                typeof(Document).GetField("_settings", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_settings", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private XDocument _endnotes
        {
            get =>
                typeof(Document).GetField("_endnotes", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_endnotes", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        internal XDocument _footnotes
        {
            get =>
                typeof(Document).GetField("_footnotes", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_footnotes", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private XDocument _stylesWithEffects
        {
            get =>
                typeof(Document).GetField("_stylesWithEffects", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_stylesWithEffects", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private XDocument _numbering
        {
            get =>
                typeof(Document).GetField("_numbering", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_numbering", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private XDocument _fontTable
        {
            get =>
                typeof(Document).GetField("_fontTable", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as XDocument;
            set => typeof(Document).GetField("_fontTable", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private Package _package =>
            typeof(Document).GetField("_package", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.GetValue(Document) as Package;

        private static List<string> _imageContentTypes
        {
            get =>
                typeof(Document).GetField("_imageContentTypes", BindingFlags.NonPublic | BindingFlags.Static)
                    ?.GetValue(null) as List<string>;
        }

        private PackagePart _settingsPart
        {
            get =>
                typeof(Document).GetField("_mainDoc", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_mainDoc", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _endnotesPart
        {
            get =>
                typeof(Document).GetField("_endnotesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_endnotesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _footnotesPart
        {
            get =>
                typeof(Document).GetField("_footnotesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_footnotesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _stylesPart
        {
            get =>
                typeof(Document).GetField("_stylesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_stylesPart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _stylesWithEffectsPart
        {
            get =>
                typeof(Document).GetField("_stylesWithEffectsPart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_stylesWithEffectsPart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _numberingPart
        {
            get =>
                typeof(Document).GetField("_numberingPart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_numberingPart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private PackagePart _fontTablePart
        {
            get =>
                typeof(Document).GetField("_fontTablePart", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?.GetValue(Document) as PackagePart;
            set => typeof(Document).GetField("_fontTablePart", BindingFlags.NonPublic | BindingFlags.Instance)
                ?.SetValue(Document, value);
        }

        private void merge_customs(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc)
        {
            var method = typeof(Document)
                .GetMethod("merge_customs", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc
                }
            );
        }

        private void merge_styles(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc,
            Document remote,
            XDocument remote_footnotes,
            XDocument remote_endnotes,
            MergingMode mergingMode)
        {
            var method = typeof(Document)
                .GetMethod("merge_styles", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc,
                    remote,
                    remote_footnotes,
                    remote_endnotes,
                    mergingMode,
                }
            );
        }

        private void merge_numbering(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc,
            Document remote)
        {
            var method = typeof(Document)
                .GetMethod("merge_numbering", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc,
                    remote
                }
            );
        }

        private void merge_images(
            PackagePart remote_pp,
            Document remote_document,
            XDocument remote_mainDoc,
            string contentType)
        {
            var method = typeof(Document)
                .GetMethod("merge_images", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    remote_document,
                    remote_mainDoc,
                    contentType
                }
            );
        }

        private void merge_fonts(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc,
            Document remote)
        {
            var method = typeof(Document)
                .GetMethod("merge_fonts", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc,
                    remote
                }
            );
        }

        private void merge_footnotes(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc,
            Document remote,
            XDocument remote_footnotes)
        {
            var method = typeof(Document)
                .GetMethod("merge_footnotes", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc,
                    remote,
                    remote_footnotes
                }
            );
        }

        private void merge_endnotes(
            PackagePart remote_pp,
            PackagePart local_pp,
            XDocument remote_mainDoc,
            Document remote,
            XDocument remote_endnotes)
        {
            var method = typeof(Document)
                .GetMethod("merge_endnotes", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(
                Document,
                new object[]
                {
                    remote_pp,
                    local_pp,
                    remote_mainDoc,
                    remote,
                    remote_endnotes
                }
            );
        }

        private PackagePart clonePackagePart(PackagePart pp)
        {
            var method = typeof(Document)
                .GetMethod("clonePackagePart", BindingFlags.NonPublic | BindingFlags.Instance);
            return method?.Invoke(
                Document,
                new object[]
                {
                    pp
                }
            ) as PackagePart;
        }

        private PackagePart clonePackageRelationship(
            Document remote_document,
            PackagePart pp,
            XDocument remote_mainDoc)
        {
            var method = typeof(Document)
                .GetMethod("clonePackageRelationship", BindingFlags.NonPublic | BindingFlags.Instance);
            return method?.Invoke(
                Document,
                new object[]
                {
                    remote_document,
                    pp,
                    remote_mainDoc
                }
            ) as PackagePart;
        }

        private void UpdateCacheSections()
        {
            var method = typeof(Document)
                .GetMethod("UpdateCacheSections", BindingFlags.NonPublic | BindingFlags.Instance);
            method?.Invoke(Document, Array.Empty<object>());
        }


        public CustomDocument(Document document) => Document = document;

        /// <summary>
        /// Вставка документа после заданного параграфа
        /// </summary>
        /// <param name="remote_document">документ-источник</param>
        /// <param name="paragraph">параграф</param>
        /// <param name="mergingMode">режим</param>
        public void InsertDocument(
            CustomDocument remote_document,
            Paragraph paragraph,
            MergingMode mergingMode = MergingMode.Both)
        {
            XDocument remote_mainDoc = new XDocument(remote_document._mainDoc);
            XDocument remote_footnotes = (XDocument) null;
            if (remote_document._footnotes != null)
                remote_footnotes = new XDocument(remote_document._footnotes);
            XDocument remote_endnotes = (XDocument) null;
            if (remote_document._endnotes != null)
                remote_endnotes = new XDocument(remote_document._endnotes);
            XElement xelement = remote_mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
            XElement first = this._mainDoc.Root.Element(XName.Get("body", w.NamespaceName));
            remote_mainDoc.Descendants(XName.Get("headerReference", w.NamespaceName)).Remove<XElement>();
            remote_mainDoc.Descendants(XName.Get("footerReference", w.NamespaceName)).Remove<XElement>();
            PackagePartCollection parts = remote_document._package.GetParts();
            List<string> stringList = new List<string>()
            {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
                "application/vnd.openxmlformats-package.core-properties+xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                "application/vnd.openxmlformats-package.relationships+xml"
            };
            foreach (PackagePart packagePart1 in parts)
            {
                if (!stringList.Contains(packagePart1.ContentType) &&
                    !_imageContentTypes.Contains(packagePart1.ContentType))
                {
                    if (this._package.PartExists(packagePart1.Uri))
                    {
                        PackagePart part = this._package.GetPart(packagePart1.Uri);
                        switch (packagePart1.ContentType)
                        {
                            case "application/vnd.ms-word.stylesWithEffects+xml":
                                this.merge_styles(packagePart1, part, remote_mainDoc, remote_document.Document,
                                    remote_footnotes, remote_endnotes, mergingMode);
                                break;
                            case "application/vnd.openxmlformats-officedocument.custom-properties+xml":
                                this.merge_customs(packagePart1, part, remote_mainDoc);
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
                                this.merge_endnotes(packagePart1, part, remote_mainDoc, remote_document.Document,
                                    remote_endnotes);
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
                                this.merge_fonts(packagePart1, part, remote_mainDoc, remote_document.Document);
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
                                this.merge_footnotes(packagePart1, part, remote_mainDoc, remote_document.Document,
                                    remote_footnotes);
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
                                this.merge_numbering(packagePart1, part, remote_mainDoc, remote_document.Document);
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
                                this.merge_styles(packagePart1, part, remote_mainDoc, remote_document.Document,
                                    this._footnotes, this._endnotes, mergingMode);
                                break;
                        }
                    }
                    else
                    {
                        PackagePart packagePart2 = this.clonePackagePart(packagePart1);
                        switch (packagePart1.ContentType)
                        {
                            case "application/vnd.ms-word.stylesWithEffects+xml":
                                this._stylesWithEffectsPart = packagePart2;
                                using (TextReader textReader =
                                       (TextReader) new StreamReader(this._stylesWithEffectsPart.GetStream()))
                                {
                                    this._stylesWithEffects = XDocument.Load(textReader);
                                    break;
                                }
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml":
                                this._endnotesPart = packagePart2;
                                this._endnotes = remote_endnotes;
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml":
                                this._fontTablePart = packagePart2;
                                using (TextReader textReader =
                                       (TextReader) new StreamReader(this._fontTablePart.GetStream()))
                                {
                                    this._fontTable = XDocument.Load(textReader);
                                    break;
                                }
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml":
                                this._footnotesPart = packagePart2;
                                this._footnotes = remote_footnotes;
                                break;
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml":
                                this._numberingPart = packagePart2;
                                using (TextReader textReader =
                                       (TextReader) new StreamReader(this._numberingPart.GetStream()))
                                {
                                    this._numbering = XDocument.Load(textReader);
                                    break;
                                }
                            case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml":
                                this._stylesPart = packagePart2;
                                using (TextReader textReader =
                                       (TextReader) new StreamReader(this._stylesPart.GetStream()))
                                {
                                    this._styles = XDocument.Load(textReader);
                                    break;
                                }
                        }

                        this.clonePackageRelationship(remote_document.Document, packagePart1, remote_mainDoc);
                    }
                }
            }
            /*if (useSectionBreak)
                this.MoveSectionIntoLastParagraph(append ? first : xelement);
            else if (append)
                this.ReplaceLastSection(first, xelement);
            else
                this.RemoveLastSection(xelement);*/

            xelement.Elements(XName.Get("sectPr", w.NamespaceName)).LastOrDefault<XElement>()?.Remove();

            foreach (PackageRelationship packageRelationship in remote_document.Document.PackagePart
                         .GetRelationshipsByType(
                             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"))
            {
                string id1 = packageRelationship.Id;
                string id2 = Document.PackagePart.CreateRelationship(packageRelationship.TargetUri,
                    packageRelationship.TargetMode, packageRelationship.RelationshipType).Id;
                foreach (XElement descendant in remote_mainDoc.Descendants(XName.Get("hyperlink", w.NamespaceName)))
                {
                    XAttribute xattribute = descendant.Attribute(XName.Get("id", r.NamespaceName));
                    if (xattribute != null && xattribute.Value == id1)
                        xattribute.SetValue((object) id2);
                }
            }

            foreach (PackageRelationship packageRelationship in remote_document.Document.PackagePart
                         .GetRelationshipsByType(
                             "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"))
            {
                string id3 = packageRelationship.Id;
                string id4 = Document.PackagePart.CreateRelationship(packageRelationship.TargetUri,
                    packageRelationship.TargetMode, packageRelationship.RelationshipType).Id;
                foreach (XElement descendant in remote_mainDoc.Descendants(XName.Get("OLEObject",
                             "urn:schemas-microsoft-com:office:office")))
                {
                    XAttribute xattribute = descendant.Attribute(XName.Get("id", r.NamespaceName));
                    if (xattribute != null && xattribute.Value == id3)
                        xattribute.SetValue((object) id4);
                }
            }

            PackageRelationshipCollection relationshipsByType1 =
                remote_document.Document.PackagePart.GetRelationshipsByType(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
            PackageRelationshipCollection relationshipsByType2 =
                Document.PackagePart.GetRelationshipsByType(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering");
            if (relationshipsByType1.Count<PackageRelationship>() > 0 &&
                relationshipsByType2.Count<PackageRelationship>() == 0)
            {
                foreach (PackageRelationship packageRelationship in relationshipsByType1)
                    Document.PackagePart.CreateRelationship(packageRelationship.TargetUri,
                        packageRelationship.TargetMode, packageRelationship.RelationshipType);
            }

            if (remote_document._fontTablePart != null && this._fontTablePart != null)
            {
                PackageRelationshipCollection relationshipsByType3 =
                    remote_document._fontTablePart.GetRelationshipsByType(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font");
                PackageRelationshipCollection relationshipsByType4 =
                    this._fontTablePart.GetRelationshipsByType(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font");
                if (relationshipsByType3.Count<PackageRelationship>() > 0 &&
                    relationshipsByType4.Count<PackageRelationship>() == 0)
                {
                    foreach (PackageRelationship packageRelationship in relationshipsByType3)
                        this._fontTablePart.CreateRelationship(packageRelationship.TargetUri,
                            packageRelationship.TargetMode, packageRelationship.RelationshipType);
                }
            }

            foreach (PackagePart remote_pp in parts)
            {
                if (_imageContentTypes.Contains(remote_pp.ContentType))
                    this.merge_images(remote_pp, remote_document.Document, remote_mainDoc, remote_pp.ContentType);
            }

            int num1 = 0;
            foreach (XElement descendant in this._mainDoc.Root.Descendants(XName.Get("docPr", wp.NamespaceName)))
            {
                XAttribute xattribute = descendant.Attribute(XName.Get("id"));
                int result;
                if (xattribute != null && OtherUtils.TryParseInt(xattribute.Value, out result) && result > num1)
                    num1 = result;
            }

            int num2 = num1 + 1;
            foreach (XElement descendant in xelement.Descendants(XName.Get("docPr", wp.NamespaceName)))
            {
                descendant.SetAttributeValue(XName.Get("id"), (object) num2);
                ++num2;
            }

            /*if (append)
                first.Add((object)xelement.Elements());
            else
                first.AddFirst((object)xelement.Elements());*/
            paragraph.InsertPureXmlAfterSelf(xelement, Document);
            foreach (XAttribute attribute in remote_mainDoc.Root.Attributes())
            {
                if (this._mainDoc.Root.Attribute(attribute.Name) == null)
                    this._mainDoc.Root.SetAttributeValue(attribute.Name, (object) attribute.Value);
            }

            this.UpdateCacheSections();
        }

        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
