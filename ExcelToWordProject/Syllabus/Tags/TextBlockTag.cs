using ExcelToWordProject.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace ExcelToWordProject.Syllabus.Tags
{
    public class TextBlockTag : BaseSyllabusTag
    {
        public TextBlockCondition[] Conditions { get; set; }

        public TextBlockTag(
            string key,
            string listName,
            TextBlockCondition[] conditions,
            string description = ""
        ) : base(
            key: key,
            listName: listName,
            description: description
        )
        {
            Conditions = conditions.OrderBy(condition => condition.Tag.Key).ToArray();
        }

        public TextBlockTag() { } // для сериализации

        public override string GetValue(Module module = null, List<Content> contentList = null, DataSet excelData = null)
        {
            throw new NotImplementedException();
        }

        public string ToXml()
        {
            XmlSerializer serializer = new XmlSerializer(typeof(TextBlockTag));
            using (StringWriter textWriter = new StringWriter())
            {
                serializer.Serialize(textWriter, this);
                return textWriter.ToString();
            }
        }

        public static TextBlockTag FromXml(string xml)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(TextBlockTag));
            using (TextReader reader = new StringReader(xml))
            {
                return (TextBlockTag)serializer.Deserialize(reader);
            }
        }
    }

    public class TextBlockCondition
    {
        [System.Xml.Serialization.XmlElementAttribute("SmartSyllabusTag", typeof(SmartSyllabusTag))]
        [System.Xml.Serialization.XmlElementAttribute("DefaultSyllabusTag", typeof(DefaultSyllabusTag))]
        [System.Xml.Serialization.XmlElementAttribute("TextBlockTag", typeof(TextBlockTag))]
        public BaseSyllabusTag Tag { get; set; }
        public string[] Condition { get; set; } // для тега, содержащего несколько значений, имеем массив
        public string Delimiter { get; set; }

        public TextBlockCondition(BaseSyllabusTag tag, string[] condition, string delimiter = null)
        {
            Tag = tag;
            Condition = condition;
            Delimiter = delimiter;
        }

        public TextBlockCondition() { } // для сериализации

    }
}
