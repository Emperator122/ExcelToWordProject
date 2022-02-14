using ExcelToWordProject.Models;
using Microsoft.Data.Sqlite;
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
            Conditions = conditions.OrderBy(condition => condition.TagName).ToArray();
        }

        public TextBlockTag() { } // для сериализации

        public override string GetValue(Module module = null, List<Content> contentList = null, DataSet excelData = null)
        {
            string sqlExpression =
                $"SELECT * FROM {DatabaseStrings.TextBlockValueColumnName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{Key}\' "
                + $"AND {DatabaseStrings.TextBlockConditionColumnName} = \'{ToXml()}\'";
            using (var connection = new SqliteConnection(DatabaseStrings.DatabasePath))
            {
                connection.Open();
                SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                using (SqliteDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        while (reader.Read())   // построчно считываем данные
                        {
                            var id = reader.GetValue(0);
                            var name = reader.GetValue(1);
                            var age = reader.GetValue(2);

                            Console.WriteLine($"{id} \t {name} \t {age}");
                        }
                    }
                }
            }
            return "";
        }

        public void SaveToDatabase(string value) 
        {
            
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
        public string TagName { get; set; }
        public string[] Condition { get; set; } // для тега, содержащего несколько значений, имеем массив
        public string Delimiter { get; set; }

        public TextBlockCondition(string tagName, string[] condition, string delimiter = null)
        {
            TagName = tagName;
            Condition = condition;
            Delimiter = delimiter;
        }

        public TextBlockCondition() { } // для сериализации

    }
}
