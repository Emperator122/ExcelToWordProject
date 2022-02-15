using ExcelToWordProject.Models;
using Microsoft.Data.Sqlite;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
                $"SELECT * FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{Key}\' "
                + $"AND {DatabaseStrings.TextBlockConditionColumnName} = \'{ToXml()}\'";

            string result = "";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                using (SqliteDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        int i = 0;
                        while (reader.Read())   // построчно считываем данные
                        {
                            string value = reader.GetString(3);
                            result += value;
                            i++;
                            break; // TODO: remove
                        }
                    }
                }
            }
            return result;
        }

        public void SaveToDatabase(string value) 
        {
            string sqlExpression =
                $"INSERT INTO {DatabaseStrings.TextBlockTagTableName} " +
                $"({DatabaseStrings.TextBlockKeyColumnName}, {DatabaseStrings.TextBlockConditionColumnName}, " +
                $"{DatabaseStrings.TextBlockValueColumnName}) "
                + $"VALUES (@key, @condition, @value)";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", Key));
                command.Parameters.Add(new SqliteParameter("@condition", ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.ExecuteNonQuery();
            }
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
        public string Condition { get; set; }
        public string Delimiter { get; set; }

        public TextBlockCondition(string tagName, string condition, string delimiter = null)
        {
            TagName = tagName;
            Condition = condition;
            Delimiter = delimiter;
        }

        public TextBlockCondition() { } // для сериализации

    }
}
