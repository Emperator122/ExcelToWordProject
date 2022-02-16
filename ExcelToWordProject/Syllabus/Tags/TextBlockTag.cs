using ExcelToWordProject.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ExcelToWordProject.Syllabus.Tags
{
    public class TextBlockTag : BaseSyllabusTag
    {

        private int MaxConditionLength { get => Conditions.Max(condition => condition.Length); }

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
            return "";
        }

        public string[] GetValue2(Module module = null, List<Content> contentList = null, DataSet excelData = null)
        {
            string DatabaseLoadHandler(TextBlockTag tag)
            {
                string sqlExpression =
                $"SELECT * FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{tag.Key}\' "
                + $"AND {DatabaseStrings.TextBlockConditionColumnName} = \'{tag.ToXml()}\'";

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
                                break; // TODO: обработка множества значений
                            }
                        }
                    }
                }
                return result;
            }

            // Все это вообще не оптимально, но хоть как
            int maxConditionLength = MaxConditionLength;
            if (maxConditionLength == 1)
                return new string[] { DatabaseLoadHandler(this) };

            string[] values = new string[maxConditionLength];
            TextBlockTag[] splitted = Split();
            for (int i = 0; i < splitted.Length; i++)
                values[i] = splitted[i].GetValue2()[0];
            return values;
        }

        public void SaveToDatabase(string[] data) 
        {
            Debug.Assert(data != null && data.Count() != 0);
            void DatabaseSaveHandler(string value, TextBlockTag tag)
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
                    command.Parameters.Add(new SqliteParameter("@key", tag.Key));
                    command.Parameters.Add(new SqliteParameter("@condition", tag.ToXml()));
                    command.Parameters.Add(new SqliteParameter("@value", value));
                    command.ExecuteNonQuery();
                }
            }

            // Все это вообще не оптимально, но хоть как

            int maxConditionLength = MaxConditionLength;
            if (maxConditionLength == 1)
            {
                Debug.Assert(data.Count() == 1);
                DatabaseSaveHandler(data[0], this);
                return;
            }

            TextBlockTag[] splitted = Split();
            Debug.Assert(data.Count() == splitted.Length);
            for (int i = 0; i < splitted.Length; i++)
                splitted[i].SaveToDatabase(new string[] { data[i] });
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

        public TextBlockTag[] Split()
        {
            int maxConditionLength = Conditions.Max(condition => condition.Length);
            int minDelimeteredConditionLength =
                Conditions.Min(condition => (condition.Delimiter != null) ? condition.Length : maxConditionLength);
            Debug.Assert(minDelimeteredConditionLength == maxConditionLength);

            List<TextBlockCondition[]> splittedConditions = new List<TextBlockCondition[]>();
            foreach (TextBlockCondition condition in Conditions)
                splittedConditions.Add(condition.Split());


            TextBlockTag[] splittedBlockTags = new TextBlockTag[maxConditionLength];
            for (int i = 0; i < maxConditionLength; i++)
            {
                TextBlockCondition[] conditions = new TextBlockCondition[Conditions.Length];
                for (int j = 0; j < conditions.Length; j++)
                    conditions[j] = splittedConditions[j][splittedConditions[j].Length > 1 ? i : 0];
                splittedBlockTags[i] = new TextBlockTag(
                    key: Key,
                    listName: ListName,
                    conditions: conditions,
                    description: Description
                );
            }

            return splittedBlockTags;
        }

        public static List<TextBlockTag> GetAllTextBlockTags()
        {
            string sqlExpression =
                $"SELECT {DatabaseStrings.TextBlockConditionColumnName} FROM {DatabaseStrings.TextBlockTagTableName}";

            List<TextBlockTag> result = new List<TextBlockTag>();
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                using (SqliteDataReader reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        while (reader.Read())   // построчно считываем данные
                        {
                            string xml = reader.GetString(0);
                            result.Add(FromXml(xml));
                        }
                    }
                }
            }
            return result;
        }
    }

    public class TextBlockCondition
    {
        public string TagName { get; set; }
        public string Condition { get; set; }
        public string Delimiter { get; set; }

        public int Length { get => Subconditions.Length; }

        public string[] Subconditions 
        { 
            get 
            {
                if (Delimiter == null)
                    return new string[] { Condition };
                else
                    return Condition.Split(new string[] { Delimiter }, StringSplitOptions.None); ;
            } 
        }

        public TextBlockCondition(string tagName, string condition, string delimiter = null)
        {
            TagName = tagName;
            Condition = condition;
            Delimiter = delimiter;
        }

        public TextBlockCondition() { } // для сериализации

        public TextBlockCondition[] Split()
        {
            string[] subconditions = Subconditions;
            TextBlockCondition[] splittedConditions = new TextBlockCondition[subconditions.Length];
            for (int i = 0; i < subconditions.Length; i++)
            {
                splittedConditions[i] = 
                    new TextBlockCondition(
                        tagName: TagName,
                        condition: subconditions[i],
                        delimiter: Delimiter
                    );
            }
            return splittedConditions;
        }

    }
}
