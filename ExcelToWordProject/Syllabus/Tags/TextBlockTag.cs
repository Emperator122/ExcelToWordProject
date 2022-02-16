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

        public string GetValue2(Module module = null, List<Content> contentList = null, DataSet excelData = null)
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
                            break; // TODO: обработка множества значений
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

        public bool CheckConditions(List<BaseSyllabusTag> tags, Module module = null, List<Content> contentList = null, DataSet excelData = null)
        {
            foreach (TextBlockCondition condition in Conditions)
            {
                BaseSyllabusTag tag = tags.Find(
                    _tag => _tag.Key == condition.TagName
                );
                if (tag == null) return false;
                string tagValue = 
                    tag.GetValue(
                        module: module,
                        contentList: contentList,
                        excelData: excelData
                    );
                if (condition.Delimiter == null)
                {
                    if (tagValue != condition.Condition) 
                        return false;
                }
                else
                {
                    bool check = tagValue.Split(new string[] { condition.Delimiter }, StringSplitOptions.None)
                        .Any(el => el == condition.Condition);
                    if (!check) return false;
                }
            }
            return true;
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
        public string Delimiter { get => delimiter; set => delimiter = value != "" ? value : "\n"; }

        private string delimiter;

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
