using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using ExcelToWordProject.Models;
using Microsoft.Data.Sqlite;

namespace ExcelToWordProject.Syllabus.Tags
{
    public class TextBlockTag : BaseSyllabusTag
    {
        public TextBlockCondition[] Conditions { get; set; }

        public bool ValueIsPureXML { get; set; }

        public string DefaultValue => GetDefaultValue(Key);

        public KeyValuePair<TextBlockTag, string> DefaultTag => GetDefaultTag(Key);

        public bool IsDefault => Conditions.Length == 0;

        public TextBlockTag(
            string key,
            IEnumerable<TextBlockCondition> conditions,
            string description = "",
            bool valueIsPureXml = false
        ) : base(key, "", description)
        {
            Conditions = conditions.OrderBy(condition => condition.TagName).ToArray();
            ValueIsPureXML = valueIsPureXml;
        }

        public TextBlockTag() // для сериализации
        {
        }

        public override string GetValue(Module module = null, List<Content> contentList = null,
            DataSet excelData = null)
        {
            return "";
        }

        public string GetValue2(Module module = null, List<Content> contentList = null, DataSet excelData = null)
        {
            var sqlExpression =
                $"SELECT * FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{Key}\' "
                + $"AND {DatabaseStrings.TextBlockConditionColumnName} = \'{ToXml()}\'";

            var result = "";
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        var i = 0;
                        while (reader.Read()) // построчно считываем данные
                        {
                            var value = reader.GetString(3);
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
            var sqlExpression =
                $"INSERT INTO {DatabaseStrings.TextBlockTagTableName} " +
                $"({DatabaseStrings.TextBlockKeyColumnName}, {DatabaseStrings.TextBlockConditionColumnName}, " +
                $"{DatabaseStrings.TextBlockValueColumnName}) "
                + "VALUES (@key, @condition, @value)";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", Key));
                command.Parameters.Add(new SqliteParameter("@condition", ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.ExecuteNonQuery();
            }
        }

        public void SetAsDefault()
        {
            SetDefaultValue(Key, GetValue2(), Description, ValueIsPureXML);
        }

        public string ToXml()
        {
            var serializer = new XmlSerializer(typeof(TextBlockTag));
            using (var textWriter = new StringWriter())
            {
                serializer.Serialize(textWriter, this);
                return textWriter.ToString();
            }
        }

        public static TextBlockTag FromXml(string xml)
        {
            var serializer = new XmlSerializer(typeof(TextBlockTag));
            using (TextReader reader = new StringReader(xml))
            {
                return (TextBlockTag) serializer.Deserialize(reader);
            }
        }

        public bool CheckConditions(List<BaseSyllabusTag> tags, Module module = null, List<Content> contentList = null,
            DataSet excelData = null)
        {
            foreach (var condition in Conditions)
            {
                var tag = tags.Find(
                    baseSyllabusTag => baseSyllabusTag.Key == condition.TagName && baseSyllabusTag.Active
                );
                if (tag == null) return false;
                var tagValue = tag.GetValue(module, contentList, excelData);
                if (condition.Delimiter == null)
                {
                    if (tagValue != condition.Condition)
                        return false;
                }
                else
                {
                    var check = tagValue.Split(new[] {condition.Delimiter}, StringSplitOptions.None)
                        .Any(el => el == condition.Condition);
                    if (!check) return false;
                }
            }

            return true;
        }

        public static List<TextBlockTag> GetAllTextBlockTags()
        {
            var sqlExpression =
                $"SELECT {DatabaseStrings.TextBlockConditionColumnName} FROM {DatabaseStrings.TextBlockTagTableName}";

            var result = new List<TextBlockTag>();
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                        while (reader.Read()) // построчно считываем данные
                        {
                            var xml = reader.GetString(0);
                            result.Add(FromXml(xml));
                        }
                }
            }

            return result;
        }

        public static string GetDefaultValue(string tagKey)
        {
            var sqlExpression =
                $"SELECT * FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{tagKey}\' "
                + $"AND {DatabaseStrings.TextBlockIsDefaultColumnName} = 1";

            string result = null;
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        result = string.Empty;
                        var i = 0;
                        while (reader.Read()) // построчно считываем данные
                        {
                            var value = reader.GetString(3);
                            result += value;
                            i++;
                            break; // TODO: обработка множества значений
                        }
                    }
                }
            }

            return result;
        }

        public static KeyValuePair<TextBlockTag, string> GetDefaultTag(string tagKey)
        {
            var sqlExpression =
                $"SELECT * FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = \'{tagKey}\' "
                + $"AND {DatabaseStrings.TextBlockIsDefaultColumnName} = 1";

            string tagValue = null;
            TextBlockTag tag = null;
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        reader.Read();

                        tagValue = reader.GetString(3);
                        tag = FromXml(reader.GetString(2));
                    }
                }
            }

            return new KeyValuePair<TextBlockTag, string>(tag, tagValue);
        }

        public static void SetDefaultValue(string tagKey, string value, string description="", bool valueIsPureXml = false)
        {
            if (GetDefaultValue(tagKey) != null)
                throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.UniqueDefaultValueError);

            var sqlExpression =
                $"INSERT INTO {DatabaseStrings.TextBlockTagTableName} " +
                $"({DatabaseStrings.TextBlockKeyColumnName}, {DatabaseStrings.TextBlockConditionColumnName}, " +
                $"{DatabaseStrings.TextBlockValueColumnName}, {DatabaseStrings.TextBlockIsDefaultColumnName}) "
                + "VALUES (@key, @condition, @value, @isDefault)";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var tempCondition = new TextBlockTag(tagKey, Array.Empty<TextBlockCondition>(), description, valueIsPureXml);
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", tagKey));
                command.Parameters.Add(new SqliteParameter("@condition", tempCondition.ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.Parameters.Add(new SqliteParameter("@isDefault", 1));
                command.ExecuteNonQuery();
            }
        }

    }

    public class TextBlockCondition
    {
        private string _delimiter;

        public TextBlockCondition(string tagName, string condition, string delimiter = null)
        {
            TagName = tagName;
            Condition = condition;
            Delimiter = delimiter;
        }

        public TextBlockCondition()
        {
        } // для сериализации

        public string TagName { get; set; }
        public string Condition { get; set; }

        public string Delimiter
        {
            get => _delimiter;
            set => _delimiter = value != "" ? value : "\n";
        }
    }
}