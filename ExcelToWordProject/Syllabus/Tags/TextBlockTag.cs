﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using ExcelToWordProject.Models;
using Microsoft.Data.Sqlite;

namespace ExcelToWordProject.Syllabus.Tags
{
    public class TextBlockTag : BaseSyllabusTag
    {
        /// <summary>
        /// Используется для получения Id из БД.
        /// Изменение вне работы с БД не предусмотрено.
        /// </summary>
        [XmlIgnore]
        private int Id { get; set; }

        public bool HasId => Id != -1;

        public bool CanStoreInDataBase => HasId || GetValue2() == null;
        public bool HasDefaultValue => DefaultValue != null;

        public string DefaultValue => GetDefaultValue(Key);

        public TextBlockCondition[] Conditions { get => _conditions; set => _conditions = value.OrderBy(condition => condition.TagName).ToArray(); }
        private TextBlockCondition[] _conditions;

        /// <summary>
        /// Используется для получения IsDefault из БД.
        /// Изменение вне работы с БД не предусмотрено.
        /// </summary>
        [XmlIgnore]
        public bool IsDefault { get; set; }

        public bool CanBeDefault => Conditions.Length == 0;

        public TextBlockTag(
            string key,
            TextBlockCondition[] conditions,
            string description = ""
        ) : base(key, "", description)
        {
            Conditions = conditions;
            Id = -1;
            IsDefault = false;
        }
        
        public TextBlockTag() // для сериализации
        {
            Id = -1;
            IsDefault = false;
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

            string result = null;
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        while (reader.Read()) // построчно считываем данные
                        {
                            var value = reader.GetString(3);
                            result += value;
                            break; // TODO: обработка множества значений
                        }
                    }
                }
            }

            return result;
        }

        public void SaveToDatabase(string value)
        {
            if (Id == -1)
                CreateInDatabase(value);
            else
                EditInDatabase(value);

        }

        private void CreateInDatabase(string value)
        {
            var sqlExpression =
                    $"INSERT INTO {DatabaseStrings.TextBlockTagTableName} " +
                    $"({DatabaseStrings.TextBlockKeyColumnName}, {DatabaseStrings.TextBlockConditionColumnName}, " +
                    $"{DatabaseStrings.TextBlockValueColumnName}) "
                    + "VALUES (@key, @condition, @value)";

            var rowIdSqlExpression = "SELECT last_insert_rowid() as id";
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", Key));
                command.Parameters.Add(new SqliteParameter("@condition", ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.ExecuteNonQuery();

                var rowIdCommand = new SqliteCommand(rowIdSqlExpression, connection);
                Id = Convert.ToInt32(rowIdCommand.ExecuteScalar());
            }
        }

        private void EditInDatabase(string value)
        {
            if (!HasId)
            {
                throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.UpdateError,
                    "Данный тег не найден в БД");
            }

            var sqlExpression =
                $"UPDATE {DatabaseStrings.TextBlockTagTableName} " +
                $"SET {DatabaseStrings.TextBlockKeyColumnName} = @key, " +
                $"{DatabaseStrings.TextBlockConditionColumnName} = @condition, " +
                $"{DatabaseStrings.TextBlockValueColumnName} = @value " +
                $"WHERE `{DatabaseStrings.TextBlockIdColumnName}` = @index";
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", Key));
                command.Parameters.Add(new SqliteParameter("@condition", ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.Parameters.Add(new SqliteParameter("@index", Id));
                command.ExecuteNonQuery();
            }
        }
        
        public void RemoveFromDatabase()
        {
            if (!HasId)
            {
                throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.DeleteError,
                    "Данный тег не найден в БД");
            }

            var sqlExpression =
                $"DELETE FROM {DatabaseStrings.TextBlockTagTableName} " +
                $"WHERE `{DatabaseStrings.TextBlockIdColumnName}` = @index";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@index", Id));
                command.ExecuteNonQuery();
            }
        }

        public void UpdateFormDatabase()
        {
            if (!HasId)
            {
                throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.UpdateError,
                    "Данный тег не имеет ID");
            }

            var sqlExpression =
                $"SELECT {DatabaseStrings.TextBlockIsDefaultColumnName}, " +
                $"{DatabaseStrings.TextBlockConditionColumnName} " +
                $"FROM {DatabaseStrings.TextBlockTagTableName} "
                + $"WHERE `{DatabaseStrings.TextBlockIdColumnName}` = @index";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@index", Id));
                using (var reader = command.ExecuteReader())
                {
                    if (reader.HasRows) // если есть данные
                    {
                        if (!reader.Read()) return;
                        var isDefault = reader.GetBoolean(0);
                        var xml = reader.GetString(1);
                            
                        // Мб надо что-то еще добавить...
                        var tag = FromDatabaseData(xml, Id, isDefault);
                        IsDefault = tag.IsDefault;
                        Conditions = tag.Conditions;
                        Key = tag.Key;
                        Active = tag.Active;
                        RegularEx = tag.RegularEx;
                    }
                    else
                    {
                        throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.UpdateError,
                            "Тег с данным ID не найден в БД");
                    }
                }
            }
        }

        public void SetIsDefaultState(bool state)
        {
            if (!HasId)
            {
                throw new SyllabusDatabaseException(SyllabusDatabaseErrorType.UpdateError,
                    "Данный тег не найден в БД");
            }

            var sqlExpression =
                $"UPDATE {DatabaseStrings.TextBlockTagTableName} " +
                $"SET {DatabaseStrings.TextBlockIsDefaultColumnName} = @isDefault " +
                $"WHERE `{DatabaseStrings.TextBlockIdColumnName}` = @index";
            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@isDefault", Convert.ToInt32(state)));
                command.Parameters.Add(new SqliteParameter("@index", Id));
                command.ExecuteNonQuery();
                IsDefault = state;
            }
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

        public static TextBlockTag FromDatabaseData(string xml, int id = -1, bool isDefault = false)
        {
            var serializer = new XmlSerializer(typeof(TextBlockTag));
            using (TextReader reader = new StringReader(xml))
            {
                var tag = (TextBlockTag) serializer.Deserialize(reader);
                tag.Id = id;
                tag.IsDefault = isDefault;
                return tag;
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

        public string ConditionsToGuiString()
        {
            var result = "";
            if (Conditions.Length == 0)
                return result;

            for (var i = 0; i < Conditions.Length-1; i++)
            {
                result += $"<{Conditions[i].TagName}> = `{Conditions[i].Condition}`, ";
            }
            result += $"<{Conditions[Conditions.Length-1].TagName}> = `{Conditions[Conditions.Length-1].Condition}`.";

            return result;
        }

        public TextBlockTag[] Split()
        {
            var maxConditionLength = Conditions.Max(condition => condition.Length);
            var minDelimeteredConditionLength =
                Conditions.Min(condition => condition.Delimiter != null ? condition.Length : maxConditionLength);
            Debug.Assert(minDelimeteredConditionLength == maxConditionLength);

            var splittedConditions = new List<TextBlockCondition[]>();
            foreach (var condition in Conditions)
                splittedConditions.Add(condition.Split());


            var splittedBlockTags = new TextBlockTag[maxConditionLength];
            for (var i = 0; i < maxConditionLength; i++)
            {
                var conditions = new TextBlockCondition[Conditions.Length];
                for (var j = 0; j < conditions.Length; j++)
                    conditions[j] = splittedConditions[j][splittedConditions[j].Length > 1 ? i : 0];
                splittedBlockTags[i] = new TextBlockTag(
                    Key,
                    conditions,
                    Description
                );
            }

            return splittedBlockTags;
        }

        public static List<TextBlockTag> GetAllTextBlockTags()
        {
            var sqlExpression =
                $"SELECT `{DatabaseStrings.TextBlockIdColumnName}`, {DatabaseStrings.TextBlockConditionColumnName}, " +
                $"{DatabaseStrings.TextBlockIsDefaultColumnName} " +
                $"FROM {DatabaseStrings.TextBlockTagTableName}";

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
                            var id = reader.GetInt32(0);
                            var xml = reader.GetString(1);
                            var isDefault = reader.GetBoolean(2);
                            result.Add(FromDatabaseData(xml, id, isDefault));
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
                        while (reader.Read()) // построчно считываем данные
                        {
                            var value = reader.GetString(3);
                            result += value;
                            break; // TODO: обработка множества значений
                        }
                    }
                }
            }

            return result;
        }

        public static void SetDefaultValue(string tagKey, string value, string description="")
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
                var tempCondition = new TextBlockTag(tagKey, Array.Empty<TextBlockCondition>(), description);
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", tagKey));
                command.Parameters.Add(new SqliteParameter("@condition", tempCondition.ToXml()));
                command.Parameters.Add(new SqliteParameter("@value", value));
                command.Parameters.Add(new SqliteParameter("@isDefault", 1));
                command.ExecuteNonQuery();
            }
        }

        public static void RemoveTagFromDatabase(string tagKey)
        {
            var sqlExpression =
                $"DELETE FROM {DatabaseStrings.TextBlockTagTableName} " +
                $"WHERE {DatabaseStrings.TextBlockKeyColumnName} = @key";

            using (var connection = new SqliteConnection(DatabaseStrings.ConnectionString))
            {
                connection.Open();
                var command = new SqliteCommand(sqlExpression, connection);
                command.Parameters.Add(new SqliteParameter("@key", tagKey));
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

        public string Delimiter // TODO: remove
        {
            get => _delimiter;
            set => _delimiter = value != "" ? value : "\n";
        }

        public int Length => Subconditions.Length;

        public string[] Subconditions => Delimiter == null
            ? new[] {Condition}
            : Condition.Split(new[] {Delimiter}, StringSplitOptions.None);

        public TextBlockCondition[] Split()
        {
            var subconditions = Subconditions;
            var splittedConditions = new TextBlockCondition[subconditions.Length];
            for (var i = 0; i < subconditions.Length; i++)
                splittedConditions[i] =
                    new TextBlockCondition(
                        TagName,
                        subconditions[i],
                        Delimiter
                    );

            return splittedConditions;
        }
    }
}