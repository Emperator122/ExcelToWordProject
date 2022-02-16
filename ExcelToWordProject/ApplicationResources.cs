﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject
{
    internal static class DatabaseStrings
    {
        public static readonly string DatabasePath = "database/storage.db";
        public static string ConnectionString { get => "Data Source=" + DatabasePath; }

        // Константы для TextBlockTag
        public static readonly string TextBlockTagTableName = "syllabus_text_blocks";
        public static readonly string TextBlockKeyColumnName = "tag_key";
        public static readonly string TextBlockConditionColumnName = "condition_object";
        public static readonly string TextBlockValueColumnName = "value";

    }
}