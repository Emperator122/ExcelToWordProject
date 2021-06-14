using ExcelToWordProject.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToWordProject
{
    public class Module
    {
        public string Index;
        public string Name;
        public string[] ContentIndexes;

        public Module(string index, string name, string сontentIndexesStr, char contentDelimiter = ';')
        {
            Index = index;
            Name = name;
            ContentIndexes = сontentIndexesStr.Split(contentDelimiter);
            for (int i = 0; i < ContentIndexes.Length; i++)
            {
                ContentIndexes[i] = ContentIndexes[i].Trim();
            }
        }

        public ModuleProperties ParseProperties(DataSet excelData, SyllabusParameters parameters)
        {
            ModuleProperties properties = new ModuleProperties(); // Результирующая переменная

            // Для начала найдем Index модуля на листе "План"
            // По дороге захватим имя блока и части
            // При этом номер части может быть не всегда (см. ФТД.Факультативы)
            // Поиск начнется с 3 индекса для скипа хедеров
            
            // Код работает на вере, что в 0 столбце будут либо плюсики с минусиками, либо имена блоков, либо null

            int rowIndex = 0;

            SmartSyllabusTag blockNameTag =
                parameters.Tags.Find(
                    tag => tag is SmartSyllabusTag && (tag as SmartSyllabusTag).Type == SmartTagType.BlockName) as SmartSyllabusTag;
            SmartSyllabusTag partNameTag =
                parameters.Tags.Find(
                    tag => tag is SmartSyllabusTag && (tag as SmartSyllabusTag).Type == SmartTagType.PartName) as SmartSyllabusTag;
            var rows = excelData.Tables[blockNameTag.ListName].Rows;

            // Маркеры, которые могут быть в 0 столбце
            string[] moduleMarkers = { "+", "-" };
            string[] skipMarkers = { null, ""};
            for (rowIndex = 3; rowIndex < rows.Count; rowIndex++)
            {
                string val = rows[rowIndex][0] as string;
                if (moduleMarkers.Contains(val)) // Если строка содержит модуль
                {
                    if(rows[rowIndex][1] as string == Index) // если это наш модуль, то останавливаем процесс
                        break;
                }
                else
                {
                    // Содержит ли данная строка имя части
                    string maybePart = rows[rowIndex][partNameTag.ColumnIndex] as string ?? "";
                    if (!skipMarkers.Contains(maybePart.Trim()))
                    {
                        if (maybePart.ToLower().Contains("часть"))
                        {
                            properties.PartName = maybePart;
                        }
                        else
                        {
                            // Содержит ли данная строка имя блока
                            string maybeBlock = rows[rowIndex][blockNameTag.ColumnIndex] as string ?? "";
                            if (!skipMarkers.Contains(maybeBlock.Trim()))
                            {
                                properties.PartName = "";
                                properties.BlockName = maybeBlock;
                            }
                        }
                    }
                }
            }

            // если мы не нашли нужную строку
            if (rowIndex == rows.Count)
                return new ModuleProperties("error", "error", new List<ControlForm>(), -1);

            // Тянем доп данные

            // Номер блока
            string temp = Regex.Replace(properties.BlockName, @"[^\d]+", "");
            properties.BlockNumber = (temp != "") ? Convert.ToInt32(temp) : -1;

            // Форма контроля
            SmartSyllabusTag controlTag =
                parameters.Tags.Find(
                    tag => tag is SmartSyllabusTag && (tag as SmartSyllabusTag).Type == SmartTagType.Control) as SmartSyllabusTag; 
            if (rows[rowIndex][controlTag.ColumnIndex] as string != null && rows[rowIndex][controlTag.ColumnIndex] as string != "")
            {
                properties.Control.Add(ControlForm.Exam);
            }
            if (rows[rowIndex][controlTag.ColumnIndex+2] as string != null && rows[rowIndex][controlTag.ColumnIndex+2] as string != "")
            {
                properties.Control.Add(ControlForm.GradedCredit);
            }
            if (rows[rowIndex][controlTag.ColumnIndex + 1] as string != null && rows[rowIndex][controlTag.ColumnIndex + 1] as string != "")
            {
                properties.Control.Add(ControlForm.Credit);
            }

            // Зачетные единицы
            SmartSyllabusTag сreditUnitsTag =
                parameters.Tags.Find(
                    tag => tag is SmartSyllabusTag && (tag as SmartSyllabusTag).Type == SmartTagType.CreditUnits) as SmartSyllabusTag;
            properties.CreditUnits = -1;
            if (rows[rowIndex][сreditUnitsTag.ColumnIndex] as string != "")
            {
                properties.CreditUnits = Convert.ToInt32(rows[rowIndex][сreditUnitsTag.ColumnIndex] as string ?? "-1");
            }


            return properties;
        }
        public List<Content> ParseContentList(DataSet excelData, SyllabusParameters parameters)
        {
            SmartSyllabusTag tag =
                parameters.Tags.Find(
                    tag_ => tag_ is SmartSyllabusTag && (tag_ as SmartSyllabusTag).Type == SmartTagType.Content) as SmartSyllabusTag;
            List<Content> contentList = new List<Content>();

            var rows = excelData.Tables[tag.ListName].Rows;

            for (int i = 2; i < rows.Count; i++)
            {

                // Если справа мы встретили не пустой столбец "Тип"
                if ((rows[i][tag.ColumnIndex + 1] as string ?? "").Trim() != "")
                    foreach (string contentIndex in ContentIndexes) // Проверяем все индексы
                    {
                        if ((rows[i][tag.ColumnIndex - 2] as string) == contentIndex) //на совпадение
                        {
                            contentList.Add(new Content(contentIndex, rows[i][tag.ColumnIndex] as string));
                            break;
                        }
                    }
            }
          
            return contentList;
        }

        public override string ToString()
        {
            return Index + " " + Name;
        }
    }
}
