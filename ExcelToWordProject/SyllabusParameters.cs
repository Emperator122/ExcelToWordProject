using ExcelToWordProject.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject
{
    public class SyllabusParameters
    {
        public string ModulesContentListName;
        public string PlanListName;
        public string ModulesListName;

        public string[] ModuleNameStopWords;

        [System.Xml.Serialization.XmlElementAttribute("SmartSyllabusTag", typeof(SmartSyllabusTag))]
        [System.Xml.Serialization.XmlElementAttribute("DefaultSyllabusTag", typeof(DefaultSyllabusTag))]
        public List<BaseSyllabusTag> Tags;

        public SyllabusParameters() { }

        public SyllabusParameters(bool fillWithValues) {
            if (!fillWithValues) return;

            // Инициализация базового набора параметров
            ModulesContentListName = "Компетенции";
            PlanListName = "План";
            ModulesListName = "Компетенции(2)";

            ModuleNameStopWords = new string[] { "дисциплины ", " часть", "факультативы", "наименование", "практики", "практика", "часть,",
                "государственная итоговая аттестация"};


            Tags = new List<BaseSyllabusTag>()
            {
                new SmartSyllabusTag(0, "BlockName", PlanListName, SmartTagType.BlockName),
                new SmartSyllabusTag(0, "BlockNumber", PlanListName, SmartTagType.BlockNumber),
                new SmartSyllabusTag(0, "PartName", PlanListName, SmartTagType.PartName),

                new SmartSyllabusTag(7, "CreditUnits", PlanListName, SmartTagType.CreditUnits),
                new SmartSyllabusTag(3, "Content", ModulesContentListName, SmartTagType.Content),
                new SmartSyllabusTag(3, "Control", PlanListName, SmartTagType.Control),
                new SmartSyllabusTag(4, "ModuleName", ModulesListName, SmartTagType.ModuleName),
                new SmartSyllabusTag(3, "ModuleIndex", ModulesListName, SmartTagType.ModuleIndex),
                new DefaultSyllabusTag(15, 1, "Code", "Титул")
            };
        }
    }


    public class SmartSyllabusTag : BaseSyllabusTag
    {
        public int ColumnIndex; // номер столбца, в котором будут искаться значения
        public SmartTagType Type;

        public SmartSyllabusTag() { }

        public SmartSyllabusTag(int columnIndex, string key, string listName, SmartTagType type) : base(key, listName)
        {
            ColumnIndex = columnIndex;
            Type = type;
        }

        public string ExtractDataFromModule(Module module, List<Content> contentList = null, ModuleProperties properties = null)
        {
            switch (Type)
            {
                case SmartTagType.ModuleName:
                    return module.Name;
                case SmartTagType.ModuleIndex:
                    return module.Index;
                case SmartTagType.Content:
                    string contentStr = "";
                    contentList.ForEach(content => contentStr += content.Value + ";\r\n");
                    return contentStr;
                case SmartTagType.BlockName:
                    return properties.BlockName;
                case SmartTagType.BlockNumber:
                    return properties.BlockNumber.ToString();
                case SmartTagType.PartName:
                    return properties.PartName;
                case SmartTagType.Control:
                    if (properties.Control == ControlForm.Exam)
                        return "экзамен";
                    else if (properties.Control == ControlForm.Credit)
                        return "зачет";
                    else if (properties.Control == ControlForm.GradedCredit)
                        return "зачет с оценкой";
                    return "error";
                case SmartTagType.CreditUnits:
                    return properties.CreditUnits.ToString();

            }
            return "";
        }
    }

    public class DefaultSyllabusTag : BaseSyllabusTag
    {
        public int ColumnIndex;
        public int RowIndex;
        
        public DefaultSyllabusTag(int rowIndex, int columnIndex, string key, string listName) : base(key, listName)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }

        public DefaultSyllabusTag() { }

    }

    public class BaseSyllabusTag
    {
        public string Key;
        public string ListName;
        public bool Active = true;
        public BaseSyllabusTag(string key, string listName)
        {
            Key = key;
            ListName = listName;
        }

        public BaseSyllabusTag() { }
    }

    // Зачетные единицы, Номер блока, Имя блока, Имя модуля, Компетенции, Форма контроля, Индекс модуля
    public enum SmartTagType { CreditUnits, BlockNumber, BlockName, ModuleName, Content, Control, ModuleIndex, PartName }
}
