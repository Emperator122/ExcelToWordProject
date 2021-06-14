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

            // Дефолтный набор тегов
            Tags = new List<BaseSyllabusTag>()
            {
                // Умные теги
                new SmartSyllabusTag(0, "BlockName", PlanListName, SmartTagType.BlockName,
                "Полное наименование блока, в который входит дисциплина. \r\nНапр.: Блок 1.Дисциплины (модули)."),
                new SmartSyllabusTag(0, "BlockNumber", PlanListName, SmartTagType.BlockNumber,
                "Номер блока, в который входит дисциплина. \r\nНапр.: Если дисциплина находится в Блоке 1, то тег будет содержать в себе значение \"1\"."),
                new SmartSyllabusTag(0, "PartName", PlanListName, SmartTagType.PartName,
                "Имя части, в которую входить дисциплина.\r\nНапр.: Базовая часть/Вариативная; может быть пустой строкой (в случае с факультативами)."),
                new SmartSyllabusTag(7, "CreditUnits", PlanListName, SmartTagType.CreditUnits, 
                "Зачетные единицы. По-умолчанию используются экспертные. Можно изменить на фактические, увеличив номер столбца на 1.\r\n Напр.: 4"),
                new SmartSyllabusTag(3, "Content", ModulesContentListName, SmartTagType.Content,
                "Список компетенций.\r\n" +
                "Для корректного отображения необходимо, чтобы тег был единственным элементом в абзаце (без каких-либо других символов в абзаце кроме тега).\r\n" +
                "Напр.:\r\nспособностью использовать основы философских знаний для формирования мировоззренческой позиции;\r\n" +
                "способностью анализировать основные этапы и закономерности исторического развития общества для формирования гражданской позиции."),
                new SmartSyllabusTag(3, "ExtendedContent", ModulesContentListName, SmartTagType.ExtendedContent,
                "Расширенный список компетенций.\r\n" +
                "Для корректного отображения необходимо, чтобы тег был единственным элементом в абзаце (без каких-либо других символов в абзаце кроме тега).\r\n" +
                "Напр.:\r\nспособностью использовать основы философских знаний для формирования мировоззренческой позиции (ОК-1);\r\n" +
                "способностью анализировать основные этапы и закономерности исторического развития общества для формирования гражданской позиции (ОК-2)."),
                new SmartSyllabusTag(3, "Control", PlanListName, SmartTagType.Control, 
                "Форма контроля. \r\nНапр.: Экзамен; Зачет; Зачет с оценкой; Экзамент, зачет; ...\r\nТочки в конце нет!"),
                new SmartSyllabusTag(4, "ModuleName", ModulesListName, SmartTagType.ModuleName,
                "Имя дисциплины. \r\nНапр.: Математика"),
                new SmartSyllabusTag(3, "ModuleIndex", ModulesListName, SmartTagType.ModuleIndex,
                "Индекс дисциплины. \r\nНапр.: Б1.Б.05"),

                // Обычные теги
                new DefaultSyllabusTag(15, 1, "Code", "Титул", "Номер направления.\r\nНапр.:09.03.01")
            };
        }
    }


    public class SmartSyllabusTag : BaseSyllabusTag
    {
        public int ColumnIndex; // номер столбца, в котором будут искаться значения
        public SmartTagType Type;

        public SmartSyllabusTag() { }

        public SmartSyllabusTag(int columnIndex, string key, string listName, SmartTagType type, string description = "") : base(key, listName, description)
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
                    for (int i = 0; i < contentList.Count(); i++)
                    {
                        Content content = contentList[i];
                        contentStr += (i == contentList.Count() - 1) ? content.Value + "." : content.Value + ";\n";
                    }
                    return contentStr;
                case SmartTagType.ExtendedContent:
                    string extContentStr = "";
                    for (int i = 0; i < contentList.Count(); i++)
                    {
                        Content content = contentList[i];
                        extContentStr += 
                            (i == contentList.Count() - 1) ? content.Value + " (" + content.Index + ")." : content.Value + " (" + content.Index + ");\n";
                    }
                    return extContentStr;
                case SmartTagType.BlockName:
                    return properties.BlockName;
                case SmartTagType.BlockNumber:
                    return properties.BlockNumber.ToString();
                case SmartTagType.PartName:
                    return properties.PartName;
                case SmartTagType.Control:

                    if (properties.Control.Count == 0)
                        return "-";
                    string controlString = "";
                    for (int i = 0; i < properties.Control.Count(); i++)
                    {
                        ControlForm controlForm = properties.Control[i];
                        if (controlForm == ControlForm.Exam)
                            controlString += "экзамен";
                        if (controlForm == ControlForm.Credit)
                            controlString += "зачет";
                        if (controlForm == ControlForm.GradedCredit)
                            controlString += "зачет с оценкой";

                        if (i != properties.Control.Count() - 1)
                            controlString += ", ";
                    }
                    return controlString;
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
        
        public DefaultSyllabusTag(int rowIndex, int columnIndex, string key, string listName, string description = "") : base(key, listName, description)
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
        public string Description;

        public string Tag { get => "<" + Key + ">"; }

        public BaseSyllabusTag(string key, string listName, string description)
        {
            Key = key;
            ListName = listName;
            Description = description;
        }

        public BaseSyllabusTag() { }
    }

    // Зачетные единицы, Номер блока, Имя блока, Имя модуля, Компетенции, Форма контроля, Индекс модуля
    public enum SmartTagType { CreditUnits, BlockNumber, BlockName, ModuleName, Content, Control, ModuleIndex, PartName, ExtendedContent}
}
