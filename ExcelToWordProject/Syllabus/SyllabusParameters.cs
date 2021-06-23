using ExcelToWordProject.Models;
using ExcelToWordProject.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToWordProject.Syllabus
{
    public class SyllabusParameters
    {
        public string ModulesContentListName;
        public string PlanListName;
        public string ModulesListName;
        public int PlanListHeaderRowIndex;

        [System.Xml.Serialization.XmlIgnore]
        public Dictionary<string, string> planListHeaderNames;
        public List<TempDictionaryItem> tempPlanListHeaderNames; // для сериализации словаря


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

            // Стоп слова при парсинге имен дисциплин
            ModuleNameStopWords = new string[] { "дисциплины ", " часть", "факультативы", "наименование", "практики", "практика", "часть,",
                "государственная итоговая аттестация"};


            // Индекс строки с хедером на листе "План"
            PlanListHeaderRowIndex = 2;

            // хедеры с листа "План"
            planListHeaderNames = new Dictionary<string, string>();
            planListHeaderNames["LecturesHoursHeaderName"] = "Лек";
            planListHeaderNames["LaboratoryLessonsHoursHeaderName"] = "Лаб";
            planListHeaderNames["PracticalLessonsHoursHeaderName"] = "Пр";
            planListHeaderNames["ControlHoursHeaderName"] = "Конт роль";
            planListHeaderNames["IndependentWorkHoursHeaderName"] = "СР";
            planListHeaderNames["TotalHoursByPlanHeaderName"] = "По плану";
            planListHeaderNames["SemesterCountingCreditUnitsPlanHeaderName"] = "з.е.";
            // заполняем временный массив для сериализации
            tempPlanListHeaderNames = new List<TempDictionaryItem>(planListHeaderNames.Select(kv 
                => new TempDictionaryItem() { Name = kv.Key, Value = kv.Value }).ToArray());


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
                "Расширенный список компетенций (с идентификатором в конце).\r\n" +
                "Для корректного отображения необходимо, чтобы тег был единственным элементом в абзаце (без каких-либо других символов в абзаце кроме тега).\r\n" +
                "Напр.:\r\nспособностью использовать основы философских знаний для формирования мировоззренческой позиции (ОК-1);\r\n" +
                "способностью анализировать основные этапы и закономерности исторического развития общества для формирования гражданской позиции (ОК-2)."),

                new SmartSyllabusTag(2, "ContentIndex", ModulesContentListName, SmartTagType.ContentIndex, "Идентификатор компетенции. \r\nПример: ОК-2"),

                new SmartSyllabusTag(3, "Control", PlanListName, SmartTagType.Control,
                "Форма контроля. \r\nНапр.: Экзамен; Зачет; Зачет с оценкой; Экзамент, зачет; ...\r\nТочки в конце нет!"),

                new SmartSyllabusTag(4, "ModuleName", ModulesListName, SmartTagType.ModuleName,
                "Имя дисциплины. \r\nНапр.: Математика"),

                new SmartSyllabusTag(3, "ModuleIndex", ModulesListName, SmartTagType.ModuleIndex,
                "Индекс дисциплины. \r\nНапр.: Б1.Б.05"),

                new SmartSyllabusTag(-1, "Years", PlanListName, SmartTagType.Years, "Курсы, на которых преподается дисциплина.\r\n" +
                "Напр.: 1, 2, 3"),

                new SmartSyllabusTag(-1, "Semesters", PlanListName, SmartTagType.Semesters, "Семестры, на которых преподается дисциплина.\r\n" +
                "Напр.: 1, 2, 3"),

                new SmartSyllabusTag(-1, "LecturesHours", PlanListName, SmartTagType.LecturesHours, "Общее количество акад. часов на лекции.\r\n" +
                "Напр.: 200"),

                new SmartSyllabusTag(-1, "PracticalLessonsHours", PlanListName, SmartTagType.PracticalLessonsHours, "Общее количество акад. часов на практики.\r\n" +
                "Напр.: 150"),

                new SmartSyllabusTag(-1, "LaboratoryLessonsHours", PlanListName, SmartTagType.LaboratoryLessonsHours, "Общее количество акад. часов на лабораторные.\r\n" +
                "Напр.: 140"),

                new SmartSyllabusTag(-1, "IndependentWorkHours", PlanListName, SmartTagType.IndependentWorkHours, "Общее количество акад. часов на самостоятельную работу.\r\n" +
                "Напр.: 100"),

                new SmartSyllabusTag(-1, "ControlHours", PlanListName, SmartTagType.ControlHours, "Общее количество акад. часов контроль\r\n" +
                "Напр.: 200"),

                new SmartSyllabusTag(-1, "TotalHoursByPlan", PlanListName, SmartTagType.TotalHoursByPlan, "Общее количество акад. часов.\r\n" +
                "Напр.: 700"),

                new SmartSyllabusTag(-1, "TotalLessons", PlanListName, SmartTagType.TotalLessons, "Итого аудиторных занятий. Сумма лаб, лекций и практик.\r\n" +
                "Напр.: 200"),

                // Обычные теги
                new DefaultSyllabusTag(15, 1, "DirectionCode", "Титул", "Номер направления.\r\nНапр.:09.03.01"),

                new DefaultSyllabusTag(17, 1, "DirectionName", "Титул", "Имя направления. \r\nНапр.: Прикладная математика и информатика", true, 
                            new RegExpData(){ Expression = @"\d ((.*)(?= Программа)|(.*))",
                                GroupIndex = 1, RegexOptions = System.Text.RegularExpressions.RegexOptions.Multiline }),

                new DefaultSyllabusTag(12, 0, "ProtocolInfo", "Титул", "План одобрен Ученым советом вуза. Протокол №... \r\nНапр.: № 12 от 29.02.2020", true,
                            new RegExpData(){ Expression = @"Протокол (.*)",
                                GroupIndex = 1, RegexOptions = System.Text.RegularExpressions.RegexOptions.Multiline }),

                new DefaultSyllabusTag(30, 19, "EducationalStandard", "Титул", "Образовательный стандарт (ФГОС). \r\nНапр.:  № 12 от 10.01.2018"),

                new DefaultSyllabusTag(30, 0, "StudyForm", "Титул", "Форма обучения.\r\nНапр.: Очная", true,
                            new RegExpData(){ Expression = @".*: (.*)",
                                GroupIndex = 1, RegexOptions = System.Text.RegularExpressions.RegexOptions.Multiline }),
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

        public override string GetValue(Module module = null, List<Content> contentList = null, ModuleProperties properties = null, DataSet excelData = null)
        {
            return ExtractDataFromModule(this, module, contentList, properties);
        }

        public static string ExtractDataFromModule(SmartSyllabusTag tag, Module module, List<Content> contentList = null, ModuleProperties properties = null)
        {
            switch (tag.Type)
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
                        contentStr += (i == contentList.Count() - 1) ? content.Value + "" : content.Value + "\n";
                    }
                    return contentStr;
                case SmartTagType.ContentIndex:
                    string contentIndexesStr = "";
                    for (int i = 0; i < contentList.Count(); i++)
                    {
                        Content content = contentList[i];
                        contentIndexesStr += (i == contentList.Count() - 1) ? content.Index + "" : content.Index + "\n";
                    }
                    return contentIndexesStr;
                case SmartTagType.ExtendedContent:
                    string extContentStr = "";
                    for (int i = 0; i < contentList.Count(); i++)
                    {
                        Content content = contentList[i];
                        extContentStr += 
                            (i == contentList.Count() - 1) ? content.Value + " (" + content.Index + ")" : content.Value + " (" + content.Index + ")\n";
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

                case SmartTagType.LecturesHours:
                    return properties.LecturesHours.ToString();

                case SmartTagType.PracticalLessonsHours:
                    return properties.PracticalLessonsHours.ToString();

                case SmartTagType.LaboratoryLessonsHours:
                    return properties.LaboratoryLessonsHours.ToString();

                case SmartTagType.TotalHoursByPlan:
                    return properties.TotalHoursByPlan.ToString();

                case SmartTagType.ControlHours:
                    return properties.ControlHours.ToString();

                case SmartTagType.IndependentWorkHours:
                    return properties.IndependentWorkHours.ToString();

                case SmartTagType.Years:
                    return OtherUtils.ListToDelimiteredString(", ", "", properties.Years);

                case SmartTagType.Semesters:
                    return OtherUtils.ListToDelimiteredString(", ", "", properties.Semesters);

                case SmartTagType.TotalLessons:
                    return properties.TotalLessonsHours.ToString();
            }
            return "";
        }
    }

    public class DefaultSyllabusTag : BaseSyllabusTag
    {
        public int ColumnIndex;
        public int RowIndex;
        
        public DefaultSyllabusTag(int rowIndex, int columnIndex, string key, string listName, 
            string description = "", bool active = true, RegExpData regularEx = null) : base(key, listName, description, active, regularEx)
        {
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }


        public override string GetValue(Module module = null, List<Content> contentList = null, ModuleProperties properties = null, DataSet excelData = null)
        {
            try
            {
                string result = excelData.Tables[ListName].Rows[RowIndex][ColumnIndex] as string ?? "";
                if (RegularEx.Expression != "")
                {
                    try
                    {
                        var match = Regex.Match(result, RegularEx.Expression, RegularEx.RegexOptions);
                        result = match.Groups[RegularEx.GroupIndex].Value;
                    }
                    catch
                    {
                    }
                }
                return result;
            }
            catch
            {
                return "Error<EDTD>";
            }
        }

        public DefaultSyllabusTag() { }

    }

    public abstract class BaseSyllabusTag
    {
        public string Key;
        public string ListName;
        public bool Active;
        public string Description;
        public RegExpData RegularEx;

        public string Tag { get => "<" + Key + ">"; }

        public BaseSyllabusTag(string key, string listName, string description, bool active = true, RegExpData regularEx = null)
        {
            if (regularEx == null)
                RegularEx = new RegExpData() { Expression = "", GroupIndex = 0, RegexOptions = System.Text.RegularExpressions.RegexOptions.None };
            else
                RegularEx = regularEx;
            Key = key;
            ListName = listName;
            Description = description;
            Active = active;
        }

        /// <summary>
        /// Получение значения тега
        /// </summary>
        /// <param name="module">Для smart</param>
        /// <param name="contentList">Для smart</param>
        /// <param name="properties">Для smart</param>
        /// <param name="excelData">Для Default</param>
        /// <returns></returns>
        public abstract string GetValue(Module module = null, List<Content> contentList = null, ModuleProperties properties = null, DataSet excelData = null);



        public BaseSyllabusTag() { }
    }

    // Зачетные единицы, Номер блока, Имя блока, Имя модуля, Компетенции, Форма контроля, Индекс модуля...
    public enum SmartTagType
    {
        CreditUnits,
        BlockNumber,
        BlockName,
        ModuleName,
        Content,
        Control,
        ModuleIndex,
        PartName,
        ExtendedContent,
        ContentIndex,

        Years,
        Semesters,
        LecturesHours,
        PracticalLessonsHours,
        LaboratoryLessonsHours,
        IndependentWorkHours,
        ControlHours,
        TotalHoursByPlan,
        TotalLessons
    }
}
