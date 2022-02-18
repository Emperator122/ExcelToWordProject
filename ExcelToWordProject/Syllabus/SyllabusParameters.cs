using ExcelToWordProject.Models;
using ExcelToWordProject.Syllabus.Tags;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelToWordProject.Syllabus
{
    public class SyllabusParameters
    {
        [System.Xml.Serialization.XmlIgnore]
        public bool HasActiveSmartTags
        {
            get { return Tags.FindIndex(el => el is SmartSyllabusTag && el.Active) != -1; }
        }

        public string ModulesContentListName;
        public string PlanListName;
        public int PlanListHeaderRowIndex;

        [System.Xml.Serialization.XmlIgnore]
        public Dictionary<string, string> planListHeaderNames;
        public List<TempDictionaryItem> tempPlanListHeaderNames; // для сериализации словаря

        public int[] ModulesYears;

        [System.Xml.Serialization.XmlElementAttribute("SmartSyllabusTag", typeof(SmartSyllabusTag))]
        [System.Xml.Serialization.XmlElementAttribute("DefaultSyllabusTag", typeof(DefaultSyllabusTag))]
        public List<BaseSyllabusTag> Tags;

        [System.Xml.Serialization.XmlIgnore]
        public List<TextBlockTag> TextBlockTags { get => TextBlockTag.GetAllTextBlockTags(); }
        
        public IEnumerable<IGrouping<string, TextBlockTag>> GroupedTextBlockTags =>
            TextBlockTags
                .GroupBy(tag => tag.ToXml());

        /// <summary>
        /// Конструктор без параметров. В основном нужен для сериализации.
        /// </summary>
        public SyllabusParameters() { }

        /// <summary>
        /// Конструктор с набором стандартных параметров
        /// </summary>
        /// <param name="fillWithValues">Заполнить стандартными параметрами</param>
        public SyllabusParameters(bool fillWithValues)
        {
            if (!fillWithValues) return;

            ModulesYears = new int[0];

            // Инициализация базового набора параметров
            PlanListName = "План";
            ModulesContentListName = "Компетенции";

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
            planListHeaderNames["DepartmentName"] = "Наименование";
            planListHeaderNames["Competitions"] = "Компетенции";
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

                new SmartSyllabusTag(-1, "ModuleContentIndexes", PlanListName, SmartTagType.ModuleContentIndexes, "Совпадает с ContentIndex. Играет служебную роль."),

                new SmartSyllabusTag(3, "Control", PlanListName, SmartTagType.Control,
                "Форма контроля. \r\nНапр.: Экзамен; Зачет; Зачет с оценкой; Экзамен, зачет; ...\r\nТочки в конце нет!"),

                new SmartSyllabusTag(2, "ModuleName", PlanListName, SmartTagType.ModuleName,
                "Имя дисциплины. \r\nНапр.: Математика"),

                new SmartSyllabusTag(1, "ModuleIndex", PlanListName, SmartTagType.ModuleIndex,
                "Индекс дисциплины. \r\nНапр.: Б1.Б.05"),

                new SmartSyllabusTag(-1, "Years", PlanListName, SmartTagType.Years, "Курсы, на которых преподается дисциплина.\r\n" +
                "Напр.: 1, 2, 3"),

                new SmartSyllabusTag(-1, "Semesters", PlanListName, SmartTagType.Semesters, "Семестры, на которых преподается дисциплина.\r\n" +
                "Напр.: 1, 2, 3"),

                new SmartSyllabusTag(-1, "TotalLecturesHours", PlanListName, SmartTagType.TotalLecturesHours, "Общее количество акад. часов на лекции.\r\n" +
                "Напр.: 200"),

                new SmartSyllabusTag(-1, "TotalPracticalLessonsHours", PlanListName, SmartTagType.TotalPracticalLessonsHours, "Общее количество акад. часов на практики.\r\n" +
                "Напр.: 150"),

                new SmartSyllabusTag(-1, "TotalLaboratoryLessonsHours", PlanListName, SmartTagType.TotalLaboratoryLessonsHours, "Общее количество акад. часов на лабораторные.\r\n" +
                "Напр.: 140"),

                new SmartSyllabusTag(-1, "TotalIndependentWorkHours", PlanListName, SmartTagType.TotalIndependentWorkHours, "Общее количество акад. часов на самостоятельную работу.\r\n" +
                "Напр.: 100"),

                new SmartSyllabusTag(-1, "TotalControlHours", PlanListName, SmartTagType.TotalControlHours, "Общее количество акад. часов на экзамен\r\n" +
                "Напр.: 200"),

                new SmartSyllabusTag(-1, "TotalHoursByPlan", PlanListName, SmartTagType.TotalHoursByPlan, "Общее количество акад. часов.\r\n" +
                "Напр.: 700"),

                new SmartSyllabusTag(-1, "TotalLessons", PlanListName, SmartTagType.TotalLessons,
                "Итого аудиторных занятий. Сумма лаб, лекций и практик.\r\n" +
                "Напр.: 200"),


                new SmartSyllabusTag(-1, "LecturesHoursBySemesters", PlanListName, SmartTagType.LecturesHoursBySemesters,
                "Количество акад. часов на лекции по семестрам.\r\n" +
                "Напр.: 50/60/17"),

                new SmartSyllabusTag(-1, "PracticalLessonsHoursBySemesters", PlanListName, SmartTagType.PracticalLessonsHoursBySemesters,
                "Количество акад. часов на практики по семестрам.\r\n" +
                "Напр.: 14/25/30"),

                new SmartSyllabusTag(-1, "LaboratoryLessonsHoursBySemesters", PlanListName, SmartTagType.LaboratoryLessonsHoursBySemesters,
                "Количество акад. часов на лабораторные по семестрам.\r\n" +
                "Напр.: 60/80/40"),

                new SmartSyllabusTag(-1, "IndependentWorkHoursBySemesters", PlanListName, SmartTagType.IndependentWorkHoursBySemesters,
                "Количество акад. часов на самостоятельную работу по семестрам.\r\n" +
                "Напр.: 40/70/0"),

                new SmartSyllabusTag(-1, "ControlHoursBySemesters", PlanListName, SmartTagType.ControlHoursBySemesters,
                "Количество акад. часов контроль по семестрам.\r\n" +
                "Напр.: 20/0/60"),

                new SmartSyllabusTag(-1, "TotalLessonsBySemesters", PlanListName, SmartTagType.TotalLessonsBySemesters, "Итого аудиторных занятий по семестрам. " +
                "Сумма лаб, лекций и практик по семестрам.\r\n" +
                "Напр.: 100/110/120"),

                new SmartSyllabusTag(-1, "isCreditBySemesters", PlanListName, SmartTagType.isCreditBySemesters, "Перечисление значений: будет ли в данном семестре " +
                "зачет. Семестры берутся из списка семестров, в которые ведется предмет.\r\nНапр.:+/+/-/+"),

                new SmartSyllabusTag(6, "isCourseWork", PlanListName, SmartTagType.isCourseWork, "+ если по предмету есть курсова, иначе -"),

                new SmartSyllabusTag(-1, "DepartmentName", PlanListName, SmartTagType.DepartmentName, "Имя кафедры.\r\nНапр.: Экономики"),



                // Обычные теги
                new DefaultSyllabusTag(15, 1, "DirectionCode", "Титул", "Номер направления.\r\nНапр.:09.03.01"),

                new DefaultSyllabusTag(17, 1, "DirectionName", "Титул", "Имя направления. \r\nНапр.: Прикладная математика и информатика", true,
                            new RegExpData(){ Expression = @"\d ((.*)(?=[ \n](?=Программа|Профиль))|(.*))",
                                GroupIndex = 1, RegexOptions = RegexOptions.Singleline }),

                new DefaultSyllabusTag(17, 1, "ProgramValue", "Титул", "Название программы/профиля.", true,
                            new RegExpData(){ Expression = @"((Программа|Профиль).*)",
                                GroupIndex = 1, RegexOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase }),

                new DefaultSyllabusTag(12, 0, "ProtocolInfo", "Титул", "План одобрен Ученым советом вуза. Протокол №... \r\nНапр.: № 12 от 29.02.2020", true,
                            new RegExpData(){ Expression = @"Протокол (.*)",
                                GroupIndex = 1, RegexOptions = RegexOptions.Singleline }),

                new DefaultSyllabusTag(30, 19, "EducationalStandard", "Титул", "Образовательный стандарт (ФГОС). \r\nНапр.:  № 12 от 10.01.2018"),

                new DefaultSyllabusTag(30, 0, "StudyForm", "Титул", "Форма обучения.\r\nНапр.: Очная", true,
                            new RegExpData(){ Expression = @".*: (.*)",
                                GroupIndex = 1, RegexOptions = RegexOptions.Singleline }),

                new DefaultSyllabusTag(28, 0, "Qualification", "Титул", "Квалификация.\r\nНапр.: Бакалавр", true,
                            new RegExpData(){ Expression = @"Квалификация: (.*)",
                                GroupIndex = 1, RegexOptions = RegexOptions.Singleline | RegexOptions.IgnoreCase }),
            };
        }

        /// <summary>
        /// Отключить все смарт теги
        /// </summary>
        public void DisableSmartTags()
        {
            Tags.ForEach(el =>
            {
                if (el is SmartSyllabusTag)
                    el.Active = false;
            });
        }
    }
}