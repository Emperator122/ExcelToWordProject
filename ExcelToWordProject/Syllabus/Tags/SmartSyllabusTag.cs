// ReSharper disable CommentTypo
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using ExcelToWordProject.Models;
using ExcelToWordProject.Utils;

namespace ExcelToWordProject.Syllabus.Tags
{
    public class SmartSyllabusTag : BaseSyllabusTag
    {
        public int ColumnIndex; // номер столбца, в котором будут искаться значения
        public SmartTagType Type;

        public SmartSyllabusTag()
        {
        }

        public SmartSyllabusTag(int columnIndex, string key, string listName, SmartTagType type,
            string description = "") : base(key, listName, description)
        {
            ColumnIndex = columnIndex;
            Type = type;
        }

        public override string GetValue(Module module = null, List<Content> contentList = null,
            DataSet excelData = null)
        {
            return ExtractDataFromModule(this, module, contentList, module.Properties);
        }

        public static string ExtractDataFromModule(SmartSyllabusTag tag, Module module,
            List<Content> contentList = null, ModuleProperties properties = null)
        {
            List<int> tempList;
            switch (tag.Type)
            {
                case SmartTagType.ModuleName:
                    return module.Name;

                case SmartTagType.ModuleIndex:
                    return module.Index;

                case SmartTagType.Content:
                    var contentStr = "";
                    for (var i = 0; i < contentList.Count(); i++)
                    {
                        var content = contentList[i];
                        contentStr += i == contentList.Count() - 1 ? content.Value + "" : content.Value + "\n";
                    }

                    return contentStr;

                case SmartTagType.ModuleContentIndexes:
                case SmartTagType.ContentIndex:
                    var contentIndexesStr = "";
                    for (var i = 0; i < contentList.Count(); i++)
                    {
                        var content = contentList[i];
                        contentIndexesStr += i == contentList.Count() - 1 ? content.Index + "" : content.Index + "\n";
                    }

                    return contentIndexesStr;

                case SmartTagType.ExtendedContent:
                    var extContentStr = "";
                    for (var i = 0; i < contentList.Count(); i++)
                    {
                        var content = contentList[i];
                        extContentStr +=
                            i == contentList.Count() - 1
                                ? content.Value + " (" + content.Index + ")"
                                : content.Value + " (" + content.Index + ")\n";
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
                    var controlString = "";
                    for (var i = 0; i < properties.Control.Count(); i++)
                    {
                        var controlForm = properties.Control[i];
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
                    return properties.CreditUnits;

                case SmartTagType.TotalLecturesHours:
                    return properties.TotalLecturesHours.ToString();

                case SmartTagType.TotalPracticalLessonsHours:
                    return properties.TotalPracticalLessonsHours.ToString();

                case SmartTagType.TotalLaboratoryLessonsHours:
                    return properties.TotalLaboratoryLessonsHours.ToString();

                case SmartTagType.TotalHoursByPlan:
                    return properties.TotalHoursByPlan.ToString();

                case SmartTagType.TotalControlHours:
                    return properties.TotalControlHours.ToString();

                case SmartTagType.TotalIndependentWorkHours:
                    return properties.TotalIndependentWorkHours.ToString();

                case SmartTagType.Years:
                    return OtherUtils.ListToDelimiteredString("/", "", properties.Years);

                case SmartTagType.Semesters:
                    return OtherUtils.ListToDelimiteredString("/", "", properties.Semesters);

                case SmartTagType.TotalLessons:
                    return properties.TotalLessonsHours.ToString();

                case SmartTagType.LecturesHoursBySemesters:
                    if (properties.LecturesHoursBySemesters.Count(el => el == 0) ==
                        properties.LecturesHoursBySemesters.Count)
                        return "-";

                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.LecturesHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);

                case SmartTagType.PracticalLessonsHoursBySemesters:
                    if (properties.PracticalLessonsHoursBySemesters.Count(el => el == 0) ==
                        properties.PracticalLessonsHoursBySemesters.Count)
                        return "-";

                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.PracticalLessonsHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);

                case SmartTagType.LaboratoryLessonsHoursBySemesters:
                    if (properties.LaboratoryLessonsHoursBySemesters.Count(el => el == 0) ==
                        properties.LaboratoryLessonsHoursBySemesters.Count)
                        return "-";

                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.LaboratoryLessonsHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);

                case SmartTagType.IndependentWorkHoursBySemesters:
                    if (properties.IndependentWorkHoursBySemesters.Count(el => el == 0) ==
                        properties.IndependentWorkHoursBySemesters.Count)
                        return "-";

                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.IndependentWorkHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);

                case SmartTagType.ControlHoursBySemesters:
                    if (properties.ControlHoursBySemesters.Count(el => el == 0) ==
                        properties.ControlHoursBySemesters.Count)
                        return "-";

                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.ControlHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);

                case SmartTagType.TotalLessonsBySemesters:
                    // Если за семестр не было аудиторных занятий
                    // то пропуск
                    if (properties.TotalLessonsHoursBySemesters.Count(el => el == 0) ==
                        properties.TotalLessonsHoursBySemesters.Count)
                        return "-";

                    // Иначе выводим инфу
                    tempList = new List<int>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        tempList.Add(properties.TotalLessonsHoursBySemesters[semesterNumber - 1]);
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", tempList);


                case SmartTagType.isCreditBySemesters:
                    if ((properties.ControlFormsBySemesters[ControlForm.Credit]?.Count ?? 0) == 0)
                        return "-";

                    var isCreditBySemesters = new List<string>();
                    properties.Semesters.ForEach(semesterNumber =>
                    {
                        if (properties.ControlFormsBySemesters[ControlForm.Credit]?.Contains(semesterNumber) == true)
                            isCreditBySemesters.Add("+");
                        else
                            isCreditBySemesters.Add("-");
                    });
                    return OtherUtils.ListToDelimiteredString("/", "", isCreditBySemesters);

                case SmartTagType.isCourseWork:
                    return properties.isCourseWork ? "+" : "-";

                case SmartTagType.DepartmentName:
                    return properties.DepartmentName;
            }

            return "";
        }
    }

    // Зачетные единицы, Номер блока, Имя блока, Имя модуля, Компетенции, Форма контроля, Индекс модуля...
    public enum SmartTagType
    {
        None,
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

        TotalLecturesHours,
        TotalPracticalLessonsHours,
        TotalLaboratoryLessonsHours,
        TotalIndependentWorkHours,
        TotalControlHours,

        LecturesHoursBySemesters,
        PracticalLessonsHoursBySemesters,
        LaboratoryLessonsHoursBySemesters,
        IndependentWorkHoursBySemesters,
        ControlHoursBySemesters,
        TotalLessonsBySemesters,
        isCreditBySemesters,

        TotalHoursByPlan,
        TotalLessons,
        isCourseWork,
        DepartmentName,
        ModuleContentIndexes
    }
}