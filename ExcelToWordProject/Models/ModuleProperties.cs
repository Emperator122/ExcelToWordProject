using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToWordProject.Models
{
    public class ModuleProperties
    {
        public int TotalLessonsHours { get => TotalLecturesHours + TotalPracticalLessonsHours + TotalLaboratoryLessonsHours; } // Итого аудиторных занятий

        public List<int> TotalLessonsHoursBySemesters
        {
            get {
                List<int> result = new List<int>(LecturesHoursBySemesters);
                for (int i = 0; i < LecturesHoursBySemesters.Count; i++)
                    result[i] += PracticalLessonsHoursBySemesters[i] + LaboratoryLessonsHoursBySemesters[i];
                return result;
            }
        }
        public List<int> Years // Курсы, на которых преподается модуль
        {
            get
            {
                List<int> years = new List<int>();
                Semesters.ForEach(semester =>
                {
                    int year = semester % 2 == 0 ? semester / 2 : semester / 2 + 1;
                    if (!years.Contains(year))
                        years.Add(year);
                });

                return years;
            }
        }

        public int TotalLecturesHours { get => LecturesHoursBySemesters.Sum(); }
        public int TotalPracticalLessonsHours { get => PracticalLessonsHoursBySemesters.Sum(); }
        public int TotalLaboratoryLessonsHours { get => LaboratoryLessonsHoursBySemesters.Sum(); }
        public int TotalIndependentWorkHours { get => IndependentWorkHoursBySemesters.Sum(); }
        public int TotalControlHours { get => ControlHoursBySemesters.Sum(); }

        public List<ControlForm> Control
        {
            get
            {
                List<ControlForm> controls = new List<ControlForm>();
                Array enumValues = Enum.GetValues(typeof(ControlForm));
                for (int i = 0; i < enumValues.Length - 1; i++)
                {
                    ControlForm controlForm = (ControlForm)enumValues.GetValue(i);
                    if (ControlFormsBySemesters.ContainsKey(controlForm) && ControlFormsBySemesters[controlForm]?.Count > 0)
                        controls.Add(controlForm);
                }
                if (controls.Count == 0)
                    controls.Add(ControlForm.Error);
                return controls;
            }
        }

        public int BlockNumber = -1;
        public string BlockName = "";
        public string PartName = "";
        public string CreditUnits = "-1";
        public string DepartmentName = "";

        public List<int> LecturesHoursBySemesters = new List<int>();
        public List<int> PracticalLessonsHoursBySemesters = new List<int>();
        public List<int> LaboratoryLessonsHoursBySemesters = new List<int>();
        public List<int> IndependentWorkHoursBySemesters = new List<int>();
        public List<int> ControlHoursBySemesters = new List<int>();
        public List<int> Semesters = new List<int>();

        // Соответсвие между формой контроля и перечислением семестров
        public Dictionary<ControlForm, List<int>> ControlFormsBySemesters 
                                    = new Dictionary<ControlForm, List<int>>() {
                                        { ControlForm.Credit, new List<int>() },
                                        { ControlForm.Exam, new List<int>() },
                                        { ControlForm.GradedCredit, new List<int>() },
                                    };

        public int TotalHoursByPlan = 0;
        public bool isCourseWork = false;






        public ModuleProperties() { }

        public ModuleProperties(string blockName, string partName, List<ControlForm> controlForm, string creditUnits, int blockNumber = -1)
        {
            BlockName = blockName;
            BlockNumber = blockNumber;
            PartName = partName;
            CreditUnits = creditUnits;
        }
    }

    // Формы контроля: Экзамен, Зачет, Зачет с оценкой
    public enum ControlForm { Exam, Credit, GradedCredit, Error }
}
