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
        public int TotalLessonsHours { get => LecturesHours + PracticalLessonsHours + LaboratoryLessonsHours; } // Итого аудиторных занятий
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

        public int BlockNumber = -1;
        public string BlockName = "";
        public string PartName = "";
        public List<ControlForm> Control = new List<ControlForm>();
        public int CreditUnits = -1;

        public List<int> Semesters = new List<int>();
        public int LecturesHours = 0;
        public int PracticalLessonsHours = 0;
        public int LaboratoryLessonsHours = 0;
        public int IndependentWorkHours = 0;
        public int ControlHours = 0;
        public int TotalHoursByPlan = 0;
        




        public ModuleProperties() { }

        public ModuleProperties(string blockName, string partName, List<ControlForm> controlForm, int creditUnits, int blockNumber = -1)
        {
            BlockName = blockName;
            BlockNumber = blockNumber;
            PartName = partName;
            Control = controlForm;
            CreditUnits = creditUnits;
        }
    }

    // Формы контроля: Экзамен, Зачет, Зачет с оценкой
    public enum ControlForm { Exam, Credit, GradedCredit, Error }
}
