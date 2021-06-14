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
        public int BlockNumber = -1;
        public string BlockName = "";
        public string PartName = "";
        public List<ControlForm> Control = new List<ControlForm>();
        public int CreditUnits = -1;

        public ModuleProperties() { }

        public ModuleProperties(string blockName, string partName, List<ControlForm> controlForm, int creditUnits, int blockNumber = -1)
        {
            BlockName = blockName;
            BlockNumber = blockNumber;
            PartName = partName;
            Control = controlForm;
            CreditUnits = creditUnits;
        }

        public ModuleProperties(string parentModuleIndex, char del = '.')
        {
            string[] props = parentModuleIndex.Split(del);

            if (props.Length > 0)
            {
                BlockName = props[0];

                string temp = Regex.Replace(BlockName, @"[^\d]+", "");
                BlockNumber = (temp != "") ? Convert.ToInt32(temp) : -1;

                PartName = "";
            }
            else // если ничего не получилось распарсить
            {
                BlockName = "";
                BlockNumber = -1;
                PartName = "";
            }

        }

    }

    // Формы контроля: Экзамен, Зачет, Зачет с оценкой
    public enum ControlForm { Exam, Credit, GradedCredit, Error }
}
