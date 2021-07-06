using ExcelToWordProject.Models;
using ExcelToWordProject.Syllabus;
using ExcelToWordProject.Utils;
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

        public ModuleProperties Properties;

        public Module(string index, string name,string сontentIndexesStr, char contentDelimiter = ';', ModuleProperties moduleProperties = null)
        {
            Properties = moduleProperties;
            Index = index;
            Name = name;
            ContentIndexes = сontentIndexesStr.Split(contentDelimiter);
            for (int i = 0; i < ContentIndexes.Length; i++)
            {
                ContentIndexes[i] = ContentIndexes[i].Trim();
            }
        }

        public override string ToString()
        {
            return Index + " " + Name;
        }
    }
}
