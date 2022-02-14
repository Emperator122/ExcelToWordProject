using ExcelToWordProject.Models;
using System.Collections.Generic;
using System.Data;

namespace ExcelToWordProject.Syllabus.Tags
{
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
        public abstract string GetValue(Module module = null, List<Content> contentList = null, DataSet excelData = null);



        public BaseSyllabusTag() { }
    }
}
