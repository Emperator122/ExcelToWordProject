using ExcelToWordProject.Models;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace ExcelToWordProject.Syllabus.Tags
{
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


        public override string GetValue(Module module = null, List<Content> contentList = null, DataSet excelData = null)
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
}
