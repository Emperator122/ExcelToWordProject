﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject.Models
{
    class ExcelTableList
    {
        public string ListName;
        public DataSet ExcelData;

        public int RowHeaderIndex;
        public int ColumnHeaderIndex;

        /*
        **  +    val2    ...     valn     <-- RowHeaderIndex
        *   v1      
        *   
        *   v2
        *   
        *   ...
        *   
        *   vn    <-- ColumnHeaderIndex
        * 
        */

        public ExcelTableList(string listName, DataSet excelData, int rowHeaderIndex = -1, int columnHeaderIndex = -1)
        {
            ListName = listName;
            ExcelData = excelData;
            RowHeaderIndex = rowHeaderIndex;
            ColumnHeaderIndex = columnHeaderIndex;
        }

        public string GetCellValue(int rowIndex, int columnIndex)
        {
            return ExcelData.Tables[ListName].Rows[rowIndex][columnIndex] as string;
        }

        public List<string> GetCellValue(int rowIndex, string rowHeaderValue, bool first = false)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < ExcelData.Tables[ListName].Columns.Count; i++)
            {
                string val = ExcelData.Tables[ListName].Rows[RowHeaderIndex][i] as string;
                if (val == rowHeaderValue)
                {
                    result.Add(ExcelData.Tables[ListName].Rows[rowIndex][i] as string);
                    if (first) return result;
                }
                    
            }
            return result;
        }

        public List<string> GetCellValue(string columnHeaderValue, int columnIndex, bool first = false)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < ExcelData.Tables[ListName].Rows.Count; i++)
            {
                string val = ExcelData.Tables[ListName].Rows[i][ColumnHeaderIndex] as string;
                if (val == columnHeaderValue)
                {
                    result.Add(ExcelData.Tables[ListName].Rows[i][columnIndex] as string);
                    if (first) return result;
                }
            }
            return result;
        }

        public List<string> GetCellValue(string columnHeaderValue, string rowHeaderValue, bool first = false)
        {
            List<string> result = new List<string>();
            for (int i = 0; i < ExcelData.Tables[ListName].Rows.Count; i++)
            {
                string val = ExcelData.Tables[ListName].Rows[i][ColumnHeaderIndex] as string;
                if (val == columnHeaderValue)
                {
                    result.AddRange(GetCellValue(i, rowHeaderValue, first));
                    if (first) return result;
                }
            }
            return result;
        }

    }
}
