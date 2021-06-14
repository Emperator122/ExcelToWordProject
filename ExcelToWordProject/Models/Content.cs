using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject.Models
{
    public class Content
    {
        public string Index;
        public string Value;

        public Content(string index, string value)
        {
            Index = index;
            Value = value;
        }
    }
}
