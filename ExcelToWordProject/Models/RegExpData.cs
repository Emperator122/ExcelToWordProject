using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelToWordProject.Models
{
    public class RegExpData
    {
        public string Expression;
        public RegexOptions RegexOptions;
        public int GroupIndex;

        public RegExpData() { }
    }
}
