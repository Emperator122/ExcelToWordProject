using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWordProject.Syllabus
{
    internal enum SyllabusDatabaseErrorType
    {
        UniqueDefaultValueError,
        DeleteError,
    }

    internal abstract class SyllabusException : Exception
    {
        protected SyllabusException(string message = "") : base(message) { }
    }

    internal class SyllabusDatabaseException : SyllabusException
    {
        public SyllabusDatabaseErrorType ErrorType { get; }
        public SyllabusDatabaseException(SyllabusDatabaseErrorType errorType, string message = "") : base(message)
        {
            ErrorType = errorType;
        }
    }
}
