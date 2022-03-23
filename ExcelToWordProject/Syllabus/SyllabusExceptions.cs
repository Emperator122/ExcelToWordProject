using System;

namespace ExcelToWordProject.Syllabus
{
    internal enum SyllabusDatabaseErrorType
    {
        UniqueDefaultValueError,
        DeleteError,
        UpdateError,
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
