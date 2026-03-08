using System;

namespace SpreadsheetToPdf.Core
{
    internal class ExcelInteropException : Exception
    {
        public ExcelInteropException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    internal class ExcelNotInstalledException : Exception
    {
        public ExcelNotInstalledException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    internal class WorksheetNotFoundException : Exception
    {
        public WorksheetNotFoundException(string message)
            : base(message)
        {
        }
    }
}
