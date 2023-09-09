using System;
using System.Runtime.Serialization;

namespace JdeJabali.JXLDataTableExtractor.Exceptions
{
    public class DuplicateWorksheetException : Exception
    {
        public DuplicateWorksheetException()
        {
        }

        public DuplicateWorksheetException(string message) : base(message)
        {
        }

        public DuplicateWorksheetException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DuplicateWorksheetException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
