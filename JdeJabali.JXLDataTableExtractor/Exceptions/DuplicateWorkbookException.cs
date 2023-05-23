using System;
using System.Runtime.Serialization;

namespace JdeJabali.JXLDataTableExtractor.Exceptions
{
    [Serializable]
    public class DuplicateWorkbookException : Exception
    {
        public DuplicateWorkbookException()
        {
        }

        public DuplicateWorkbookException(string message) : base(message)
        {
        }

        public DuplicateWorkbookException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DuplicateWorkbookException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}