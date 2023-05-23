using System;
using System.Runtime.Serialization;

namespace JdeJabali.JXLDataTableExtractor.Exceptions
{
    [Serializable]
    public class DuplicateColumnException : Exception
    {
        public DuplicateColumnException()
        {
        }

        public DuplicateColumnException(string message) : base(message)
        {
        }

        public DuplicateColumnException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected DuplicateColumnException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
