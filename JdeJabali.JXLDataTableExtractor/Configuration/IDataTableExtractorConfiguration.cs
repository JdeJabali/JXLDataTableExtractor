using JdeJabali.JXLDataTableExtractor.Exceptions;
using System;
using System.Collections.Generic;

namespace JdeJabali.JXLDataTableExtractor.Configuration
{
    public interface IDataTableExtractorConfiguration
    {
        /// <summary>
        /// The route of the workbook to read
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"/>
        /// <exception cref="DuplicateWorkbookException"/>
        IDataTableExtractorWorkbookConfiguration Workbook(string workbook);

        /// <summary>
        /// The route of the workbooks to read
        /// </summary>
        /// <param name="workbooks"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"/>        
        /// <exception cref="DuplicateWorkbookException"/>
        IDataTableExtractorWorkbookConfiguration Workbooks(IEnumerable<string> workbooks);
    }
}
