using JdeJabali.JXLDataTableExtractor.Exceptions;
using System;

namespace JdeJabali.JXLDataTableExtractor.Configuration
{
    public interface IDataTableExtractorSearchConfiguration
    {
        /// <exception cref="ArgumentException"/>
        /// <exception cref="DuplicateWorksheetException"/>
        IDataTableExtractorSearchConfiguration Worksheet(int worksheetIndex);

        /// <exception cref="ArgumentException"/>
        /// <exception cref="DuplicateWorksheetException"/>
        IDataTableExtractorSearchConfiguration Worksheets(int[] worksheetIndexes);

        /// <exception cref="ArgumentException"/>
        /// <exception cref="DuplicateWorksheetException"/>
        IDataTableExtractorSearchConfiguration Worksheet(string worksheet);

        /// <exception cref="ArgumentException"/>
        /// <exception cref="DuplicateWorksheetException"/>
        IDataTableExtractorSearchConfiguration Worksheets(string[] worksheets);

        /// <summary>
        /// Read all the worksheets in the workbook(s) specified.
        /// </summary>
        /// <returns></returns>
        IDataTableExtractorWorksheetConfiguration ReadAllWorksheets();
        IDataTableExtractorWorksheetConfiguration ReadOnlyTheIndicatedSheets();
    }
}