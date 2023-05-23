namespace JdeJabali.JXLDataTableExtractor.Configuration
{
    public interface IDataTableExtractorSearchConfiguration
    {
        /// <exception cref="ArgumentException"/>
        IDataTableExtractorSearchConfiguration Worksheet(int worksheetIndex);

        /// <exception cref="ArgumentException"/>
        IDataTableExtractorSearchConfiguration Worksheets(int[] worksheetIndexes);

        /// <exception cref="ArgumentException"/>
        IDataTableExtractorSearchConfiguration Worksheet(string worksheet);

        /// <exception cref="ArgumentException"/>
        IDataTableExtractorSearchConfiguration Worksheets(string[] worksheets);

        /// <summary>
        /// Read all the worksheets in the workbook(s) specified.
        /// </summary>
        /// <returns></returns>
        IDataTableExtractorWorksheetConfiguration ReadAllWorksheets();
        IDataTableExtractorWorksheetConfiguration ReadOnlyTheIndicatedSheets();
    }
}