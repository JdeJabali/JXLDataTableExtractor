using JdeJabali.JXLDataTableExtractor.JXLExtractedData;
using System.Collections.Generic;
using System.Data;

namespace JdeJabali.JXLDataTableExtractor.Configuration
{
    public interface IDataExtraction
    {
        /// <summary>
        /// Extract all the data from the workbook(s).
        /// </summary>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"/>
        List<JXLWorkbookData> GetWorkbooksData();

        /// <summary>
        /// Only retrieves the extracted rows from all the workbooks and worksheets.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"/>
        List<JXLExtractedRow> GetExtractedRows();

        /// <summary>    
        /// Convert the result of <see cref="GetWorkbooksData"/> to a DataTable.
        /// </summary>
        /// <returns>The <see cref="DataTable"/> with all the rows read.</returns>
        /// <exception cref="InvalidOperationException"/>
        DataTable GetDataTable();
    }
}