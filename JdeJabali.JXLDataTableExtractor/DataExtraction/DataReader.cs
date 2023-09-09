using JdeJabali.JXLDataTableExtractor.JXLExtractedData;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace JdeJabali.JXLDataTableExtractor.DataExtraction
{
    internal class DataReader
    {
        public HashSet<string> Workbooks { get; set; } = new HashSet<string>();
        public int SearchLimitRow { get; set; }
        public int SearchLimitColumn { get; set; }
        public HashSet<int> WorksheetIndexes { get; set; } = new HashSet<int>();
        public HashSet<string> Worksheets { get; set; } = new HashSet<string>();
        public bool ReadAllWorksheets { get; set; }
        public List<HeaderToSearch> HeadersToSearch { get; set; } = new List<HeaderToSearch>();

        public DataTable GetDataTable()
        {
            List<JXLWorkbookData> data = GetWorkbooksData();

            DataTable dataTable = new DataTable();

            List<HeaderToSearch> orderedColumns = HeadersToSearch.OrderBy(column => column.HeaderCoord.Column).ToList();

            foreach (HeaderToSearch headerCoord in orderedColumns)
            {
                if (!string.IsNullOrEmpty(headerCoord.ColumnHeaderName))
                {
                    dataTable.Columns.Add(headerCoord.ColumnHeaderName);
                }
                else
                {
                    dataTable.Columns.Add("Column " + headerCoord.ColumnIndex.ToString() + ":");
                }
            }

            if (data.Count == 0)
            {
                return dataTable;
            }

            List<JXLExtractedRow> extractedRows = data
                .SelectMany(workbookData => workbookData.WorksheetsData
                .SelectMany(worksheetData => worksheetData.Rows))
                .ToList();

            foreach (JXLExtractedRow extractedRow in extractedRows)
            {
                string[] columns = new string[extractedRow.Columns.Count];

                for (int i = 0; i < extractedRow.Columns.Count; i++)
                {
                    string columnName = string.Empty;

                    if (!string.IsNullOrEmpty(orderedColumns[i].ColumnHeaderName))
                    {
                        columnName = orderedColumns[i].ColumnHeaderName;
                    }
                    else
                    {
                        columnName = orderedColumns[i].HeaderCoord.Column.ToString();
                    }

                    if (extractedRow.Columns.TryGetValue(columnName, out string value))
                    {
                        columns[i] = value;
                    }
                }

                dataTable.Rows.Add(columns);
            }

            return dataTable;
        }

        public List<JXLExtractedRow> GetJXLExtractedRows()
        {
            List<JXLWorkbookData> data = GetWorkbooksData();

            List<JXLExtractedRow> extractedRows = data
                .SelectMany(workbookData => workbookData.WorksheetsData
                .SelectMany(worksheetData => worksheetData.Rows))
                .ToList();

            return extractedRows;
        }

        public List<JXLWorkbookData> GetWorkbooksData()
        {
            List<JXLWorkbookData> workbooksData = new List<JXLWorkbookData>();

            foreach (string workbook in Workbooks)
            {
                using (ExcelPackage excel = new ExcelPackage(DataReaderHelpers.GetFileStream(workbook)))
                {
                    JXLWorkbookData workbookData = new JXLWorkbookData();
                    workbookData.WorkbookPath = workbook;
                    workbookData.WorkbookName = Path.GetFileNameWithoutExtension(workbook);

                    List<JXLWorksheetData> worksheetsData = new List<JXLWorksheetData>();

                    Func<string, ExcelPackage, List<JXLWorksheetData>> getWorksheetsData =
                        ResolveGettingWorksheetsMethod();

                    if (getWorksheetsData != null)
                    {
                        worksheetsData = getWorksheetsData(workbook, excel);
                    }

                    workbookData.WorksheetsData.AddRange(worksheetsData);
                    workbooksData.Add(workbookData);
                }
            }

            return workbooksData;
        }

        private Func<string, ExcelPackage, List<JXLWorksheetData>> ResolveGettingWorksheetsMethod()
        {
            if (ReadAllWorksheets)
            {
                return GetAllWorksheetsData;
            }

            return GetWorksheetsByNameOrIndex;
        }

        /// <summary>
        /// It will not throw an exception if two worksheets match by index.
        /// Only the first will be read, and the second will be ignored.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="excel"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        private List<JXLWorksheetData> GetWorksheetsByNameOrIndex(string workbook, ExcelPackage excel)
        {
            List<JXLWorksheetData> worksheetsData = new List<JXLWorksheetData>();
            JXLWorksheetData worksheetData;

            List<int> worksheetsIndexes = new List<int>();

            foreach (string worksheetName in Worksheets)
            {
                try
                {
                    ExcelWorksheet sheet = DataReaderHelpers.GetWorksheetByName(worksheetName, excel);

                    if (sheet is null)
                    {
                        throw new IndexOutOfRangeException($@"Worksheet name not found: ""{worksheetName}"" in workbook: ""{workbook}"".");
                    }

                    worksheetsIndexes.Add(sheet.Index);

                    worksheetData = ExtractRows(sheet);
                    worksheetData.WorksheetName = sheet.Name;
                }
                catch
                {
                    throw new InvalidOperationException($@"Error reading worksheet by name: ""{worksheetName}"" " +
                        $@"in workbook: ""{workbook}""");
                }

                if (worksheetData != null)
                {
                    worksheetsData.Add(worksheetData);
                }
            }

            foreach (int worksheetIndex in WorksheetIndexes)
            {
                try
                {
                    ExcelWorksheet sheet = DataReaderHelpers.GetWorksheetByIndex(worksheetIndex, excel);

                    if (sheet is null)
                    {
                        throw new IndexOutOfRangeException($@"Worksheet index not found: ""{worksheetIndex}"" " +
                            $@"in workbook: ""{workbook}"".");
                    }

                    worksheetData = ExtractRows(sheet);
                    worksheetData.WorksheetName = sheet.Name;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($@"Error reading worksheet by index: ""{worksheetIndex}"" " +
                        $@"in workbook: ""{workbook}"". {ex.Message}");
                }

                if (worksheetData != null)
                {
                    worksheetsData.Add(worksheetData);
                }
            }

            return worksheetsData;
        }

        private List<JXLWorksheetData> GetAllWorksheetsData(string workbook, ExcelPackage excel)
        {
            List<JXLWorksheetData> worksheetsData = new List<JXLWorksheetData>();
            JXLWorksheetData worksheetData;

            foreach (ExcelWorksheet worksheet in excel.Workbook.Worksheets)
            {
                try
                {
                    worksheetData = ExtractRows(worksheet);
                    worksheetData.WorksheetName = worksheet.Name;
                }
                catch
                {
                    throw new InvalidOperationException($@"Error reading worksheet: ""{worksheet.Name}"" " +
                        $@"in workbook: ""{workbook}""");
                }

                if (worksheetData != null)
                {
                    worksheetsData.Add(worksheetData);
                }
            }

            return worksheetsData;
        }

        private JXLWorksheetData ExtractRows(ExcelWorksheet sheet)
        {
            JXLWorksheetData worksheetData = new JXLWorksheetData()
            {
                WorksheetName = sheet.Name
            };

            GetHeadersCoordinates(sheet);

            if (!AreHeadersInTheSameRow())
            {
                throw new InvalidOperationException(
                    $"The headers to look for found in {worksheetData.WorksheetName} do not match in the same row. Cannot continue.");
            }

            worksheetData = GetWorksheetData(sheet);

            return worksheetData;
        }

        private JXLWorksheetData GetWorksheetData(ExcelWorksheet sheet)
        {
            JXLWorksheetData worksheetData = new JXLWorksheetData();

            if (HeadersToSearch.Count == 0)
            {
                return worksheetData;
            }

            HeaderCoord? firstColumnWithHeader = HeadersToSearch
                .FirstOrDefault(h => !string.IsNullOrEmpty(h.ColumnHeaderName))?
                .HeaderCoord;

            // If there is no column with a header, it means that all columns must be found by their indexes.
            // In that case the rows will begin to be read from the first
            if (firstColumnWithHeader is null)
            {
                firstColumnWithHeader = new HeaderCoord
                {
                    Row = 0,
                    Column = 1,
                };
            }

            // sheet.Dimension? Because might be empty
            for (int row = firstColumnWithHeader.Value.Row + 1; row <= sheet.Dimension?.Rows; row++)
            {
                JXLExtractedRow extractedRow = GetRow(row, sheet, out bool canAddRow);

                if (extractedRow != null && canAddRow && extractedRow.Columns.Count > 0)
                {
                    worksheetData.Rows.Add(extractedRow);
                }
            }

            return worksheetData;
        }

        private JXLExtractedRow GetRow(int row, ExcelWorksheet sheet, out bool canAddRow)
        {
            canAddRow = true;

            JXLExtractedRow extractedRow = new JXLExtractedRow();

            foreach (HeaderToSearch headerToSearch in HeadersToSearch)
            {
                string cellValue = DataReaderHelpers
                    .GetCellValue(row, headerToSearch.HeaderCoord.Column, sheet);

                if (!headerToSearch.ConditionalToReadRow(cellValue))
                {
                    canAddRow = false;
                }

                if (!string.IsNullOrEmpty(headerToSearch.ColumnHeaderName))
                {
                    extractedRow.Columns.Add(headerToSearch.ColumnHeaderName, cellValue);
                }
                else
                {
                    extractedRow.Columns.Add(headerToSearch.HeaderCoord.Column.ToString(),
                        cellValue);
                }
            }

            return extractedRow;
        }

        private void GetHeadersCoordinates(ExcelWorksheet sheet)
        {
            AddHeaderCoordsFromWorksheetColumnConfigurations(sheet);

            AddHeaderCoordsFromWorksheetColumnIndexesConfigurations();

            AddHeaderCoordsFromCustomColumnHeaderMatch(sheet);
        }

        private void AddHeaderCoordsFromWorksheetColumnConfigurations(ExcelWorksheet sheet)
        {
            List<HeaderToSearch> headersToSearch = HeadersToSearch
                .Where(h => !string.IsNullOrEmpty(h.ColumnHeaderName))
                .ToList();

            foreach (HeaderToSearch headerToSearch in headersToSearch)
            {
                bool headerFound = false;

                for (int row = 1; row <= SearchLimitRow; row++)
                {
                    if (headerFound)
                    {
                        break;
                    }

                    for (int column = 1; column <= SearchLimitColumn; column++)
                    {
                        string cellValue = DataReaderHelpers.GetCellValue(row, column, sheet);

                        if (string.Equals(cellValue, headerToSearch.ColumnHeaderName, StringComparison.Ordinal))
                        {
                            headerToSearch.HeaderCoord = new HeaderCoord
                            {
                                Row = row,
                                Column = column,
                            };

                            headerFound = true;
                            break;
                        }
                    }
                }
            }
        }

        private void AddHeaderCoordsFromWorksheetColumnIndexesConfigurations()
        {
            List<HeaderToSearch> headersToSearch = HeadersToSearch
                .Where(h =>
                    string.IsNullOrEmpty(h.ColumnHeaderName) &&
                    h.ConditionalToReadColumnHeader is null &&
                    h.ColumnIndex != null)
                .ToList();

            foreach (HeaderToSearch headerToSearch in headersToSearch)
            {
                headerToSearch.HeaderCoord = new HeaderCoord
                {
                    Row = 1,
                    Column = headerToSearch.ColumnIndex.Value,
                };
            }
        }

        private void AddHeaderCoordsFromCustomColumnHeaderMatch(ExcelWorksheet sheet)
        {
            List<HeaderToSearch> headersToSearch = HeadersToSearch
                .Where(h =>
                    string.IsNullOrEmpty(h.ColumnHeaderName) &&
                    h.ConditionalToReadColumnHeader != null &&
                    h.ColumnIndex is null)
                .ToList();

            foreach (HeaderToSearch headerToSearch in headersToSearch)
            {
                bool headerFound = false;

                for (int row = 1; row <= SearchLimitRow; row++)
                {
                    if (headerFound)
                    {
                        break;
                    }

                    for (int column = 1; column <= SearchLimitColumn; column++)
                    {
                        string cellValue = DataReaderHelpers.GetCellValue(row, column, sheet);

                        if (headerToSearch.ConditionalToReadColumnHeader(cellValue))
                        {
                            headerToSearch.ColumnHeaderName = cellValue;
                            headerToSearch.HeaderCoord = new HeaderCoord
                            {
                                Row = row,
                                Column = column,
                            };

                            headerFound = true;
                            break;
                        }
                    }
                }
            }
        }

        private bool AreHeadersInTheSameRow()
        {
            if (!HeadersToSearch.Any())
            {
                throw new InvalidOperationException($"{nameof(HeadersToSearch)} is empty.");
            }

            int firstHeaderRow = HeadersToSearch.First().HeaderCoord.Row;

            return HeadersToSearch
                .Where(h => !string.IsNullOrEmpty(h.ColumnHeaderName))
                .ToList()
                .All(header => header.HeaderCoord.Row == firstHeaderRow);
        }
    }
}