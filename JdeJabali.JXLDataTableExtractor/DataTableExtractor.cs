using JdeJabali.JXLDataTableExtractor.Configuration;
using JdeJabali.JXLDataTableExtractor.DataExtraction;
using JdeJabali.JXLDataTableExtractor.Exceptions;
using JdeJabali.JXLDataTableExtractor.JXLExtractedData;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace JdeJabali.JXLDataTableExtractor
{
    public class DataTableExtractor :
        IDataTableExtractorConfiguration,
        IDataTableExtractorWorkbookConfiguration,
        IDataTableExtractorSearchConfiguration,
        IDataTableExtractorWorksheetConfiguration
    {
        private bool _readAllWorksheets;
        private int _searchLimitRow;
        private int _searchLimitColumn;

        private readonly HashSet<string> _workbooks = new HashSet<string>();
        private readonly HashSet<int> _worksheetIndexes = new HashSet<int>();
        private readonly HashSet<string> _worksheets = new HashSet<string>();

        private readonly List<HeaderToSearch> _headersToSearch = new List<HeaderToSearch>();
        private HeaderToSearch _headerToSearch;

        private DataReader _reader;

        private DataTableExtractor()
        {
        }

        public static IDataTableExtractorConfiguration Configure()
        {
            return new DataTableExtractor();
        }

        public IDataTableExtractorWorkbookConfiguration Workbook(string workbook)
        {
            if (string.IsNullOrEmpty(workbook))
            {
                throw new ArgumentException($"{nameof(workbook)} cannot be null or empty.");
            }

            if (!_workbooks.Add(workbook))
            {
                throw new DuplicateWorkbookException("Cannot search for more than one workbook with the same name: " +
                    $@"""{workbook}"".");
            }

            return this;
        }

        public IDataTableExtractorWorkbookConfiguration Workbooks(string[] workbooks)
        {
            if (workbooks is null)
            {
                throw new ArgumentNullException($"{nameof(workbooks)} cannot be null.");
            }

            foreach (string workbook in workbooks)
            {
                Workbook(workbook);
            }

            return this;
        }

        public IDataTableExtractorSearchConfiguration SearchLimits(int searchLimitRow, int searchLimitColumn)
        {
            _searchLimitRow = searchLimitRow;
            _searchLimitColumn = searchLimitColumn;

            return this;
        }

        public IDataTableExtractorSearchConfiguration Worksheet(int worksheetIndex)
        {
            if (worksheetIndex < 0)
            {
                throw new ArgumentException($"{nameof(worksheetIndex)} cannot be less than zero.");
            }

            if (!_worksheetIndexes.Add(worksheetIndex))
            {
                throw new DuplicateWorksheetException("Cannot search for more than " +
                    $@"one worksheet with the same index: ""{worksheetIndex}"".");
            }

            return this;
        }

        public IDataTableExtractorSearchConfiguration Worksheets(int[] worksheetIndexes)
        {
            if (worksheetIndexes is null)
            {
                throw new ArgumentException($"{nameof(worksheetIndexes)} cannot be null or empty.");
            }

            foreach (int worksheet in worksheetIndexes)
            {
                Worksheet(worksheet);
            }

            return this;
        }

        public IDataTableExtractorSearchConfiguration Worksheet(string worksheet)
        {
            if (string.IsNullOrEmpty(worksheet))
            {
                throw new ArgumentException($"{nameof(worksheet)} cannot be null or empty.");
            }

            if (!_worksheets.Add(worksheet))
            {
                throw new DuplicateWorksheetException("Cannot search for more than " +
                    $@"one worksheet with the same name: ""{worksheet}"".");
            }

            return this;
        }

        public IDataTableExtractorSearchConfiguration Worksheets(string[] worksheets)
        {
            if (worksheets is null)
            {
                throw new ArgumentException($"{nameof(worksheets)} cannot be null or empty.");
            }

            foreach (string worksheet in worksheets)
            {
                Worksheet(worksheet);
            }

            return this;
        }

        public IDataTableExtractorWorksheetConfiguration ReadOnlyTheIndicatedSheets()
        {
            _readAllWorksheets = false;

            if (_worksheetIndexes.Count == 0 && _worksheets.Count == 0)
            {
                throw new InvalidOperationException("No worksheets selected.");
            }

            return this;
        }

        public IDataTableExtractorWorksheetConfiguration ReadAllWorksheets()
        {
            _readAllWorksheets = true;

            return this;
        }

        IDataTableExtractorWorksheetConfiguration IDataTableColumnsToSearch.ColumnHeader(string columnHeader)
        {
            if (string.IsNullOrEmpty(columnHeader))
            {
                throw new ArgumentException($"{nameof(columnHeader)} cannot be null or empty.");
            }

            if (_headersToSearch.FirstOrDefault(h => h.ColumnHeaderName == columnHeader) != null)
            {
                throw new DuplicateColumnException("Cannot search for more than one column header with the same name: " +
                    $@"""{columnHeader}"".");
            }

            _headerToSearch = new HeaderToSearch()
            {
                ColumnHeaderName = columnHeader,
            };

            _headersToSearch.Add(_headerToSearch);

            return this;
        }

        IDataTableExtractorWorksheetConfiguration IDataTableColumnsToSearch.ColumnIndex(int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentException($"{nameof(columnIndex)} cannot be less than zero.");
            }

            if (_headersToSearch.FirstOrDefault(h => h.ColumnIndex == columnIndex) != null)
            {
                throw new DuplicateColumnException("Cannot search for more than one column with the same index: " +
                    $@"""{columnIndex}"".");
            }

            _headerToSearch = new HeaderToSearch()
            {
                ColumnIndex = columnIndex,
            };

            _headersToSearch.Add(_headerToSearch);

            return this;
        }

        IDataTableExtractorWorksheetConfiguration IDataTableColumnsToSearch.CustomColumnHeaderMatch(Func<string, bool> conditional)
        {
            if (conditional is null)
            {
                throw new ArgumentNullException("Conditional cannot be null.");
            }

            _headerToSearch = new HeaderToSearch()
            {
                ConditionalToReadColumnHeader = conditional,
            };

            _headersToSearch.Add(_headerToSearch);

            return this;
        }

        IDataTableExtractorWorksheetConfiguration IDataTableExtractorColumnConfiguration.ConditionToExtractRow(Func<string, bool> conditional)
        {
            if (conditional is null)
            {
                throw new ArgumentNullException("Conditional cannot be null.");
            }

            if (_headerToSearch is null)
            {
                throw new InvalidOperationException(nameof(_headerToSearch));
            }

            _headerToSearch.ConditionalToReadRow = conditional;

            return this;
        }

        public List<JXLWorkbookData> GetWorkbooksData()
        {
            _reader = new DataReader()
            {
                Workbooks = _workbooks,
                SearchLimitRow = _searchLimitRow,
                SearchLimitColumn = _searchLimitColumn,
                WorksheetIndexes = _worksheetIndexes,
                Worksheets = _worksheets,
                ReadAllWorksheets = _readAllWorksheets,
                HeadersToSearch = _headersToSearch,
            };

            return _reader.GetWorkbooksData();
        }

        public List<JXLExtractedRow> GetExtractedRows()
        {
            _reader = new DataReader()
            {
                Workbooks = _workbooks,
                SearchLimitRow = _searchLimitRow,
                SearchLimitColumn = _searchLimitColumn,
                WorksheetIndexes = _worksheetIndexes,
                Worksheets = _worksheets,
                ReadAllWorksheets = _readAllWorksheets,
                HeadersToSearch = _headersToSearch,
            };

            return _reader.GetJXLExtractedRows();
        }

        public DataTable GetDataTable()
        {
            _reader = new DataReader()
            {
                Workbooks = _workbooks,
                SearchLimitRow = _searchLimitRow,
                SearchLimitColumn = _searchLimitColumn,
                WorksheetIndexes = _worksheetIndexes,
                Worksheets = _worksheets,
                ReadAllWorksheets = _readAllWorksheets,
                HeadersToSearch = _headersToSearch,
            };

            return _reader.GetDataTable();
        }
    }
}