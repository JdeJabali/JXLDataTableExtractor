using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

internal static class DataReaderHelpers
{
    public static string GetCellValue(int row, int column, ExcelWorksheet sheet)
    {
        return sheet.Cells[row, column].Value is null
            ? string.Empty
            : sheet.Cells[row, column].Value.ToString();
    }

    public static FileStream GetFileStream(string workbookFilename)
    {
        // Activate asynchronous read, or not?
        return new FileStream(
            workbookFilename,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            bufferSize: 4096,
            useAsync: false);
    }

    public static ExcelWorksheet GetWorksheetByIndex(int index, string workbook, ExcelPackage excel)
    {
        ExcelWorksheet worksheet = excel.Workbook.Worksheets
            .AsEnumerable()
            .FirstOrDefault(ws => ws.Index == index);

        if (worksheet is null)
        {
            throw new IndexOutOfRangeException($@"Worksheet index not found: ""{index}"" in ""{workbook}"".");
        }

        return worksheet;
    }

    public static ExcelWorksheet GetWorksheetByName(string worksheetName, string workbook, ExcelPackage excel)
    {
        ExcelWorksheet worksheet = excel.Workbook.Worksheets
            .AsEnumerable()
            .FirstOrDefault(ws => ws.Name == worksheetName);

        if (worksheet is null)
        {
            throw new IndexOutOfRangeException($@"Worksheet name not found: ""{worksheetName}"" in ""{workbook}"".");
        }

        return worksheet;
    }
}