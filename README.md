
# JdeJabali.JXLDataTableExtractor
[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)

Extract data as tables from Excel. Search columns by their header or index number. Sets conditions for extracting the rows. 

Read one or many workbooks. Select what worksheets should be read, by index number or name.

Get the result in a DataTable or in a collection of rows.



## Demo

    string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
    string subFolderPath = Path.Combine(path, "Daily store sales");
    string[] workbooks = Directory.GetFiles(subFolderPath, "*.xls");

    DataTable dataTable = DataTableExtractor
        .Configure()
        .Workbooks(workbooks)
        .SearchLimits(searchLimitRow: 10, searchLimitColumn: 30)
        .Worksheet(worksheetIndex: 0)
        .ReadOnlyTheIndicatedSheets()
        .ColumnHeader("Description")
        .ColumnHeader("Sales Value")
        .ColumnHeader("Discounts")
        .ColumnHeader("VAT")
            .ConditionToExtractRow(ConditionalsToExtractRow.HasNumericValueAboveZero)
        .ColumnIndex(columnIndex: 7)
        .CustomColumnHeaderMatch(cellValue => cellValue.Contains("Total"))
            .ConditionToExtractRow(cellValue => !string.IsNullOrEmpty(cellValue))
        .GetDataTable();


## Documentation

"Search Limits" apply to all the worksheets to read.

Instead of 

    .ReadOnlyTheIndicatedSheets()

use
    
    .ReadAllWorksheets()

to read every worksheet in every workbook.

If after this line for example

    .ColumnHeader("Description")

this other line is not specified

    .ConditionToExtractRow(condition)    

Then the row will always be extracted (although for that all other conditions must be met).

You may want to get a collection of rows instead of a DataTable.

for that, change 

    .GetDataTable();

with

    .GetExtractedRows();

at the end.

If you need more details about the rows, for example which workbook or sheet they were extracted from, then you might want to use this line at the end.

    .GetWorkbooksData();

That's all
## Feedback

If you have any feedback, please reach out to me at jdejabali@outlook.com

