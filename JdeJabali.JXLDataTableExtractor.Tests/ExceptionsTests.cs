using FluentAssertions;
using JdeJabali.JXLDataTableExtractor.Exceptions;

namespace JdeJabali.JXLDataTableExtractor.Tests;

public class ExceptionsTests
{
    [Fact]
    public void AddingDuplicateWorkbooks_Return_DuplicateWorkbookException()
    {
        string[] books = { "1.xls", "1.xls" };

        Action act = () => DataTableExtractor.Configure().Workbooks(books);

        act.Should().Throw<DuplicateWorkbookException>();
    }

    [Fact]
    public void AddingWorksheetByIndexLessThanZero_Return_ArgumentException()
    {
        string[] worksheets = { "sheet1", "sheet1" };

        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
        .Worksheet(int.MinValue);

        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddingWorksheetByEmptyName_Return_ArgumentException()
    {
        string[] worksheets = { "sheet1", "sheet1" };

        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
        .Worksheet(string.Empty);

        act.Should().Throw<ArgumentException>();
    }

    [Fact]
    public void AddingDuplicateWorksheet_Return_DuplicateWorksheetException()
    {
        string[] worksheets = { "sheet1", "sheet1" };

        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
        .Worksheets(worksheets);

        act.Should().Throw<DuplicateWorksheetException>();
    }

    [Fact]
    public void AddingDuplicateWorksheetIndex_Return_DuplicateWorksheetException()
    {
        int[] worksheets = { 1, 1 };

        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
            .Worksheets(worksheets);

        act.Should().Throw<DuplicateWorksheetException>();
    }

    [Fact]
    public void SearchForBothWorksheetIndexesAndNames_Return_DuplicateWorksheetException()
    {
        string[] worksheets = { "sheet1", "sheet1" };
        int[] worksheetsIndexes = { 1, 1 };

        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
            .Worksheets(worksheets)
            .Worksheets(worksheetsIndexes)
            .ReadOnlyTheIndicatedSheets()
            .GetExtractedRows();

        act.Should().Throw<DuplicateWorksheetException>();
    }

    [Fact]
    public void AddingDuplicateColumnHeader_Return_DuplicateColumnException()
    {
        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
            .Worksheet("sheet1")
            .ReadOnlyTheIndicatedSheets()
            .ColumnHeader("Aquatic")
            .ColumnHeader("Aquatic");

        act.Should().Throw<DuplicateColumnException>();
    }

    [Fact]
    public void AddingDuplicateColumnIndex_Return_DuplicateColumnException()
    {
        Action act = () =>
            DataTableExtractor.Configure().Workbook("1.xls").SearchLimits(1, 20)
            .Worksheet("sheet1")
            .ReadOnlyTheIndicatedSheets()
            .ColumnIndex(1)
            .ColumnIndex(1);

        act.Should().Throw<DuplicateColumnException>();
    }
}
