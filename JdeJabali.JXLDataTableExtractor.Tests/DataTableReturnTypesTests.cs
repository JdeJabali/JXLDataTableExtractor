using FluentAssertions;
using JdeJabali.JXLDataTableExtractor.Configuration;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace JdeJabali.JXLDataTableExtractor.Tests;

public class DataTableReturnTypesTests
{
    [Fact]
    public void DataTableExtractor_Configure_Return_IDataTableExtractorConfiguration()
    {
        DataTableExtractor.Configure().Should().BeAssignableTo<IDataTableExtractorConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_Workbook_Return_IDataTableExtractorWorkbookConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .Should().BeAssignableTo<IDataTableExtractorWorkbookConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_Workbooks_Return_IDataTableExtractorWorkbookConfiguration()
    {
        DataTableExtractor.Configure().Workbooks(Array.Empty<string>())
            .Should().BeAssignableTo<IDataTableExtractorWorkbookConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_SearchLimits_Return_IDataTableExtractorSearchConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1).Should().BeAssignableTo<IDataTableExtractorSearchConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_WorksheetByWorksheetIndex_Return_IDataTableExtractorSearchConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .Worksheet(0)
            .Should().BeAssignableTo<IDataTableExtractorSearchConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_WorksheetsByWorksheetIndex_Return_IDataTableExtractorSearchConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .Worksheets(Array.Empty<int>())
            .Should().BeAssignableTo<IDataTableExtractorSearchConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_WorksheetByWorksheetName_Return_IDataTableExtractorSearchConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .Worksheet("test")
            .Should().BeAssignableTo<IDataTableExtractorSearchConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_WorksheetsByWorksheetName_Return_IDataTableExtractorSearchConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .Worksheets(Array.Empty<string>())
            .Should().BeAssignableTo<IDataTableExtractorSearchConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_ReadAllWorksheets_Return_IDataTableExtractorWorksheetConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .ReadAllWorksheets()
            .Should().BeAssignableTo<IDataTableExtractorWorksheetConfiguration>();
    }

    [Fact]
    public void DataTableExtractor_ReadOnlyTheIndicatedSheets_Return_IDataTableExtractorWorksheetConfiguration()
    {
        DataTableExtractor.Configure().Workbook("test")
            .SearchLimits(1, 1)
            .Worksheets(new string[2] { "sheet1", "sheet2" })
            .ReadOnlyTheIndicatedSheets()
            .Should().BeAssignableTo<IDataTableExtractorWorksheetConfiguration>();
    }
}