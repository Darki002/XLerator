﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;
using XLerator.ExcelUtility.ExcelEditing.Creator;
using XLerator.Tests.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.ExcelEditingTests.Creator;

[TestFixture]
public class ExcelCreatorTest
{
    [Test]
    public void CreateExcel_ReturnsNewIExcelEditor()
    {
        // Arrange
        const string filePath = "./CreateExcel_ReturnsNewIExcelEditor.xlsx";
        XLeratorTest.FilePaths.Add(filePath);
        
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };

        var testee = ExcelCreator<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Act
        var excelEditor = testee.CreateExcel();
        
        // Assert
        excelEditor.Should().NotBeNull();
        excelEditor.Dispose();
    }
    
    [Test]
    public void CreateExcel_CreatesANewExcelFile()
    {
        // Arrange
        const string filePath = "./CreateExcel_CreatesANewExcelFile.xlsx";
        XLeratorTest.FilePaths.Add(filePath);
        
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };

        var testee = ExcelCreator<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Act
        var editor = testee.CreateExcel();
        editor.Dispose();
        
        // Assert
        var fileExist = File.Exists(filePath);
        fileExist.Should().BeTrue();
    }

    [Test]
    public void CreateExcel_CreatesHeader_WhenSetTrue()
    {
        // Arrange
        const string filePath = "./CreateExcel_CreatesHeader_WhenSetTrue.xlsx";
        XLeratorTest.FilePaths.Add(filePath);
        
        const string sheetName = "Sheet";
        var options = new XLeratorOptions
        {
            FilePath = filePath,
            SheetName = sheetName,
            HeaderLength = 2
        };

        var excelMapper = new ExcelMapperBaseFake();
        excelMapper.AddHeaderMap("Id", "Index");
        excelMapper.AddPropertyIndexMap("Id", 1);
        excelMapper.AddHeaderMap("Name", "Name");
        excelMapper.AddPropertyIndexMap("Name", 2);

        var testee = ExcelCreator<HeaderedExcelClass>.Create(options, excelMapper);
        
        // Act
        var excelEditor = testee.CreateExcel();
        excelEditor.Dispose();
        
        // Assert
        using var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart?.WorksheetParts.First();
        var sheetData = worksheetPart?.Worksheet.Elements<SheetData>().First();
        var rows = sheetData?.Elements<Row>().ToList();

        // Assert
        rows.Should().NotBeNull();
        rows!.Count.Should().Be(1);

        var headerRow = rows.First();
        headerRow.RowIndex.Should().Be(2u);
        var cells = headerRow.Elements<Cell>().ToList();
            
        // Assert
        cells.Count.Should().Be(2);

        var firstHeaderValue = cells[0].InnerText;
        var secondHeaderValue = cells[1].InnerText;

        // Assert
        firstHeaderValue.Should().Be("Index");
        secondHeaderValue.Should().Be("Name");
    }
}