﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Editor;
using XLerator.Tests.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.Editor;

[TestFixture]
public class ExcelEditorTest
{
    [Test]
    public void Write_AddsNewRowToSpreadSheet()
    {
        // Arrange
        const string filePath = "./Write_AddsNewRowToSpreadSheet.xlsx";
        TestEnvironment.FilePaths.Add(filePath);
        
        var options = new XLeratorOptions
        {
            FilePath = filePath,
            SheetName = "Sheet1"
        };
        var spreadsheet = Spreadsheet.Create(options);

        var mapper = new ExcelMapperBaseFake();
        mapper.AddPropertyIndexMap(nameof(HeaderedExcelClass.Id), 1);
        mapper.AddPropertyIndexMap(nameof(HeaderedExcelClass.Name), 2);
        
        var testee = ExcelEditor<HeaderedExcelClass>.CreateFrom(spreadsheet, mapper);
        
        // Act
        var data = new HeaderedExcelClass
        {
            Id = 42,
            Name = "Test"
        };
        testee.Write(data);
        testee.Dispose();
        
        // Assert
        using (var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            var worksheetPart = workbookPart?.WorksheetParts.First();
            var sheetData = worksheetPart?.Worksheet.Elements<SheetData>().First();
            var rows = sheetData?.Elements<Row>().ToList();
            
            // Assert
            rows.Should().NotBeNull();
            rows!.Count.Should().Be(1);
            
            var newRow = rows.First();
            var cells = newRow.Elements<Cell>().ToList();
            
            // Assert
            cells.Count.Should().Be(2);
            
            var firstHeaderValue = cells[0].InnerText;
            var secondHeaderValue = cells[1].InnerText;
            
            // Assert
            firstHeaderValue.Should().Be(data.Id.ToString());
            secondHeaderValue.Should().Be(data.Name);
        }
    }
    
    [Test]
    public void WriteMany_AddsNewRowsToSpreadSheet()
    {
        // Arrange
        const string filePath = "./WriteMany_AddsNewRowsToSpreadSheet.xlsx";
        TestEnvironment.FilePaths.Add(filePath);
        
        var options = new XLeratorOptions
        {
            FilePath = filePath,
            SheetName = "Sheet1"
        };
        var spreadsheet = Spreadsheet.Create(options);

        var mapper = new ExcelMapperBaseFake();
        mapper.AddPropertyIndexMap(nameof(HeaderedExcelClass.Id), 1);
        mapper.AddPropertyIndexMap(nameof(HeaderedExcelClass.Name), 2);
        
        var testee = ExcelEditor<HeaderedExcelClass>.CreateFrom(spreadsheet, mapper);
        
        // Act
        var data = new HeaderedExcelClass
        {
            Id = 42,
            Name = "Test"
        };
        
        var data2 = new HeaderedExcelClass
        {
            Id = 69,
            Name = "Test"
        };
        testee.WriteRows(new []{data, data2});
        testee.Dispose();
        
        // Assert
        using (var spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            var worksheetPart = workbookPart?.WorksheetParts.First();
            var sheetData = worksheetPart?.Worksheet.Elements<SheetData>().First();
            var rows = sheetData?.Elements<Row>().ToList();
            
            // Assert
            rows.Should().NotBeNull();
            rows!.Count.Should().Be(2);
            
            var firstRow = rows.First();
            var cells = firstRow.Elements<Cell>().ToList();
            
            // Assert
            cells.Count.Should().Be(2);
            
            var firstHeaderValue = cells[0].InnerText;
            var secondHeaderValue = cells[1].InnerText;
            
            // Assert
            firstHeaderValue.Should().Be(data.Id.ToString());
            secondHeaderValue.Should().Be(data.Name);
            
            var secondRow = rows[1];
            cells = secondRow.Elements<Cell>().ToList();
            
            // Assert
            cells.Count.Should().Be(2);
            
            firstHeaderValue = cells[0].InnerText;
            secondHeaderValue = cells[1].InnerText;
            
            // Assert
            firstHeaderValue.Should().Be(data2.Id.ToString());
            secondHeaderValue.Should().Be(data2.Name);
        }
    }
}