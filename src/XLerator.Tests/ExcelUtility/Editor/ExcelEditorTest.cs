using DocumentFormat.OpenXml.Packaging;
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
        
        var testee = ExcelEditor<HeaderedExcelClass>.CreateFrom(spreadsheet, mapper, options);
        
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
}