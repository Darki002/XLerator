using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Editor;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.Editor;

[TestFixture]
public class ExcelEditorTest
{
    [Test]
    public void CreateFrom_InitializedCorrectly()
    {
        // Arrange
        const string filePath = "./CreateFrom_InitializedCorrectly.xlsx";
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };
        
        var sheetId = CreateFile(filePath, "Sheet1");
        
        // Act
        var testee = ExcelEditor<Dummy>.CreateFrom(options, new ExcelMapperDummy(), sheetId, 1);
        
        // Assert
        testee.Spreadsheet.Should().NotBeNull();
        testee.SheetData.Should().NotBeNull();
        
        // Clean Up
        testee.Dispose();
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
    
    [Test]
    public void Create_InitializedCorrectly()
    {
        // Arrange
        const string filePath = "./Create_InitializedCorrectly.xlsx";
        const string sheetName = "Sheet1";
        
        var options = new XLeratorOptions
        {
            FilePath = filePath,
            SheetName = sheetName
        };
        
        CreateFile(filePath, sheetName);
        
        // Act
        var testee = ExcelEditor<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Assert
        testee.Spreadsheet.Should().NotBeNull();
        testee.SheetData.Should().NotBeNull();
        
        // Clean Up
        testee.Dispose();
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    [Test]
    public void Write_AddsNewRowToSpreadSheet()
    {
        // Arrange
        const string filePath = "./Write_AddsNewRowToSpreadSheet.xlsx";
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };
        
        var sheetId = CreateFile(filePath, "Sheet1");
        var testee = ExcelEditor<HeaderedExcelClass>.CreateFrom(options, new ExcelMapperDummy(), sheetId, 1);
        
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
        
        // Clean Up
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
    
    private static StringValue CreateFile(string filePath, string sheetName)
    {
        var spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = spreadsheet.WorkbookPart?.Workbook.AppendChild(new Sheets());
        var sheetId = workbookPart.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            SheetId = 1,
            Name = sheetName
        };
        
        sheets?.Append(sheet);
        spreadsheet.Save();
        spreadsheet.Dispose();
        return sheetId;
    }
}