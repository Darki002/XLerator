using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal class Spreadsheet : IDisposable
{
    private SpreadsheetDocument document;

    private WorksheetPart worksheetPart;
    
    public SheetData SheetData { get; private set; }

    private Spreadsheet(
        SpreadsheetDocument document, 
        WorksheetPart worksheetPart, 
        SheetData sheetData)
    {
        this.document = document;
        this.worksheetPart = worksheetPart;
        SheetData = sheetData;
    }

    public static Spreadsheet Create(XLeratorOptions options)
    {
        var document = SpreadsheetDocument.Create(options.FilePath, SpreadsheetDocumentType.Workbook);
        
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        var sheetId = workbookPart.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            SheetId = 1,
            Name = options.SheetName
        };
        
        sheets.Append(sheet);
        document.Save();
        document.Dispose();
        
        document = SpreadsheetDocument.Open(options.FilePath, true);
        var result = GetWorksheetPartByName(document, options.SheetName);
        
        var spreadsheet = new Spreadsheet(
            document: document,
            worksheetPart: result.worksheetPart,
            sheetData: result.sheetData);
        return spreadsheet;
    }

    public static Spreadsheet Open(XLeratorOptions options, bool isEditable)
    {
        var document = SpreadsheetDocument.Open(options.FilePath, isEditable);
        var result = GetWorksheetPartByName(document, options.SheetName);
        
            var spreadsheet = new Spreadsheet(
            document: document,
            worksheetPart: result.worksheetPart,
            sheetData: result.sheetData);
        return spreadsheet;
    }
    
    private static (WorksheetPart worksheetPart, SheetData sheetData) GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
    {
        var sheets = document.WorkbookPart?.Workbook.Sheets!;
        foreach (var sheet in sheets.Elements<Sheet>())
        {
            if (sheet.Name != sheetName || sheet.Id == null)
            {
                continue;
            }
            var worksheetPart = (WorksheetPart)document.WorkbookPart?.GetPartById(sheet.Id!)!;
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            return (worksheetPart, sheetData);
        }

        throw new InvalidOperationException("The SheetData was not initialized correctly.");
    }
    
    public Row? LastRowOrDefault() => SheetData.Elements<Row>().LastOrDefault();

    public void AppendRow(Row row)
    {
        var existingRow = SheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == row.RowIndex);
        if (existingRow != null)
        {
            SheetData.RemoveChild(existingRow);
        }
        
        SheetData.Append(row);
    }

    public void Save()
    {
        worksheetPart.Worksheet.Save();
        document.Save();
    }

    public void SaveWorksheet()
    {
        worksheetPart.Worksheet.Save();
    }

    public void Dispose()
    {
        document.Dispose();
        document = null!;
        worksheetPart = null!;
        SheetData = null!;
    }
}