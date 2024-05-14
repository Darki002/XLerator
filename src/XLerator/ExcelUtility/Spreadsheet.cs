using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal class Spreadsheet : IDisposable
{
    public SpreadsheetDocument Document { get; private set; }

    public WorkbookPart WorkbookPart { get; private set; }

    public WorksheetPart WorksheetPart { get; private set; }
    
    public SheetData SheetData { get; private set; }
    
    public Sheets Sheets { get; private set; }

    public Sheet Sheet { get; private set; }

    private Spreadsheet(
        SpreadsheetDocument document, 
        WorkbookPart workbookPart, 
        WorksheetPart worksheetPart, 
        SheetData sheetData,
        Sheets sheets, 
        Sheet sheet)
    {
        Document = document;
        WorkbookPart = workbookPart;
        WorksheetPart = worksheetPart;
        SheetData = sheetData;
        Sheets = sheets;
        Sheet = sheet;
    }

    public static Spreadsheet Create(XLeratorOptions options)
    {
        var document = SpreadsheetDocument.Create(options.GetFilePath(), SpreadsheetDocumentType.Workbook);
        
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
            Name = options.GetSheetNameOrDefault()
        };
        
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        
        sheets.Append(sheet);
        document.Save();
        
        var spreadsheet = new Spreadsheet(
            document: document,
            workbookPart: workbookPart,
            worksheetPart: worksheetPart,
            sheetData: sheetData,
            sheets: sheets,
            sheet: sheet);
        return spreadsheet;
    }

    public static Spreadsheet Open(XLeratorOptions options, bool isEditable)
    {
        var document = SpreadsheetDocument.Open(options.GetFilePath(), isEditable);
        var result = GetWorksheetPartByName(document, options.GetSheetNameOrDefault());
        var workbookPart = document.WorkbookPart!;
        
            var spreadsheet = new Spreadsheet(
            document: document,
            workbookPart: workbookPart,
            worksheetPart: result.worksheetPart,
            sheetData: result.sheetData,
            sheets: result.sheets,
            sheet: result.sheet);
        return spreadsheet;
    }
    
    private static (WorksheetPart worksheetPart, Sheets sheets, Sheet sheet, SheetData sheetData) GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
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
            return (worksheetPart, sheets, sheet, sheetData);
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
        WorksheetPart.Worksheet.Save();
        Document.Save();
    }

    public void SaveWorksheet()
    {
        WorksheetPart.Worksheet.Save();
    }

    public void Dispose()
    {
        Document.Dispose();
        Document = null!;
        WorkbookPart = null!;
        WorksheetPart = null!;
        SheetData = null!;
        Sheets = null!;
        Sheet = null!;
    }
}