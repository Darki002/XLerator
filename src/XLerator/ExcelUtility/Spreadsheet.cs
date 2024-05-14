using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal class Spreadsheet : IDisposable
{
    public SpreadsheetDocument Document { get; private set; }

    public WorkbookPart WorkbookPart { get; private set; }

    public WorksheetPart WorksheetPart { get; private set; }

    public Sheets Sheets { get; private set; }

    public StringValue SheetId { get; private set; }

    public Sheet Sheet { get; private set; }

    private Spreadsheet(
        SpreadsheetDocument document, 
        WorkbookPart workbookPart, 
        WorksheetPart worksheetPart, 
        Sheets sheets, 
        StringValue sheetId, 
        Sheet sheet)
    {
        Document = document;
        WorkbookPart = workbookPart;
        WorksheetPart = worksheetPart;
        Sheets = sheets;
        SheetId = sheetId;
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
        
        sheets.Append(sheet);
        document.Save();
        
        var spreadsheet = new Spreadsheet(
            document: document,
            workbookPart: workbookPart,
            worksheetPart: worksheetPart,
            sheets: sheets,
            sheetId: sheetId,
            sheet: sheet);
        return spreadsheet;
    }
    
    public SheetData? GetSheetData()
    {
        return WorksheetPart.Worksheet.GetFirstChild<SheetData>();
    }

    public void Save()
    {
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
        Sheets = null!;
        SheetId = null!;
        Sheet = null!;
    }
}