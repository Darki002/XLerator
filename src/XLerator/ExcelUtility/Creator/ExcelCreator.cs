using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.Editor;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private const uint RowIndex = 1;
    
    private readonly ExcelMapperBase excelMapper;
    private readonly XLeratorOptions xLeratorOptions;
    
    private ExcelCreator(XLeratorOptions xLeratorOptions, ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        this.xLeratorOptions = xLeratorOptions;
    }

    internal static IExcelCreator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
       return new ExcelCreator<T>(options, excelMapper);
    }
    
    public IExcelEditor<T> CreateExcel(bool addHeader)
    {
       var spreadsheetDocument = CreateFile(out var sheetId);
        
        if (addHeader)
        {
            try
            {
                AddHeader(spreadsheetDocument, sheetId);
            }
            catch
            {
                spreadsheetDocument.Save();
                spreadsheetDocument.Dispose();
                throw;
            }
        }
        spreadsheetDocument.Save();
        spreadsheetDocument.Dispose();

        return ExcelEditor<T>.CreateFrom(xLeratorOptions, excelMapper, sheetId);
    }

    private SpreadsheetDocument CreateFile(out StringValue sheetId)
    {
        var spreadsheet = SpreadsheetDocument.Create(xLeratorOptions.GetFilePath(), SpreadsheetDocumentType.Workbook);
        
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheetId = workbookPart.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            SheetId = 1,
            Name = xLeratorOptions.GetSheetNameOrDefault()
        };
        
        sheets.Append(sheet);
        spreadsheet.Save();
        return spreadsheet;
    }
    
    private void AddHeader(SpreadsheetDocument spreadsheetDocument, StringValue sheetId)
    {
       var row = ExcelHeader<T>.CreateFrom(RowIndex, excelMapper);
       
       var worksheetPart = (WorksheetPart?)spreadsheetDocument.WorkbookPart?.GetPartById(sheetId!);
       if (worksheetPart is null)
       {
           throw new InvalidOperationException("The Worksheet was not initialized correctly.");
       }
       var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
       var dataRow = new Row { RowIndex = RowIndex };
        
       Cell? lastCell = null;
       foreach (var cell in row)
       {
           var newCell = cell.ToCell();
           dataRow.InsertAfter(newCell, lastCell);
           lastCell = newCell;
       }
        
       sheetData?.Append(dataRow);
       worksheetPart.Worksheet.Save();
    }
}