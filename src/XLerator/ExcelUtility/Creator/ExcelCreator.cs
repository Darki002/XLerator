using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.Editor;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private const uint RowIndex = 0;
    
    private readonly ExcelMapperBase excelMapper;
    private readonly XLeratorOptions xLeratorOptions;
    
    private StringValue sheetId = null!;
    
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
        var rowIndex = RowIndex;
        
        using (var spreadsheetDocument = CreateFile())
        {
            if (addHeader)
            {
                AddHeader(spreadsheetDocument);
                rowIndex++;
            }
        }
        
        return ExcelEditor<T>.Create(xLeratorOptions, excelMapper)
            .SetCurrentRow(rowIndex)
            .SetSheetId(sheetId);
    }

    private SpreadsheetDocument CreateFile()
    {
        var spreadsheet = SpreadsheetDocument.Create(xLeratorOptions.GetFilePath(), SpreadsheetDocumentType.Workbook);
        
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = spreadsheet.WorkbookPart?.Workbook.AppendChild(new Sheets());
        sheetId = spreadsheet.WorkbookPart?.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            Name = xLeratorOptions.GetSheetNameOrDefault()
        };
        
        sheets?.Append(sheet);
        spreadsheet.Save();
        return spreadsheet;
    }
    
    private void AddHeader(SpreadsheetDocument spreadsheetDocument)
    {
       var row = ExcelHeader<T>.CreateFrom(RowIndex, excelMapper);
       
       var worksheetPart = (WorksheetPart?)spreadsheetDocument.WorkbookPart?.GetPartById(sheetId!);
       if (worksheetPart is null)
       {
           throw new InvalidOperationException("The Worksheet was not initialized correctly.");
       }
       var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
       var dataRow = new Row { RowIndex = 0 };
        
       foreach (var data in row)
       {
           dataRow.AppendChild(data.ToCell());
       }
        
       sheetData?.AppendChild(dataRow);
       spreadsheetDocument.Save();
    }
}