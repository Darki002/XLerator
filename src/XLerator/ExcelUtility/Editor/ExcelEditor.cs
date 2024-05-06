using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal class ExcelEditor<T> : IExcelEditor<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;

    internal SpreadsheetDocument Spreadsheet = null!;
    internal SheetData SheetData = null!;
    
    private uint currentRow;
    
    private ExcelEditor(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        currentRow = 0;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var editor = new ExcelEditor<T>(excelMapper);
        editor.Spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), true);
        editor.SheetData = editor.GetWorksheetPartByName(options.GetSheetNameOrDefault());
        return editor;
    }
    
    private SheetData GetWorksheetPartByName(string sheetName)
    {
        var sheets = Spreadsheet.WorkbookPart?.Workbook.Sheets!;
        foreach (var sheet in sheets.Elements<Sheet>())
        {
            if (sheet.Name != sheetName || sheet.Id == null)
            {
                continue;
            }
            var worksheetPart = (WorksheetPart)Spreadsheet.WorkbookPart?.GetPartById(sheet.Id!)!;
            return worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
        }

        throw new InvalidOperationException("The SheetData was not initialized correctly.");
    }
    
    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper, StringValue sheetId, uint currentRow)
    {
        var editor = new ExcelEditor<T>(excelMapper);
        
        editor.currentRow = currentRow;
        editor.Spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), true);
        
        var worksheetPart = (WorksheetPart?)editor.Spreadsheet.WorkbookPart?.GetPartById(sheetId!);
        if (worksheetPart is null)
        {
            throw new InvalidOperationException("The Worksheet was not initialized correctly.");
        }
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData is null)
        {
            throw new InvalidOperationException("The SheetData was not initialized correctly.");
        }
        editor.SheetData = sheetData;
        return editor;
    }
    
    public void Write(T data)
    {
        var row = ExcelData<T>.CreateFrom(data, currentRow, excelMapper);
        var dataRow = new Row { RowIndex = currentRow };
        
        foreach (var cell in row)
        {
            dataRow.AppendChild(cell.ToCell());
        }
        
        SheetData.AppendChild(dataRow);
        Spreadsheet.Save();
    }

    public void WriteMany(IEnumerable<T> data)
    {
        foreach (var rowData in data)
        {
            var row = ExcelData<T>.CreateFrom(rowData, currentRow, excelMapper);
            var dataRow = new Row { RowIndex = currentRow };
            
            foreach (var cell in row)
            {
                dataRow.AppendChild(cell.ToCell());
            }
            
            SheetData.AppendChild(dataRow);
        }
        
        Spreadsheet.Save();
    }
    
    public void Dispose()
    {
        Spreadsheet.Dispose();
        Spreadsheet = null!;
        SheetData = null!;
    }
}