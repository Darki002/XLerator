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
    
    private ExcelEditor(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var editor = new ExcelEditor<T>(excelMapper);
        editor.Spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), true);
        editor.SheetData = editor.GetWorksheetPartByName(options.GetSheetNameOrDefault());
        
        // TODO find currentRow out. Look through all rows until one is empty
        
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
    
    internal static ExcelEditor<T> CreateFrom(XLeratorOptions options, ExcelMapperBase excelMapper, StringValue sheetId)
    {
        var editor = new ExcelEditor<T>(excelMapper);
        
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
        try
        {
            var lastRow = SheetData.Elements<Row>().LastOrDefault();
            var index = lastRow?.RowIndex ?? 1;
            
            var row = ExcelData<T>.CreateFrom(data, index, excelMapper);
            AddRow(row, index);
            Spreadsheet.Save();
        }
        catch
        {
            Spreadsheet.Save();
            throw;
        }
    }

    public void WriteMany(IEnumerable<T> data)
    {
        try
        {
            foreach (var rowData in data)
            {
                var lastRow = SheetData.Elements<Row>().LastOrDefault();
                var index = lastRow?.RowIndex ?? 1;
                
                var row = ExcelData<T>.CreateFrom(rowData, index, excelMapper);
                AddRow(row, index);
            }
            Spreadsheet.Save();
        }
        catch
        {
            Spreadsheet.Save();
            throw;
        }
    }

    private void AddRow(ExcelData<T> row, uint index)
    {
        var dataRow = new Row { RowIndex = index };
        
        Cell? lastCell = null;
        foreach (var cell in row)
        {
            var newCell = cell.ToCell();
            dataRow.InsertAfter(newCell, lastCell);
            lastCell = newCell;
        }
        SheetData.Append(dataRow);
    }
    
    public void Dispose()
    {
        Spreadsheet.Dispose();
        Spreadsheet = null!;
        SheetData = null!;
    }
}