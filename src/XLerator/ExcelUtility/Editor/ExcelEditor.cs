using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal class ExcelEditor<T> : IExcelEditor<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;

    private SpreadsheetDocument spreadsheet = null!;
    
    private uint currentRow;
    
    private StringValue? sheetId;
    
    private ExcelEditor(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        currentRow = 0;
        sheetId = null;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var editor = new ExcelEditor<T>(excelMapper);
        editor.spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), true);
        return editor;
    }

    public ExcelEditor<T> SetCurrentRow(uint index)
    {
        currentRow = index;
        return this;
    }

    public ExcelEditor<T> SetSheetId(StringValue id)
    {
        sheetId = id;
        return this;
    }
    
    public void Write(T data)
    {
        var row = ExcelData<T>.CreateFrom(data, currentRow, excelMapper);
    }

    public void WriteMany(IEnumerable<T> data)
    {
        foreach (var cell in data)
        {
            Write(cell); // TODO optimize so only 1 Save is required
        }
    }

    private StringValue GetSheetId()
    {
        if (sheetId is not null) return sheetId;
        var worksheetPart = spreadsheet.WorkbookPart;
        sheetId = spreadsheet.WorkbookPart?.GetIdOfPart(worksheetPart!);
        return sheetId;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}