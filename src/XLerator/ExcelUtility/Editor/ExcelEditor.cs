using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal class ExcelEditor<T> : IExcelEditor<T> where T : class
{
    private readonly XLeratorOptions xLeratorOptions;
    private readonly ExcelMapperBase excelMapper;

    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelEditor(XLeratorOptions xLeratorOptions, ExcelMapperBase excelMapper)
    {
        this.xLeratorOptions = xLeratorOptions;
        this.excelMapper = excelMapper;
    }

    internal static IExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var editor = new ExcelEditor<T>(options, excelMapper);
        editor.spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), true);
        return editor;
    }
    
    public void Write(T data)
    {
        throw new NotImplementedException();
    }

    public void WriteMany(IEnumerable<T> data)
    {
        throw new NotImplementedException();
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}