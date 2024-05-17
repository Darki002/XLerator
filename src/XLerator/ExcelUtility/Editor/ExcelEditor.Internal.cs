using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal partial class ExcelEditor<T>
{
    private readonly ExcelMapperBase excelMapper;
    
    private Spreadsheet spreadsheet;
    
    private ExcelEditor(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        this.spreadsheet = spreadsheet;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, true);
        return new ExcelEditor<T>(spreadsheet, excelMapper);
    }
    
    internal static ExcelEditor<T> CreateFrom(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        return new ExcelEditor<T>(spreadsheet, excelMapper);
    }
}