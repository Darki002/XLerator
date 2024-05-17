using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal partial class ExcelReader<T>
{
    private readonly ExcelMapperBase excelMapper;
    
    private Spreadsheet spreadsheet;
    
    private ExcelReader(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.spreadsheet = spreadsheet;
        this.excelMapper = excelMapper;
    }
    
    internal static ExcelReader<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelReader<T>(spreadsheet, excelMapper);
    }
}