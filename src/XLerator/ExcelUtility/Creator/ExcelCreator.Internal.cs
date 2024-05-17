using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal partial class ExcelCreator<T>
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
}