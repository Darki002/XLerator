using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal class ExcelReader<T> : IExcelReader<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;

    private ExcelReader(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var reader = new ExcelReader<T>(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), false);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}