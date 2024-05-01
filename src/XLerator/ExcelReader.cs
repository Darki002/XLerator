using DocumentFormat.OpenXml.Packaging;
using XLerator.ExcelMappings;

namespace XLerator;

public class ExcelReader : IDisposable
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;

    private ExcelReader(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader Create<T>(string filePath, ExcelMapperBase excelMapper) where T : class
    {
        var reader = new ExcelReader(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Open(filePath, false);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}