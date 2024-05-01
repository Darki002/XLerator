using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

public class ExcelReader : IDisposable
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;

    private ExcelReader(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader Create(string filePath, ExcelMapperBase excelMapper)
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