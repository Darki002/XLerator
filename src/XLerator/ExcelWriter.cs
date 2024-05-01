using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLerator.ExcelMappings;

namespace XLerator;

public class ExcelWriter : IDisposable
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelWriter(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelWriter Create<T>(string filePath, ExcelMapperBase excelMapper) where T : class
    {
        var reader = new ExcelWriter(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}