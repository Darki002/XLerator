using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLerator.ExcelMappings;

namespace XLerator.ExcelUtility.Creator;

public class ExcelCreator : IDisposable
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelCreator(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelCreator Create(string filePath, ExcelMapperBase excelMapper)
    {
        var reader = new ExcelCreator(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}