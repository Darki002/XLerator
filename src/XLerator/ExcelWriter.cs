using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace XLerator;

public class ExcelWriter : IDisposable
{
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelWriter() { }

    internal static ExcelWriter Create<T>(string filePath) where T : class
    {
        var reader = new ExcelWriter();
        reader.spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}