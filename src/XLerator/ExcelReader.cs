using DocumentFormat.OpenXml.Packaging;

namespace XLerator;

public class ExcelReader : IDisposable
{
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelReader() { }

    internal static ExcelReader Create<T>(string filePath) where T : class
    {
        var reader = new ExcelReader();
        reader.spreadsheet = SpreadsheetDocument.Open(filePath, false);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}