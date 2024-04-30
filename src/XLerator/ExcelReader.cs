using DocumentFormat.OpenXml.Packaging;

namespace XLerator;

public class ExcelReader : IDisposable
{
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelReader() { }

    internal static ExcelReader Create<T>(ExcelConductor<T> excelConductor) where T : class
    {
        var reader = new ExcelReader();
        reader.spreadsheet = SpreadsheetDocument.Open(excelConductor.FilePath, false);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}