using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

public class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;
    
    private ExcelCreator(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelCreator<T> Create(string filePath, ExcelMapperBase excelMapper)
    {
        var reader = new ExcelCreator<T>(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
        return reader;
    }

    public void CreateExcel(IEnumerable<T> rows)
    {
        foreach (var row in rows)
        {
            
        }
    }

    private void WriteRow(T row, int rowIndex)
    {
        var propertyInfos = typeof(T).GetProperties();

        foreach (var propertyInfo in propertyInfos)
        {
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            if(col is null) continue;
            
            // TODO write cell with data
        }
    }

    private void CreateHeader()
    {
        var propertyInfos = typeof(T).GetProperties();

        foreach (var propertyInfo in propertyInfos)
        {
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            var header = excelMapper.GetHeaderFor(propertyInfo.Name);
            if(col is null || header is null) continue;
            
            // TODO write cell with header name
        }
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}