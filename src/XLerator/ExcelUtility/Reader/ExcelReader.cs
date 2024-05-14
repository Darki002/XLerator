using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal class ExcelReader<T> : IExcelReader<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private Spreadsheet spreadsheet;

    private ExcelReader(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.spreadsheet = spreadsheet;
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelReader<T>(spreadsheet, excelMapper);
    }

    public T GetCell(uint row, uint column)
    {
        throw new NotImplementedException();
    }

    public T GetCell(string cellReference)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRange(uint column, uint lowerRow, uint upperRow)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRange(Range rowRange, uint column)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRange(uint lowerRow, uint lowerColumn, uint upperRow, uint upperColumn)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRange(Range rowRange, Range columnRange)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRow(uint row)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRows(uint lowerBound, uint upperBound)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRows(Range range)
    {
        throw new NotImplementedException();
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}