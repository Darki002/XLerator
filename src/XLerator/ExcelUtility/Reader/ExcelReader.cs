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

    public T GetCell(int row, int column)
    {
        throw new NotImplementedException();
    }

    public T GetCell(string cellReference)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRange(int column, int lowerRow, int upperRow)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRange(Range rowRange, int column)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRange(int lowerRow, int lowerColumn, int upperRow, int upperColumn)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRange(Range rowRange, Range columnRange)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRow(int row)
    {
        throw new NotImplementedException();
    }

    public List<List<T>> GetRows(int lowerBound, int upperBound)
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