namespace XLerator.ExcelUtility.Reader;

/// <summary>
/// Allows to read the data in the Excel file in precise locations.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelReader<T> : IDisposable where T : class
{
    T GetCell(int row, int column);
    
    T GetCell(string cellReference);

    List<T> GetRange(int column, int lowerRow, int upperRow);
    
    List<T> GetRange(Range rowRange, int column);

    List<List<T>> GetRange(int lowerRow, int lowerColumn, int upperRow, int upperColumn);
    
    List<List<T>> GetRange(Range rowRange, Range columnRange);
    
    List<T> GetRow(int row);

    List<List<T>> GetRows(int lowerBound, int upperBound);
    
    List<List<T>> GetRows(Range range);
}