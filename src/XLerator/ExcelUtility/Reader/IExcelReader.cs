namespace XLerator.ExcelUtility.Reader;

/// <summary>
/// Allows to read the data in the Excel file in precise locations.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelReader<T> : IDisposable where T : class
{
    T GetCell(uint row, uint column);
    
    T GetCell(string cellReference);

    List<T> GetRange(uint column, uint lowerRow, uint upperRow);
    
    List<T> GetRange(Range rowRange, uint column);

    List<List<T>> GetRange(uint lowerRow, uint lowerColumn, uint upperRow, uint upperColumn);
    
    List<List<T>> GetRange(Range rowRange, Range columnRange);
    
    List<T> GetRow(uint row);

    List<List<T>> GetRows(uint lowerBound, uint upperBound);
    
    List<List<T>> GetRows(Range range);
}