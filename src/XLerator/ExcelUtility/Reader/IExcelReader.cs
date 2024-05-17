namespace XLerator.ExcelUtility.Reader;

/// <summary>
/// Allows to read the data in the Excel file in precise locations.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelReader<T> : IDisposable where T : class
{
    T GetRow(int rowIndex);

    List<T> GetRows(int lowerBound, int upperBound);
    
    List<T> GetRange(int column, int lowerRow, int upperRow);
}