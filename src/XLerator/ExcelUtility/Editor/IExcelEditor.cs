namespace XLerator.ExcelUtility.Editor;

public interface IExcelEditor<in T> : IDisposable where T : class
{
    /// <summary>
    /// Writes the data as a new row to the spreadsheet.
    /// </summary>
    /// <param name="data">The data for the new row</param>
    void Write(T data);
    
    /// <summary>
    /// Writes the data each Element of the Enumerable as new rows to the spreadsheet.
    /// </summary>
    /// <param name="data">The data for the new rows</param>
    void WriteMany(IEnumerable<T> data);
}