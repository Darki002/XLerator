namespace XLerator.ExcelUtility.Editor;

/// <summary>
/// Allows to edit the Excel file.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelEditor<in T> : IDisposable where T : class
{
    /// <summary>
    /// Writes the data as a new row to the spreadsheet.
    /// </summary>
    /// <param name="data">The data for the new row</param>
    void Write(T data);
    
    /// <summary>
    /// Writes each Element of the Enumerable as new rows to the spreadsheet.
    /// </summary>
    /// <param name="data">The data for the new rows</param>
    void WriteMany(IEnumerable<T> data);
}