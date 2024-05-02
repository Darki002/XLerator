namespace XLerator.ExcelUtility.Creator;

public interface IExcelCreator<in T> : IDisposable where T : class
{
    /// <summary>
    /// Creates a Header in the Excel, using <see cref="XLerator.Attributes.ExcelHeaderName"/> or else the Property Name.
    /// <exception cref="InvalidOperationException">If you call this methode after any other methode was called that writes data to the spreadsheet</exception>
    /// </summary>
    void CreateHeader();

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