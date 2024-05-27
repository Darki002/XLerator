namespace XLerator.ExcelUtility.ExcelReading.Iterator;

/// <summary>
/// Allows to read the data in the Excel file in iteratable steps.
/// </summary>
/// <typeparam name="T">The type to Serialize or Deserialize.</typeparam>
public interface IExcelIterator<out T> : IDisposable where T : class
{
    /// <summary>
    /// Progresses the Iterator to the next row in the spreadsheet.
    /// </summary>
    /// <returns>
    /// True: when there are more rows left to read. <br/>
    /// False: when there is no row left to read.
    /// </returns>
    bool Read();

    /// <summary>
    /// Reads the current row and returns it as a new instance of <typeparamref name="T" />.
    /// </summary>
    /// <returns>Returns a new Instance of <typeparamref name="T" />, representing the row in the spreadsheet.</returns>
    /// <exception cref="InvalidOperationException">When there is present row to read.</exception>
    T GetCurrentRow();

    /// <summary>
    /// The iterator will skip over the given amount of rows.
    /// </summary>
    /// <param name="amount">How many rows should be skipped.</param>
    /// <exception cref="ArgumentException">When <paramref name="amount"/> is less or equal to zero.</exception>
    void SkipRows(int amount);
}