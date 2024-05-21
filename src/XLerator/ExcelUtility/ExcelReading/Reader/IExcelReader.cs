namespace XLerator.ExcelUtility.ExcelReading.Reader;

/// <summary>
///     Reads rows in the spreadsheet as the given type.
/// </summary>
/// <typeparam name="T">The type of row in the spreadsheet.</typeparam>
public interface IExcelReader<T> : IDisposable where T : class
{
    /// <summary>
    ///     Reads the row on the given Index.
    /// </summary>
    /// <param name="rowIndex">The index of the row</param>
    /// <returns>Returns a new Instance of <typeparamref name="T" />, representing the row in the spreadsheet.</returns>
    /// <exception cref="ArgumentException">When <paramref name="rowIndex" /> is less then zero.</exception>
    T GetRow(int rowIndex);

    /// <summary>
    ///     Reads all rows within the specified range, from <paramref name="lowerBound" /> to <paramref name="upperBound" />
    ///     (excluding <paramref name="upperBound" />).
    /// </summary>
    /// <param name="lowerBound">The index of the first row to read.</param>
    /// <param name="upperBound">The index of the row after the last row to read (not included in the result).</param>
    /// <returns>A list of new instances of <typeparamref name="T" />, representing the rows in the specified range.</returns>
    /// <exception cref="ArgumentException">
    ///     When <paramref name="lowerBound" /> or <paramref name="upperBound" /> is less then zero.
    ///     -or- when <paramref name="upperBound" /> is less than <paramref name="lowerBound" />
    /// </exception>
    List<T> GetRows(int lowerBound, int upperBound);
}