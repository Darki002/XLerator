namespace XLerator.ExcelUtility.ExcelReading.Iterator;

/// <summary>
/// Allows to read the data in the Excel file in iteratable steps.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet.</typeparam>
public interface IExcelIterator<out T> : IDisposable
{
    bool Read();

    T GetCurrentRow();
}