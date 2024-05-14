namespace XLerator.ExcelUtility.Iterator;

/// <summary>
/// Allows to read the data in the Excel file in iteratable steps.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
public interface IExcelIterator<out T>
{
    bool Read();

    T GetCurrentRow();
    
    // TODO: add what ever an Iterator could possibly need
}