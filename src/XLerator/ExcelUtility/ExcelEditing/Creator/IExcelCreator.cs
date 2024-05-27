using XLerator.ExcelUtility.ExcelEditing.Editor;

namespace XLerator.ExcelUtility.ExcelEditing.Creator;

/// <summary>
/// Creates a new spreadsheet. It will be structure based on <typeparamref name="T"/>.
/// </summary>
/// <typeparam name="T">The type to Serialize or Deserialize.</typeparam>
public interface IExcelCreator<in T> where T : class
{
    /// <summary>
    /// Creates a new Excel file and returns a new Instance of a <see cref="IExcelEditor{T}"/>.
    /// </summary>
    /// <returns>The new Instance of a <see cref="IExcelEditor{T}"/>.</returns>
    IExcelEditor<T> CreateExcel();
}