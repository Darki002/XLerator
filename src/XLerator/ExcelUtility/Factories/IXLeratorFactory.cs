using XLerator.ExcelUtility.ExcelEditing.Creator;
using XLerator.ExcelUtility.ExcelEditing.Editor;
using XLerator.ExcelUtility.ExcelReading.Iterator;
using XLerator.ExcelUtility.ExcelReading.Reader;

namespace XLerator.ExcelUtility.Factories;

/// <summary>
/// Factory that can create different Instances of the Excel Utilities.
/// </summary>
/// <typeparam name="T">The type to Serialize or Deserialize.</typeparam>
public interface IXLeratorFactory<T> where T : class
{
    /// <summary>
    /// Creates new Instance of an <see cref="IExcelCreator{T}"/>
    /// </summary>
    /// <returns>A new Instance of a ExcelCreator</returns>
    IExcelCreator<T> CreateExcelCreator();
    
    /// <summary>
    /// Creates new Instance of an <see cref="IExcelReader{T}"/>
    /// </summary>
    /// <returns>A new Instance of a ExcelReader</returns>
    IExcelReader<T> CreateExcelReader();

    /// <summary>
    /// Creates new Instance of an <see cref="IExcelEditor{T}"/>
    /// </summary>
    /// <returns>A new Instance of a ExcelEditor</returns>
    IExcelEditor<T> CreateExcelEditor();

    /// <summary>
    /// Create new Instance of an <see cref="IExcelIterator{T}"/>
    /// </summary>
    /// <returns>A new Instance of a ExcelIterator</returns>
    IExcelIterator<T> CreateExcelIterator();
}