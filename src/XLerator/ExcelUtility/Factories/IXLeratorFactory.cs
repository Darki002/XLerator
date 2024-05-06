using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Editor;
using XLerator.ExcelUtility.Reader;

namespace XLerator.ExcelUtility.Factories;

/// <summary>
/// Factory that can create different Instances of the Excel Utilities.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet </typeparam>
public interface IXLeratorFactory<in T> where T : class
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
    IExcelReader<T> CreateReader();

    /// <summary>
    /// Creates new Instance of an <see cref="IExcelEditor{T}"/>
    /// </summary>
    /// <returns>A new Instance of a ExcelEditor</returns>
    IExcelEditor<T> CreateExcelEditor();
}