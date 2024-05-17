namespace XLerator.ExcelUtility.Factories;

/// <summary>
/// Factory that can create different Instances of the Excel Utilities.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet </typeparam>
public partial class XLeratorFactory<T> : IXLeratorFactory<T> where T : class
{
    /// <summary>
    /// Creates a new Instance of a <see cref="IXLeratorFactory{T}"/>.
    /// </summary>
    /// <param name="options">Options for the Excel file.</param>
    /// <returns>The new Instance.</returns>
    public static IXLeratorFactory<T> CreateFactory(XLeratorOptions options)
    {
        return new XLeratorFactory<T>(options);
    }
}