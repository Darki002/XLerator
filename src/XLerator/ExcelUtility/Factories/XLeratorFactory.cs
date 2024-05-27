namespace XLerator.ExcelUtility.Factories;

/// <summary>
/// Factory that can create different Instances of the Excel Utilities.
/// </summary>
/// <typeparam name="T">The type to Serialize or Deserialize.</typeparam>
public partial class XLeratorFactory<T> : IXLeratorFactory<T> where T : class
{
    private readonly XLeratorOptions options;
    
    private XLeratorFactory(XLeratorOptions options)
    {
        this.options = options;
    }

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