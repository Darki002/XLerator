using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Reader;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Factories;

/// <summary>
/// Factory that can create different different Instances of the Excel Utilities.
/// </summary>
/// <typeparam name="T">The type of data and structure of the spreadsheet </typeparam>
public class XLeratorFactory<T> : IXLeratorFactory<T> where T : class
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

    public IExcelReader CreateReader() 
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader.Create(options, mapper);
    }
    
    public IExcelCreator<T> CreateExcelCreator()
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelCreator<T>.Create(options, mapper);
    }

    internal static ExcelMapperBase CreateMapper(Type type)
    {
        if (type.IsDefined(typeof(IndexedExcel)))
        {
            return IndexedExcelMapper.CreateFrom(type);
        }
        
        return HeaderExcelMapper.CreateFrom(type);
    }
}