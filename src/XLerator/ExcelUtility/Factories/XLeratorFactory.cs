using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Reader;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Factories;

public class XLeratorFactory(XLeratorOptions options)
{
    public IExcelReader CreateReader<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader.Create(options, mapper);
    }

    /// <summary>
    /// Creates new Instance of an <see cref="IExcelCreator{T}"/>
    /// </summary>
    /// <typeparam name="T">The type of data and structure of the spreadsheet</typeparam>
    /// <returns>The new Instance</returns>
    public IExcelCreator<T> CreateExcelCreator<T>() where T : class
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