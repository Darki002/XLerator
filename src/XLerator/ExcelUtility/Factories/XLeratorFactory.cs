using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Reader;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Factories;

public class XLeratorFactory(string filePath)
{
    public ExcelReader CreateReader<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader.Create(filePath, mapper);
    }

    public IExcelCreator<T> CreateExcelCreator<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelCreator<T>.Create(filePath, mapper);
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