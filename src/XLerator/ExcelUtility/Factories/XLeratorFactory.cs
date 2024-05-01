using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelMappings;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Reader;

namespace XLerator.ExcelUtility.Factories;

public class XLeratorFactory(string filePath)
{
    public ExcelReader CreateReader<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader.Create(filePath, mapper);
    }

    public ExcelCreator CreateWriter<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelCreator.Create(filePath, mapper);
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