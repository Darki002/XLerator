using System.Reflection;
using XLerator.ExcelMappings;
using XLerator.Tags;

namespace XLerator;

public class XLeratorFactory(string filePath)
{
    public ExcelReader CreateReader<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader.Create<T>(filePath, mapper);
    }

    public ExcelWriter CreateWriter<T>() where T : class
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelWriter.Create<T>(filePath, mapper);
    }

    private static ExcelMapperBase CreateMapper(Type type)
    {
        if (type.IsDefined(typeof(IndexedExcel)))
        {
            return IndexedExcelMapper.CreateFrom(type);
        }
        
        return HeaderExcelMapper.CreateFrom(type);
    }
}