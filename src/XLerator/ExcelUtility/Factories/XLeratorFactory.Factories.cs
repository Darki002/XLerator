using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelUtility.ExcelEditing.Creator;
using XLerator.ExcelUtility.ExcelEditing.Editor;
using XLerator.ExcelUtility.ExcelReading.Iterator;
using XLerator.ExcelUtility.ExcelReading.Reader;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Factories;

public partial class XLeratorFactory<T>
{
    public IExcelCreator<T> CreateExcelCreator()
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelCreator<T>.Create(options, mapper);
    }
    
    public IExcelEditor<T> CreateExcelEditor()
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelEditor<T>.Create(options, mapper);
    }
    
    public IExcelReader<T> CreateReader() 
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelReader<T>.Create(options, mapper);
    }

    public IExcelIterator<T> CreateIterator()
    {
        var mapper = CreateMapper(typeof(T));
        return ExcelIterator<T>.Create(options, mapper);
    }
    
    internal static ExcelMapperBase CreateMapper(Type type)
    {
        if (type.IsDefined(typeof(NoExcelIndexes)))
        {
            return HeaderExcelMapper.CreateFrom(type);
        }
        
        return IndexedExcelMapper.CreateFrom(type);
    }
}