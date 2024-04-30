using System.Reflection;
using XLerator.Tags;

namespace XLerator;

public class ExcelConductor<T> where T : class
{
    internal readonly string FilePath;

    internal Dictionary<string, ExcelMap> Mapping;

    private ExcelConductor(string filePath, Dictionary<string, ExcelMap> mapping)
    {
        FilePath = filePath;
        Mapping = mapping;
    }

    public static ExcelConductor<T> Create(string filePath)
    {
        if (typeof(T).GetCustomAttribute<IndexedExcel>() is not null)
        {
            var indexMapping = CreateIndexMapping();
            return new ExcelConductor<T>(filePath, indexMapping);
        }
        
        var mapping = CreateHeaderMapping();
        return new ExcelConductor<T>(filePath, mapping);
    }

    private static Dictionary<string, ExcelMap> CreateHeaderMapping()
    {
        var mapping = new List<(string header, string property)>();
        var properties = typeof(T).GetProperties();

        if (properties.Length <= 0)
        {
            throw new ArgumentException($"Class {nameof(T)} must have at least 1 public property.");
        }

        foreach (var propertyInfo in properties)
        {
            var excelHeaderName = propertyInfo.GetCustomAttribute<ExcelHeaderName>();
            var name = excelHeaderName is not null ? excelHeaderName.Name : propertyInfo.Name;
            mapping.Add((name, propertyInfo.Name));
        }

        return mapping
            .Select((item, index) => (item, index))
            .ToDictionary(
                tuple => tuple.item.property, 
                tuple => new ExcelMap(tuple.index, tuple.item.header));
    }

    private static Dictionary<string, ExcelMap> CreateIndexMapping()
    {
        var mapping = new Dictionary<string, ExcelMap>();
        
        var properties = typeof(T).GetProperties()
            .Where(prop => Attribute.IsDefined(prop, typeof(IndexedExcel)))
            .Select(prop => (PropertyInfo: prop, Attribute: prop.GetCustomAttribute<ExcelIndex>()!)) 
            .ToList();
        
        if (properties.Count <= 0)
        {
            throw new ArgumentException($"Class {nameof(T)} must have at least 1 public property with the {nameof(ExcelIndex)} Attribute.");
        }
        
        foreach (var prop in properties)
        {
            var excelHeaderName = prop.PropertyInfo.GetCustomAttribute<ExcelHeaderName>();
            var name = excelHeaderName?.Name ?? prop.PropertyInfo.Name;
            mapping.Add(prop.PropertyInfo.Name, new ExcelMap(prop.Attribute.Index, name));
        }

        return mapping;
    }

    public ExcelReader CreateReader() => ExcelReader.Create(this);

    public ExcelWriter CreateWriter() => ExcelWriter.Create(this);
}