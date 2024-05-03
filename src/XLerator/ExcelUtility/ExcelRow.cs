using System.Collections;
using XLerator.Mappings;

namespace XLerator.ExcelUtility;

internal readonly struct ExcelRow<T> : IEnumerable<ExcelCell> where T : class
{
    private readonly T? data;

    public readonly List<ExcelCell> Row;

    public ExcelRow()
    {
        Row = new List<ExcelCell>();
    }
    
    private ExcelRow(T? data)
    {
        this.data = data;
        Row = new List<ExcelCell>();
    }

    internal static ExcelRow<T> CreateFrom(T data, uint rowIndex, ExcelMapperBase excelMapper)
    {
        var row = new ExcelRow<T>(data);
        row.CreateCells(rowIndex, excelMapper);
        return row;
    }
    
    private void CreateCells(uint rowIndex, ExcelMapperBase excelMapper)
    {
        var propertyInfos = data!.GetType().GetProperties();
        foreach (var propertyInfo in propertyInfos)
        {
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            if(col is null) continue;

            var value = propertyInfo.GetValue(data);
            var type = propertyInfo.PropertyType;
            
            if (value != null && value.GetType() != type)
            {
                value = Convert.ChangeType(value, type);
            }
            
            Row.Add(new ExcelCell(col, rowIndex, value));
        }
    }
    
    internal static ExcelRow<T> CreateHeader(uint rowIndex, ExcelMapperBase excelMapper)
    {
        var row = new ExcelRow<T>();
        row.CreateHeaderCells(rowIndex, excelMapper);
        return row;
    }

    private void CreateHeaderCells(uint rowIndex, ExcelMapperBase excelMapper)
    {
        var propertyInfos = typeof(T).GetProperties();
        foreach (var propertyInfo in propertyInfos)
        {
            var header = excelMapper.GetHeaderFor(propertyInfo.Name);
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            if(header is null || col is null) continue;
            
            Row.Add(new ExcelCell(col, rowIndex, header));
        }
    }

    internal ExcelCell this[int index] => Row[index];

    public IEnumerator<ExcelCell> GetEnumerator()
    {
        return Row.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}