using XLerator.Mappings;

namespace XLerator.ExcelUtility;

internal class ExcelData<T> : ExcelRow where T : class
{
    private readonly T? data;

    private ExcelData(T? data)
    {
        this.data = data;   
    }
    
    internal static ExcelData<T> CreateFrom(T data, uint rowIndex, ExcelMapperBase excelMapper)
    {
        var row = new ExcelData<T>(data);
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
}