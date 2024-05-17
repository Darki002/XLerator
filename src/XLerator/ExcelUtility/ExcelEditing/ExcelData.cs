using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelEditing;

internal class ExcelData<T> : ExcelRow where T : class
{
    public readonly uint RowIndex;
    
    private readonly T? data;

    private ExcelData(T? data, uint rowIndex)
    {
        this.data = data;
        this.RowIndex = rowIndex;
    }
    
    internal static ExcelData<T> CreateFrom(T data, uint rowIndex, ExcelMapperBase excelMapper)
    {
        var row = new ExcelData<T>(data, rowIndex);
        row.CreateCells(excelMapper);
        return row;
    }
    
    private void CreateCells(ExcelMapperBase excelMapper)
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
            
            Row.Add(new ExcelCell(col, RowIndex, value!));
        }
    }
}