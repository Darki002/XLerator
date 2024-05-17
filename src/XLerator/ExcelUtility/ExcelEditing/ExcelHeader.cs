using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelEditing;

internal class ExcelHeader<T> : ExcelRow where T : class
{
    internal static ExcelHeader<T> CreateFrom(uint rowIndex, ExcelMapperBase excelMapper)
    {
        var row = new ExcelHeader<T>();
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
}