using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal class ExcelReader<T> : IExcelReader<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private Spreadsheet spreadsheet;

    private ExcelReader(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.spreadsheet = spreadsheet;
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelReader<T>(spreadsheet, excelMapper);
    }

    public T GetRow(int rowIndex)
    {
        ThrowHelper.IfInvalidRowIndex(rowIndex);

        var row = spreadsheet.SheetData.Elements<Row>()
            .SingleOrDefault(r => r.RowIndex != null && r.RowIndex == rowIndex);

        ThrowHelper.ThrowIfNull(row, $"Row with index {rowIndex} does not exist.");

        var instanceType = typeof(T);
        
        var cells = row!.Elements<Cell>().ToList();
        var properties = instanceType.GetProperties();

        var instance = (T)Activator.CreateInstance(instanceType)!;

        foreach (var propertyInfo in properties)
        {
            var cellIndex = excelMapper.GetColumnIndexFor(propertyInfo.Name);
            ThrowHelper.ThrowIfNull(cellIndex, $"Excel file does not Match expected pattern of Type {typeof(T)}");
            var valueString = cells[(int)cellIndex! - 1].CellValue?.InnerText;
            
            var type = propertyInfo.PropertyType;

            if (valueString is null)
            {
                propertyInfo.SetValue(instance, Helper.GetDefaultValue(type));
                continue;
            }

            var converter = TypeDescriptor.GetConverter(type);
            var value = converter.ConvertFromString(valueString);
            
            propertyInfo.SetValue(instance, value);
        }

        return instance;
    }

    public T GetRowOrDefault(int rowIndex)
    {
        throw new NotImplementedException();
    }

    public List<T> GetRows(int lowerBound, int upperBound)
    {
        ThrowHelper.IfInvalidRowIndex(lowerBound);
        ThrowHelper.IfInvalidRowIndex(upperBound);

        throw new NotImplementedException();
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}