using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal partial class ExcelReader<T> : IExcelReader<T> where T : class
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
        
        var cells = row!.Elements<Cell>().ToList();
        
        var instanceType = typeof(T);
        var properties = instanceType.GetProperties();
        var instance = (T)Activator.CreateInstance(instanceType)!;

        foreach (var propertyInfo in properties)
        {
            var type = propertyInfo.PropertyType;
            var valueString = GetCellValue(cells, propertyInfo.Name);
            
            propertyInfo.SetValue(instance, GetValueOrDefault(type, valueString));
        }

        return instance;
    }

    public List<T> GetRows(int lowerBound, int upperBound)
    {
        ThrowHelper.IfInvalidRowIndex(lowerBound);
        ThrowHelper.IfInvalidRowIndex(upperBound);

        if (lowerBound > upperBound)
        {
            throw new ArgumentException($"{nameof(lowerBound)} must be greater or equal then {nameof(upperBound)}");
        }

        return spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex != null && r.RowIndex >= lowerBound && r.RowIndex < upperBound)
            .Select(row => GetRow((int)row.RowIndex?.Value!))
            .ToList();
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}