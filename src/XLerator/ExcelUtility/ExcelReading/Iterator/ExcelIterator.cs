using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading.Iterator;

internal class ExcelIterator<T> : IExcelIterator<T>
{
    private readonly ExcelMapperBase excelMapper;

    private Spreadsheet spreadsheet;

    private Row? currentRow;

    private ExcelIterator(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        this.excelMapper = excelMapper;
        this.spreadsheet = spreadsheet;
        currentRow = spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex != null && r.RowIndex.Value > options.HeaderLength)
            .MinBy(r => r.RowIndex?.Value);
    }

    internal static ExcelIterator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelIterator<T>(spreadsheet, excelMapper, options);
    }
    
    public bool Read()
    {
        if (currentRow is null)
        {
            return false;
        }

        var rows = spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex != null).ToList();
        
        currentRow = rows.SingleOrDefault(r => r.RowIndex!.Value == currentRow.RowIndex!.Value + 1);
        return currentRow is not null && rows.Any(r => r.RowIndex!.Value == currentRow.RowIndex!.Value + 1);
    }

    public T GetCurrentRow()
    {
        var cells = currentRow!.Elements<Cell>().ToList();

        var instanceType = typeof(T);
        var properties = instanceType.GetProperties();
        var instance = (T)Activator.CreateInstance(instanceType)!;

        foreach (var propertyInfo in properties)
        {
            var type = propertyInfo.PropertyType;
            var valueString = Helper.GetCellValue<T>(cells, excelMapper, propertyInfo.Name);

            propertyInfo.SetValue(instance, Helper.GetValueOrDefault(type, valueString));
        }

        return instance;
    }

    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}