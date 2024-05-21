using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading.Iterator;

// TODO Tests

internal class ExcelIterator<T> : IExcelIterator<T>
{
    private readonly ExcelMapperBase excelMapper;

    private readonly XLeratorOptions options;

    private Spreadsheet spreadsheet;

    private Row? currentRow;

    private uint currentRowIndex;

    private ExcelIterator(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        this.excelMapper = excelMapper;
        this.options = options;
        this.spreadsheet = spreadsheet;
        currentRow = null;
    }

    internal static ExcelIterator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelIterator<T>(spreadsheet, excelMapper, options);
    }
    
    public bool Read()
    {
        currentRow = spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex?.Value > options.HeaderLength)
            .Where(r => r.RowIndex > currentRow?.RowIndex)
            .MinBy(r => r.RowIndex?.Value);
        return currentRow is null;
    }

    public T GetCurrentRow()
    {
        if (currentRow is null)
        {
            throw new InvalidOperationException("No row was found to read.");
        }

        var cells = currentRow.Elements<Cell>().ToList();

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

    public void SkipRows(int amount)
    {
        if (amount <= 0)
        {
            throw new ArgumentException($"{nameof(amount)} must be greater then zero.");
        }
        
        currentRow = spreadsheet.SheetData.Elements<Row>()
            .SingleOrDefault(r => r.RowIndex == currentRow?.RowIndex);
    }

    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}