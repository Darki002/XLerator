using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading.Iterator;

// TODO Tests

internal class ExcelIterator<T> : IExcelIterator<T>
{
    private readonly ExcelMapperBase excelMapper;

    private Spreadsheet spreadsheet;

    private uint currentRowIndex;

    private readonly uint maxRowIndex;

    private ExcelIterator(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        this.excelMapper = excelMapper;
        this.spreadsheet = spreadsheet;
        currentRowIndex = (uint)options.HeaderLength;
        
        maxRowIndex = this.spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex != null)
            .MaxBy(r => (uint)r.RowIndex!)?
            .RowIndex?.Value ?? (uint)options.HeaderLength;
    }

    internal static ExcelIterator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelIterator<T>(spreadsheet, excelMapper, options);
    }
    
    public bool Read()
    {
        currentRowIndex++;
        return currentRowIndex < maxRowIndex;
    }

    public T GetCurrentRow()
    {
        if (currentRowIndex > maxRowIndex)
        {
            throw new InvalidOperationException("No more rows left to read in the spreadsheet.");
        }

        var row = spreadsheet.SheetData.Elements<Row>().Single(r => r.RowIndex?.Value == currentRowIndex);
        var cells = row!.Elements<Cell>().ToList();

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

        currentRowIndex += (uint)amount;
        
        var exists = spreadsheet.SheetData.Elements<Row>().Any(r => r.RowIndex?.Value == currentRowIndex);
        if (!exists)
        {
            throw new ArgumentException($"Row with the index {currentRowIndex} doesn't exist in the spreadsheet.");
        }
    }

    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}