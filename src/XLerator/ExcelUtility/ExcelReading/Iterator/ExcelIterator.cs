using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading.Iterator;

// TODO Tests

internal class ExcelIterator<T> : IExcelIterator<T> where T : class
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
        currentRowIndex = 0;
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
            .Where(r => r.RowIndex > (currentRow?.RowIndex ?? 0))
            .MinBy(r => r.RowIndex?.Value);
        currentRowIndex = currentRow?.RowIndex?.Value ?? 0;
        return spreadsheet.SheetData.Elements<Row>().Any(r => r.RowIndex?.Value > currentRowIndex);
    }

    public T GetCurrentRow()
    {
        if (currentRow is null)
        {
            throw new InvalidOperationException($"No row was found to read at index {currentRowIndex}.");
        }

        var cells = currentRow.Elements<Cell>().ToList();
        var helper = new Helper<T>(spreadsheet, excelMapper);
        return helper.DeserializerFrom(cells);
    }

    public void SkipRows(int amount)
    {
        if (amount <= 0)
        {
            throw new ArgumentException($"{nameof(amount)} must be greater then zero.");
        }

        currentRow = spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex > (currentRow?.RowIndex ?? 0))
            .MinBy(r => r.RowIndex);
        currentRowIndex = currentRow?.RowIndex?.Value ?? 0;
    }

    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}