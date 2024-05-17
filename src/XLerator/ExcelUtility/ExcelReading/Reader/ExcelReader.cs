using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading.Reader;

internal partial class ExcelReader<T> : IExcelReader<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;

    private readonly XLeratorOptions options;

    private Spreadsheet spreadsheet;

    private ExcelReader(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        this.spreadsheet = spreadsheet;
        this.excelMapper = excelMapper;
        this.options = options;
    }

    internal static ExcelReader<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, false);
        return new ExcelReader<T>(spreadsheet, excelMapper, options);
    }

    public T GetRow(int rowIndex)
    {
        ThrowHelper.IfInvalidRowIndex(rowIndex);

        rowIndex = AddDefaultConstants(rowIndex);
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
            var valueString = Helper.GetCellValue<T>(cells, excelMapper, propertyInfo.Name);

            propertyInfo.SetValue(instance, Helper.GetValueOrDefault(type, valueString));
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

        lowerBound = AddDefaultConstants(lowerBound);
        upperBound = AddDefaultConstants(lowerBound);
        return spreadsheet.SheetData.Elements<Row>()
            .Where(r => r.RowIndex != null && r.RowIndex >= lowerBound && r.RowIndex < upperBound)
            .Select(row => GetRow((int)row.RowIndex?.Value!))
            .ToList();
    }

    private int AddDefaultConstants(int index) => index + options.HeaderLength + 1;

    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}