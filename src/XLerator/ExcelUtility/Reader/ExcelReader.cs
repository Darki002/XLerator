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

        if (row is null)
        {
            throw new ArgumentException($"Row with index {rowIndex} does not exist.");
        }
        
        var cells = row.Elements<Cell>().ToList();
    }

    public List<T> GetRows(int lowerBound, int upperBound)
    {
        ThrowHelper.IfInvalidRowIndex(lowerBound);
        ThrowHelper.IfInvalidRowIndex(upperBound);

        throw new NotImplementedException();
    }

    public List<T> GetRange(int column, int lowerRow, int upperRow)
    {
        ThrowHelper.IfInvalidRowIndex(lowerRow);
        ThrowHelper.IfInvalidRowIndex(upperRow);

        throw new NotImplementedException();
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}