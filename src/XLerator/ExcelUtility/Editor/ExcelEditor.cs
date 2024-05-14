using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal class ExcelEditor<T> : IExcelEditor<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;

    private readonly XLeratorOptions options;
    
    private Spreadsheet spreadsheet;
    
    private ExcelEditor(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        this.excelMapper = excelMapper;
        this.options = options;
        this.spreadsheet = spreadsheet;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, true);
        return new ExcelEditor<T>(spreadsheet, excelMapper, options);
    }
    
    internal static ExcelEditor<T> CreateFrom(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, XLeratorOptions options)
    {
        return new ExcelEditor<T>(spreadsheet, excelMapper, options);
    }
    
    public void Write(T data)
    {
        try
        {
            var row = CreateNewRow(data);
            AddRow(row);
            spreadsheet.Save();
        }
        catch
        {
            spreadsheet.Save();
            throw;
        }
    }

    public void WriteMany(IEnumerable<T> data)
    {
        try
        {
            foreach (var rowData in data)
            {
                var row = CreateNewRow(rowData);
                AddRow(row);
            }
            spreadsheet.Save();
        }
        catch
        {
            spreadsheet.Save();
            throw;
        }
    }

    private ExcelData<T> CreateNewRow(T data)
    {
        var lastRow = spreadsheet.LastRowOrDefault();
        var index = lastRow?.RowIndex ?? 0u;
        return ExcelData<T>.CreateFrom(data, index + 1, excelMapper);
    }

    private void AddRow(ExcelData<T> row)
    {
        var dataRow = new Row { RowIndex = row.RowIndex };
        
        Cell? lastCell = null;
        foreach (var cell in row)
        {
            var newCell = cell.ToCell();
            dataRow.InsertAfter(newCell, lastCell);
            lastCell = newCell;
        }
        spreadsheet.AppendRow(dataRow);
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}