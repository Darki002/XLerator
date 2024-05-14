using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Editor;

internal class ExcelEditor<T> : IExcelEditor<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private Spreadsheet spreadsheet;
    
    private ExcelEditor(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        this.spreadsheet = spreadsheet;
    }

    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var spreadsheet = Spreadsheet.Open(options, true);
        return new ExcelEditor<T>(spreadsheet, excelMapper);
    }
    
    internal static ExcelEditor<T> CreateFrom(Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        return new ExcelEditor<T>(spreadsheet, excelMapper);
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

    public void WriteMany(params T[] data) => WriteRows(data);
    
    public void WriteMany(IEnumerable<T> data) => WriteRows(data);

    internal void WriteRows(IEnumerable<T> data)
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