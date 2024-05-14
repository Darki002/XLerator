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
            var lastRow = spreadsheet.LastRowOrDefault();
            var index = lastRow?.RowIndex ?? 1;
            
            var row = ExcelData<T>.CreateFrom(data, index, excelMapper);
            AddRow(row, index);
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
                var lastRow = spreadsheet.LastRowOrDefault();
                var index = lastRow?.RowIndex ?? 1u;
                
                var row = ExcelData<T>.CreateFrom(rowData, index, excelMapper);
                AddRow(row, index);
            }
            spreadsheet.Save();
        }
        catch
        {
            spreadsheet.Save();
            throw;
        }
    }

    private void AddRow(ExcelData<T> row, uint index)
    {
        var dataRow = new Row { RowIndex = index };
        
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