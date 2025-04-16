using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelEditing;

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
    
    internal static ExcelEditor<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper, Spreadsheet spreadsheet)
    {
        return new ExcelEditor<T>(spreadsheet, excelMapper, options);
    }
    
    public void Write(T data)
    {
        try
        {
            var rowIndex = GetNewRowIndex();
            var row = ExcelData<T>.CreateFrom(data, rowIndex, excelMapper);
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

    public void Update(int rowIndex, T data)
    {
        ThrowHelper.IfInvalidRowIndex(rowIndex);
        rowIndex += options.HeaderLength;
        try
        {
            var row = ExcelData<T>.CreateFrom(data, (uint)rowIndex, excelMapper);
            AddRow(row);
            spreadsheet.Save();
        }
        catch
        {
            spreadsheet.Save();
            throw;
        }
    }

    internal void WriteRows(IEnumerable<T> data)
    {
        try
        {
            foreach (var rowData in data)
            {
                var rowIndex = GetNewRowIndex();
                var row = ExcelData<T>.CreateFrom(rowData, rowIndex, excelMapper);
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

    private uint GetNewRowIndex()
    {
        var lastRow = spreadsheet.LastRowOrDefault();
        var oldIndex = lastRow?.RowIndex ?? 0u;
        return oldIndex + 1;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}