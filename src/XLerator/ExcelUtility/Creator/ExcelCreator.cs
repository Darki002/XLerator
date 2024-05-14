using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.Editor;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private const uint RowIndex = 1;
    
    private readonly ExcelMapperBase excelMapper;
    private readonly XLeratorOptions xLeratorOptions;
    
    private ExcelCreator(XLeratorOptions xLeratorOptions, ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        this.xLeratorOptions = xLeratorOptions;
    }

    internal static IExcelCreator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
       return new ExcelCreator<T>(options, excelMapper);
    }
    
    public IExcelEditor<T> CreateExcel(bool addHeader)
    {
       var spreadsheet = Spreadsheet.Create(xLeratorOptions);
        
        if (addHeader)
        {
            try
            {
                AddHeader(spreadsheet);
            }
            catch
            {
                spreadsheet.Save();
                spreadsheet.Dispose();
                throw;
            }
        }
        spreadsheet.Save();
        spreadsheet.Dispose();

        return ExcelEditor<T>.CreateFrom(spreadsheet, excelMapper, xLeratorOptions);
    }
    
    private void AddHeader(Spreadsheet spreadsheet)
    {
       var row = ExcelHeader<T>.CreateFrom(RowIndex, excelMapper);
       var dataRow = new Row { RowIndex = RowIndex };
        
       Cell? lastCell = null;
       foreach (var cell in row)
       {
           var newCell = cell.ToCell();
           dataRow.InsertAfter(newCell, lastCell);
           lastCell = newCell;
       }
        
       spreadsheet.AppendRow(dataRow);
       spreadsheet.SaveWorksheet();
    }
}