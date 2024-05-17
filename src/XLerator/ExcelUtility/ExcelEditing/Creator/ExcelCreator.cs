using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.ExcelEditing.Editor;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelEditing.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    private readonly XLeratorOptions options;
    
    private ExcelCreator(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        this.options = options;
    }

    internal static IExcelCreator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
       return new ExcelCreator<T>(options, excelMapper);
    }
    
    public IExcelEditor<T> CreateExcel()
    {
       var spreadsheet = Spreadsheet.Create(options);
        
        if (options.HeaderLength > 0)
        {
            try
            {
                var index = (uint)options.HeaderLength;
                AddHeader(spreadsheet, index);
            }
            catch
            {
                spreadsheet.Save();
                spreadsheet.Dispose();
                throw;
            }
        }
        spreadsheet.Save();

        return ExcelEditor<T>.CreateFrom(spreadsheet, excelMapper, options);
    }
    
    private void AddHeader(Spreadsheet spreadsheet, uint index)
    {
       var row = ExcelHeader<T>.CreateFrom(index, excelMapper);
       var dataRow = new Row { RowIndex = index };
        
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