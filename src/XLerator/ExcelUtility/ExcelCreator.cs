using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility;

internal static class ExcelCreator<T> where T : class
{
    public static Spreadsheet CreateExcel(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
       var spreadsheet = Spreadsheet.Create(options);
        
        if (options.HeaderLength > 0)
        {
            try
            {
                var index = (uint)options.HeaderLength;
                AddHeader(spreadsheet, excelMapper, index);
            }
            catch
            {
                spreadsheet.Save();
                spreadsheet.Dispose();
                throw;
            }
        }
        spreadsheet.Save();

        return spreadsheet;
    }
    
    private static void AddHeader(Spreadsheet spreadsheet, ExcelMapperBase excelMapper, uint index)
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