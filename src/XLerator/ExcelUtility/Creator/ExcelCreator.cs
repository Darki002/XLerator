using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.Editor;

namespace XLerator.ExcelUtility.Creator;

internal partial class ExcelCreator<T> : IExcelCreator<T> where T : class
{
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

        return ExcelEditor<T>.CreateFrom(spreadsheet, excelMapper);
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