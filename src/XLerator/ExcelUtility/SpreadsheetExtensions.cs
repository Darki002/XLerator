using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal static class SpreadsheetExtensions
{
    internal static void SaveRowToSpreadsheet<TRow>(this SpreadsheetDocument spreadsheetDocument, StringValue sheetId, uint rowIndex, List<ExcelCell<TRow>> row)
    {
        var worksheetPart = (WorksheetPart?)spreadsheetDocument.WorkbookPart?.GetPartById(sheetId!);
        if (worksheetPart is null)
        {
            throw new InvalidOperationException("The Worksheet was not initialized correctly.");
        }
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var dataRow = new Row { RowIndex = rowIndex };
        
        foreach (var data in row)
        {
            dataRow.AppendChild(data.ToCell());
        }
        
        sheetData?.AppendChild(dataRow);
        spreadsheetDocument.Save();
    }
}