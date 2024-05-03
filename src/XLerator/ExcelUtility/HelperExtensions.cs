using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal static class HelperExtensions
{
    internal static void SaveRowToSpreadsheet<T>(this SpreadsheetDocument spreadsheetDocument, StringValue sheetId, uint rowIndex, ExcelRow<T> row) where T : class
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