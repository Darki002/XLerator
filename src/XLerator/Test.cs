using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator;

public static class Test
{
    public static void Run(string fileName)
    {
            // Create a spreadsheet document by using the file name.  
            SpreadsheetDocument spreadsheetDocument =  
                 SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook);  

            // Add a WorkbookPart and Workbook objects.  
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();  
            workbookpart.Workbook = new Workbook();  

            // Add a WorksheetPart.  
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();  

            // Create Worksheet and SheetData objects.  
            worksheetPart.Worksheet = new Worksheet(new SheetData());  

            // Add a Sheets object.  
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook  
                .AppendChild<Sheets>(new Sheets());  

            // Append the new worksheet named "mySheet" and associate it   
            // with the workbook.  
            Sheet sheet = new Sheet()  
            {  
                Id = spreadsheetDocument.WorkbookPart  
                    .GetIdOfPart(worksheetPart),  
                SheetId = 1,  
                Name = "mySheet"  
            };  
            sheets.Append(sheet);  

            // Get the sheetData cell table.  
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();  

            // Add a row to the cell table.  
            Row row;  
            row = new Row() { RowIndex = 1 };  
            sheetData.Append(row);  

            // Add the cell to the cell table at A1.  
            Cell refCell = null;  
            Cell newCell = new Cell() { CellReference = "A1" };  
            row.InsertBefore(newCell, refCell);  

            // Set the cell value to be a numeric value of 123.  
            newCell.CellValue = new CellValue("123");  
            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);  

            // Close the document.  
            spreadsheetDocument.Save();
    }
}