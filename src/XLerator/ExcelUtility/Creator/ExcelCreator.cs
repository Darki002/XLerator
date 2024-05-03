using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;
    private StringValue sheetId = null!;

    private uint currentRow;
    
    private ExcelCreator(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
        currentRow = 0;
    }

    internal static ExcelCreator<T> Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var reader = new ExcelCreator<T>(excelMapper);
        reader.SetUpSpreadsheet(options);
        
        return reader;
    }

    private void SetUpSpreadsheet(XLeratorOptions options)
    {
        spreadsheet = SpreadsheetDocument.Create(options.FilePath, SpreadsheetDocumentType.Workbook);
        
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = spreadsheet.WorkbookPart?.Workbook.AppendChild(new Sheets());
        sheetId = spreadsheet.WorkbookPart?.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            Name = options.WorkbookName
        };
        
        sheets?.Append(sheet);
        spreadsheet.Save();
    }
    
    public void CreateHeader()
    {
        if (currentRow > 0)
        {
            throw new InvalidOperationException("Can not create Header after there was already data written to the spreadsheet.");
        }
        
        var propertyInfos = typeof(T).GetProperties();

        var row = new List<ExcelCell<string>>();
        foreach (var propertyInfo in propertyInfos)
        {
            var header = excelMapper.GetHeaderFor(propertyInfo.Name);
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            if(header is null || col is null) continue;
            
            row.Add(new ExcelCell<string>(col, currentRow, header));
        }
        SaveRowToSpreadsheet(row);
    }

    public void Write(T data)
    {
        
    }

    public void WriteMany(IEnumerable<T> rows)
    { 
        foreach (var row in rows)
        {
            
        }
    }

    private void SaveRowToSpreadsheet<TRow>(List<ExcelCell<TRow>> row)
    {
        var worksheetPart = (WorksheetPart?)spreadsheet.WorkbookPart?.GetPartById(sheetId!);
        if (worksheetPart is null)
        {
            throw new InvalidOperationException("The Worksheet was not initialized correctly.");
        }
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        var dataRow = new Row { RowIndex = currentRow };
        
        foreach (var data in row)
        {
            dataRow.AppendChild(data.ToCell());
        }
        
        sheetData?.AppendChild(dataRow);
        spreadsheet.Save();
        currentRow++;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}