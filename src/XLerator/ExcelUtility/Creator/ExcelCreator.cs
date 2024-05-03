using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility.Editor;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Creator;

internal class ExcelCreator<T> : IExcelCreator<T> where T : class
{
    private const uint RowIndex = 0;
    
    private readonly ExcelMapperBase excelMapper;
    private readonly XLeratorOptions xLeratorOptions;
    
    private StringValue sheetId = null!;
    
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
        using (var spreadsheetDocument = CreateFile())
        {
            if (addHeader)
            {
                AddHeader(spreadsheetDocument);
            }
        }
        
        return ExcelEditor<T>.Create(xLeratorOptions, excelMapper);
    }

    private SpreadsheetDocument CreateFile()
    {
        var spreadsheet = SpreadsheetDocument.Create(xLeratorOptions.GetFilePath(), SpreadsheetDocumentType.Workbook);
        
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());
        
        var sheets = spreadsheet.WorkbookPart?.Workbook.AppendChild(new Sheets());
        sheetId = spreadsheet.WorkbookPart?.GetIdOfPart(worksheetPart);
        var sheet = new Sheet
        {
            Id = sheetId,
            Name = xLeratorOptions.GetSheetNameOrDefault()
        };
        
        sheets?.Append(sheet);
        spreadsheet.Save();
        return spreadsheet;
    }
    
    private void AddHeader(SpreadsheetDocument spreadsheetDocument)
    {
       var propertyInfos = typeof(T).GetProperties();

        var row = new List<ExcelCell<string>>();
        foreach (var propertyInfo in propertyInfos)
        {
            var header = excelMapper.GetHeaderFor(propertyInfo.Name);
            var col = excelMapper.GetColumnFor(propertyInfo.Name);
            if(header is null || col is null) continue;
            
            row.Add(new ExcelCell<string>(col, RowIndex, header));
        }
        spreadsheetDocument.SaveRowToSpreadsheet(sheetId, 0, row);
    }
}