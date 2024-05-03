﻿using DocumentFormat.OpenXml.Packaging;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.Reader;

internal class ExcelReader : IDisposable, IExcelReader
{
    private readonly ExcelMapperBase excelMapper;
    
    private SpreadsheetDocument spreadsheet = null!;

    private ExcelReader(ExcelMapperBase excelMapper)
    {
        this.excelMapper = excelMapper;
    }

    internal static ExcelReader Create(XLeratorOptions options, ExcelMapperBase excelMapper)
    {
        var reader = new ExcelReader(excelMapper);
        reader.spreadsheet = SpreadsheetDocument.Open(options.GetFilePath(), false);
        return reader;
    }
    
    public void Dispose()
    {
        spreadsheet.Dispose();
        spreadsheet = null!;
    }
}