using System.Reflection;
using XLerator.Attributes;
using XLerator.ExcelUtility.ExcelEditing;
using XLerator.ExcelUtility.ExcelReading.Iterator;
using XLerator.ExcelUtility.ExcelReading.Reader;
using XLerator.Mappings;

namespace XLerator.ExcelUtility;

public class XLeratorSpreadsheet<TClass> : IDisposable, IAsyncDisposable where TClass : class
{
    private readonly XLeratorOptions options;
    private readonly Spreadsheet spreadsheet;

    private readonly ExcelReader<TClass> reader;
    private readonly ExcelIterator<TClass> iterator;
    private readonly ExcelEditor<TClass> editor;

    public IExcelReader<TClass> Reader => reader;
    
    public IExcelIterator<TClass> Iterator => iterator;
    
    public IExcelEditor<TClass> Editor => editor;

    private XLeratorSpreadsheet(XLeratorOptions options, bool enableEditing)
    {
        this.options = options;
        spreadsheet = Spreadsheet.Open(options, enableEditing);

        var mapper = CreateMapper(typeof(TClass));
        
        editor = ExcelEditor<TClass>.Create(options, mapper, spreadsheet);
        reader = ExcelReader<TClass>.Create(options, mapper, spreadsheet);
        iterator = ExcelIterator<TClass>.Create(options, mapper, spreadsheet);
    }
    
    private XLeratorSpreadsheet(XLeratorOptions options, Spreadsheet spreadsheet, ExcelMapperBase excelMapper)
    {
        this.options = options;
        this.spreadsheet = spreadsheet;
        
        editor = ExcelEditor<TClass>.Create(options, excelMapper, spreadsheet);
        reader = ExcelReader<TClass>.Create(options, excelMapper, spreadsheet);
        iterator = ExcelIterator<TClass>.Create(options, excelMapper, spreadsheet);
    }

    public static XLeratorSpreadsheet<T> Open<T>(string filePath, bool enableEditing = true, bool create = false) where T : class
    {
        if (create && !File.Exists(filePath))
        {
            return Create<T>(filePath);
        }
        
        return new XLeratorSpreadsheet<T>(new XLeratorOptions
        {
            FilePath = filePath
        }, enableEditing);
    }

    public static XLeratorSpreadsheet<T> Open<T>(string filePath, string sheetName, bool enableEditing = true, bool create = false) where T : class
    {
        if (create && !File.Exists(filePath))
        {
            return Create<T>(filePath, sheetName);
        }
        
        return new XLeratorSpreadsheet<T>(new XLeratorOptions
        {
            FilePath = filePath,
            SheetName = sheetName
        }, enableEditing);
    }

    public static XLeratorSpreadsheet<T> Open<T>(XLeratorOptions options, bool enableEditing = true, bool create = false) where T : class
    {
        if (create && !File.Exists(options.FilePath))
        {
            return Create<T>(options);
        }
        
        return new XLeratorSpreadsheet<T>(options, enableEditing);
    }
    
    public static XLeratorSpreadsheet<T> Create<T>(string filePath) where T : class
    {
        var options = new XLeratorOptions { FilePath = filePath };
        var mapper = CreateMapper(typeof(T));
        
        var spreadsheet = ExcelCreator<T>.CreateExcel(options, mapper);
        return new XLeratorSpreadsheet<T>(options, spreadsheet, mapper);
    }

    public static XLeratorSpreadsheet<T> Create<T>(string filePath, string sheetName) where T : class
    {
        var options = new XLeratorOptions { FilePath = filePath, SheetName = sheetName };
        var mapper = CreateMapper(typeof(T));
        
        var spreadsheet = ExcelCreator<T>.CreateExcel(options, mapper);
        return new XLeratorSpreadsheet<T>(options, spreadsheet, mapper);
    }

    public static XLeratorSpreadsheet<T> Create<T>(XLeratorOptions options) where T : class
    {
        var mapper = CreateMapper(typeof(T));
        
        var spreadsheet = ExcelCreator<T>.CreateExcel(options, mapper);
        return new XLeratorSpreadsheet<T>(options, spreadsheet, mapper);
    }
    
    private static ExcelMapperBase CreateMapper(Type type)
    {
        if (type.IsDefined(typeof(HeaderedExcel)))
        {
            return HeaderExcelMapper.CreateFrom(type);
        }
        
        return IndexedExcelMapper.CreateFrom(type);
    }

    public void Dispose()
    {
        spreadsheet.Dispose();
        reader.Dispose();
        iterator.Dispose();
        editor.Dispose();
    }

    public async ValueTask DisposeAsync()
    {
        await CastAndDispose(spreadsheet);
        await CastAndDispose(reader);
        await CastAndDispose(iterator);
        await CastAndDispose(editor);

        return;

        static async ValueTask CastAndDispose(IDisposable resource)
        {
            if (resource is IAsyncDisposable resourceAsyncDisposable)
                await resourceAsyncDisposable.DisposeAsync();
            else
                resource.Dispose();
        }
    }
}