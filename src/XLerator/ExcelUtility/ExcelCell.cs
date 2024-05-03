using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility;

internal struct ExcelCell(string column, uint row, object? data = null)
{
    public object? Data { get; set; } = data;

    private string CellReference => column + row;
    
    public Cell ToCell()
    {
        if (Data is null)
        {
            throw new InvalidOperationException("Not Data to convert to CellValue");
        }
        
        return new Cell(new CellValue(Data.ToString()!))
        {
            DataType = GetValueType(),
            CellReference = CellReference,
        };
    }

    private CellValues GetValueType()
    {
        var type = Data!.GetType();
        if (type == typeof(string)) return CellValues.String;
        if (type == typeof(DateTime)) return CellValues.Date;
        if (type == typeof(bool)) return CellValues.Boolean;
        return IsNumericType(type) ? CellValues.Number : CellValues.String;
    }
    
    private static bool IsNumericType(Type type)
    {
        var typeCode = Type.GetTypeCode(type);
        return typeCode is TypeCode.Int32 or TypeCode.UInt32 or TypeCode.Int16 or TypeCode.UInt16 or TypeCode.Int64 or TypeCode.UInt64 or TypeCode.Single or TypeCode.Double or TypeCode.Decimal;
    }
}