using System.Security;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility.ExcelEditing;

internal readonly struct ExcelCell(string column, uint row, object data)
{
    private object Data { get; } = data;
    
    public Cell ToCell()
    {
        var text = Data.ToString();
        if (text is null)
        {
            throw new InvalidOperationException("Not Data to convert to CellValue");
        }

        var cell = new Cell
        {
            DataType = new EnumValue<CellValues>(GetValueType()),
            CellReference = GetCellReference(),
            CellValue = new CellValue(SecurityElement.Escape(text))
        };
        return cell;
    }

    private CellValues GetValueType()
    {
        var type = Data.GetType();
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

    private string GetCellReference() => $"{column}{row}";
}