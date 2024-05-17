namespace XLerator;

internal static class ThrowHelper
{
    public static void IfInvalidRowIndex(int rowIndex)
    {
        if (rowIndex <= 0)
        {
            throw new ArgumentException("Row Index must be greater then 0.");
        }
    }

    public static void ThrowIfNull(object? obj, string message)
    {
        if (obj is null)
        {
            throw new ArgumentException(message);
        }
    }
}