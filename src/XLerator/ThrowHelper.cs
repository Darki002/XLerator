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
}