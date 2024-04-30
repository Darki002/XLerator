namespace XLerator;

public record ExcelMap(int Index, string HeaderName)
{
    public override int GetHashCode()
    {
        return Index.GetHashCode();
    }

    public virtual bool Equals(ExcelMap? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return Index == other.Index;
    }
}