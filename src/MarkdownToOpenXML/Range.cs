namespace MarkdownToOpenXML;

public class Range<T> where T : IComparable<T>
{
    public T? Minimum { get; set; }
    public T? Maximum { get; set; }

    public override string ToString() { return $"[{Minimum} - {Maximum}]"; }

    public Boolean IsValid() { return Minimum != null && Minimum.CompareTo(Maximum) <= 0; }

    public Boolean ContainsValue(T value)
    {
        return Minimum != null && Minimum.CompareTo(value) <= 0 && value.CompareTo(Maximum) <= 0;
    }
}