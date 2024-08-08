namespace MarkdownToOpenXML;

public class Ranges<T> where T : IComparable<T>
{
    private readonly List<Range<T>> _rangeList = new();

    public void Add(Range<T> range)
    {
        _rangeList.Add(range);
    }

    public int Count()
    {
        return _rangeList.Count;
    }

    public Boolean ContainsValue(T value)
    {
        return _rangeList.Any(range => range.ContainsValue(value));
    }
}