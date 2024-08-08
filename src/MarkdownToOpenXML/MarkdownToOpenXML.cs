using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML;

public class MarkdownToOpenXml
{
    public static bool ExtendedMode { get; set; } = true;

    private int _lineCount;

    private string[] _lines = Array.Empty<string>();
    private readonly string _md;
    private readonly string _path;

    private bool _skipNextLine;

    public MarkdownToOpenXml(string md, string path)
    {
        this._md = md;
        this._path = path;
    }

    public static void CreateDocX(string md, string path)
    {
        MarkdownToOpenXml inst = new(md, path);
        inst.Run();
    }

    public void Run()
    {
        Body body = new();
        int index = 0;

        _lines = _md.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
        _lineCount = _lines.Length;

        foreach (string line in _lines)
        {
            if (_skipNextLine)
            {
                index += 1;
                _skipNextLine = !_skipNextLine;
                continue;
            }

            ParagraphBuilder paragraph = new(line, GetLine(index + 1));
            _skipNextLine = paragraph.SkipNextLine;
            body.Append(paragraph.Build());
            index += 1;
        }

        DocumentBuilder file = new(body);
        file.SaveTo(_path);
    }

    private string GetLine(int n)
    {
        return InRange(n) ? _lines[n] : "";
    }

    private bool InRange(int n)
    {
        return n >= 0 && n < _lineCount;
    }
}