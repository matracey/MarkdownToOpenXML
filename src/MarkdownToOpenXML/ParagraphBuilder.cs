using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML;

internal class ParagraphBuilder
{
    private string _current;
    private readonly string _next;

    private readonly Paragraph _para = new();
    private readonly ParagraphProperties _prop = new();
    public bool SkipNextLine;

    public ParagraphBuilder(string current, string next)
    {
        this._current = current;
        this._next = next;
    }

    public Paragraph Build()
    {
        if (MarkdownToOpenXml.ExtendedMode)
        {
            DoAlignment();
        }

        DoHeaders();
        DoNumberedLists();

        _para.Append(_prop);
        RunBuilder run = new(_current, _para);
        return run.Para;
    }

    private void DoAlignment()
    {
        Dictionary<JustificationValues, Match> alignment = new()
        {
            { JustificationValues.Center, Regex.Match(_current, @"^><") }, { JustificationValues.Left, Regex.Match(_current, @"^<<") }, { JustificationValues.Right, Regex.Match(_current, @"^>>") },
            { JustificationValues.Distribute, Regex.Match(_current, @"^<>") }
        };

        foreach (KeyValuePair<JustificationValues, Match> match in alignment.Where(match => match.Value.Success))
        {
            _prop.Append(new Justification { Val = match.Key });
            _current = _current[2..];
            break;
        }
    }

    private void DoHeaders()
    {
        int headerLevel = _current.TakeWhile(x => x == '#').Count();

        if (headerLevel > 0)
        {
            _current = _current.TrimStart('#').TrimEnd('#').Trim();
        }
        else
        {
            String sTest = Regex.Replace(_next, @"\w", "");
            if (Regex.Match(sTest, @"[=]{2,}").Success)
            {
                headerLevel = 1;
                SkipNextLine = true;
            }

            if (Regex.Match(sTest, @"[-]{2,}").Success)
            {
                headerLevel = 2;
                SkipNextLine = true;
            }
        }

        if (headerLevel > 0)
        {
            _prop.Append(new ParagraphStyleId { Val = "Heading" + headerLevel });
        }
    }

    private void DoNumberedLists()
    {
        Match numberedList = Regex.Match(_current, @"^\\d\\.");

        // Set Paragraph Styles
        if (!numberedList.Success)
        {
            return;
        }

        // Doesn't work currently, needs NumberingDefinitions adding in filecreation.cs
        _current = _current[2..];
        NumberingProperties nPr = new(
            new NumberingLevelReference { Val = 0 },
            new NumberingId { Val = 1 }
        );

        _prop.Append(nPr);
    }
}