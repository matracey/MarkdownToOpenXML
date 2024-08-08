using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML;

internal class RunBuilder
{
    private readonly PatternMatcher _bold;
    private readonly PatternMatcher _hyperlinks;
    private readonly PatternMatcher _hyperlinksText;
    private readonly PatternMatcher _italic;
    private readonly string _md;
    public readonly Paragraph Para;
    private Run _run = null!;
    private readonly PatternMatcher _tab;

    private readonly Ranges<int> _tokens = new();
    private readonly PatternMatcher? _underline;

    public RunBuilder(string md, Paragraph para)
    {
        this.Para = para;
        this._md = md;

        PatternMatcher.Pattern it = MarkdownToOpenXml.ExtendedMode ? PatternMatcher.Pattern.Grave : PatternMatcher.Pattern.Asterisk;
        _italic = new PatternMatcher(it);
        _italic.FindMatches(md, ref _tokens);

        _bold = new PatternMatcher(PatternMatcher.Pattern.DblAsterisk);
        _bold.FindMatches(md, ref _tokens);

        if (MarkdownToOpenXml.ExtendedMode)
        {
            _underline = new PatternMatcher(PatternMatcher.Pattern.Underscore);
            _underline.FindMatches(md, ref _tokens);
        }

        _tab = new PatternMatcher(PatternMatcher.Pattern.Tab);
        _tab.FindMatches(md, ref _tokens);

        _hyperlinks = new PatternMatcher(PatternMatcher.Pattern.Hyperlink);
        _hyperlinks.FindMatches(md, ref _tokens);

        _hyperlinksText = new PatternMatcher(PatternMatcher.Pattern.HyperlinkText);
        _hyperlinksText.FindMatches(md, ref _tokens);

        GenerateRuns();
    }

    private bool PatternsHaveMatches()
    {
        return _bold.HasMatches() || _italic.HasMatches() || (_underline != null && _underline.HasMatches()) || _hyperlinks.HasMatches() ||
               _hyperlinksText.HasMatches() || _tab.HasMatches();
    }

    private void GenerateRuns()
    {
        // Calculate positions of all tokens and use this to set      // run styles when iterating through the string

        // in the same calculation note down location of tokens      // so they can be ignored when loading strings into the buffer

        if (!PatternsHaveMatches())
        {
            _run = new Run();
            _run.Append(new Text(_md) { Space = SpaceProcessingModeValues.Preserve });
            Para.Append(_run);
        }
        else
        {
            int pos = 0;
            string buffer = "";

            // This needs optimizing, so it builds a string buffer before adding the run itself
            while (pos < _md.Length)
            {
                if (!_tokens.ContainsValue(pos))
                {
                    buffer += _md.Substring(pos, 1);
                }
                else if (buffer.Length > 0)
                {
                    _run = new Run();
                    RunProperties rPr = new();

                    _bold.SetFlagFor(pos - 1);
                    _italic.SetFlagFor(pos - 1);
                    _underline?.SetFlagFor(pos - 1);

                    if (_bold.Flag)
                    {
                        rPr.Append(new Bold { Val = new OnOffValue(true) });
                    }

                    if (_italic.Flag)
                    {
                        rPr.Append(new Italic());
                    }

                    if (_underline?.Flag ?? false)
                    {
                        rPr.Append(new Underline { Val = UnderlineValues.Single });
                    }

                    _run.Append(rPr);
                    _run.Append(new Text(buffer) { Space = SpaceProcessingModeValues.Preserve });

                    if (_tab.ContainsValue(pos))
                    {
                        _run.Append(new TabChar());
                    }

                    Para.Append(_run);
                    buffer = "";
                }

                pos++;
            }
        }
    }
}