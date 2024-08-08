using System.Text.RegularExpressions;

namespace MarkdownToOpenXML;

internal class PatternMatcher
{
    public enum Pattern
    {
        DblAsterisk,
        Asterisk,
        Grave,
        Underscore,
        Tab,
        HyperlinkText,
        Hyperlink
    }

    private static readonly Dictionary<Pattern, Regex> Patterns = new()
    {
        { Pattern.DblAsterisk, new Regex(@"(?<!\*)(\*\*)([^\ +].+?)(\*\*)") },
        { Pattern.Asterisk, new Regex(@"(?<!\*)[^\*]?(\*)([^\*].+?)(\*)[^\*]") },
        { Pattern.Grave, new Regex(@"(?<!`)[^`]?(`)([^`].+?)(\`)[^`]?") },
        { Pattern.Underscore, new Regex(@"(?<!_)[^_]?(_)([^_].+?)(_)[^_]?") },
        { Pattern.Tab, new Regex(@"(\\t|[\\ ]{4})") },
        {
            Pattern.HyperlinkText,
            new Regex(
                @"\[[a-z]+\]\((?:[a-z]+://|www\.|ftp\.)[-A-Z0-9+&@#/%=~_|$?!:,.]*[A-Z0-9+&@#/%=~_|$]\)",
                RegexOptions.IgnoreCase)
        },
        {
            Pattern.Hyperlink,
            new Regex(@"\<(?:[a-z]+://|www\.|ftp\.)[-A-Z0-9+&@#/%=~_|$?!:,.]*[A-Z0-9+&@#/%=~_|$]\>", RegexOptions.IgnoreCase)
        }
    };

    public bool Flag;
    public readonly Ranges<int> Matches = new();

    private readonly Pattern _pattern;
    public readonly Regex Regex;

    public PatternMatcher(Pattern pattern)
    {
        this._pattern = pattern;
        Regex = Patterns[pattern];
    }

    public void FindMatches(string md, ref Ranges<int> tokens)
    {
        switch (_pattern)
        {
            case Pattern.DblAsterisk or Pattern.Asterisk or Pattern.Grave or Pattern.Underscore:
                {
                    MatchCollection mc = Regex.Matches(md);

                    foreach (Match m in mc)
                    {
                        int sToken = m.Groups[1].Index;
                        int match = m.Groups[2].Index;
                        int eToken = m.Groups[3].Index;
                        int endStr = m.Groups[3].Index + m.Groups[3].Length;

                        tokens.Add(new Range<int> { Minimum = sToken, Maximum = match - 1 });

                        Matches.Add(new Range<int> { Minimum = match, Maximum = eToken - 1 });

                        tokens.Add(new Range<int> { Minimum = eToken, Maximum = endStr - 1 });
                    }

                    break;
                }
            case Pattern.Tab:
                {
                    MatchCollection mc = Regex.Matches(md);

                    foreach (Match m in mc)
                    {
                        Matches.Add(new Range<int> { Minimum = m.Index, Maximum = m.Index });

                        tokens.Add(new Range<int> { Minimum = m.Index, Maximum = m.Index + m.Length - 1 });
                    }

                    break;
                }
            case Pattern.HyperlinkText or Pattern.Hyperlink:
                break;
            default:
                throw new InvalidOperationException("FindMatches not implemented for this pattern");
        }
    }

    public bool HasMatches()
    {
        return Matches.Count() > 0;
    }

    public bool ContainsValue(int num)
    {
        return Matches.ContainsValue(num);
    }

    public void SetFlagFor(int pos)
    {
        Flag = ContainsValue(pos);
    }
}