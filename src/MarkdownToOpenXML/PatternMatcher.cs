﻿using System.Text.RegularExpressions;

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
        Hyperlink_Text,
        Hyperlink
    }

    private static readonly Dictionary<Pattern, Regex> Patterns = new Dictionary<Pattern, Regex>
    {
        { Pattern.DblAsterisk, new Regex("(?<!\\*)(\\*\\*)([^\\ +].+?)(\\*\\*)") },
        { Pattern.Asterisk, new Regex("(?<!\\*)[^\\*]?(\\*)([^\\*].+?)(\\*)[^\\*]") },
        { Pattern.Grave, new Regex("(?<!`)[^`]?(`)([^`].+?)(\\`)[^`]?") },
        { Pattern.Underscore, new Regex("(?<!_)[^_]?(_)([^_].+?)(_)[^_]?") },
        { Pattern.Tab, new Regex(@"(\\t|[\\ ]{4})") },
        {
            Pattern.Hyperlink_Text,
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
    public Ranges<int> matches = new Ranges<int>();

    private readonly Pattern pattern;
    public Regex regex;

    public PatternMatcher(Pattern pattern)
    {
        this.pattern = pattern;
        regex = Patterns[pattern];
    }

    public void FindMatches(string md, ref Ranges<int> Tokens)
    {
        if (pattern == Pattern.DblAsterisk || pattern == Pattern.Asterisk || pattern == Pattern.Grave || pattern == Pattern.Underscore)
        {
            MatchCollection mc = regex.Matches(md);

            foreach (Match m in mc)
            {
                int sToken = m.Groups[1].Index;
                int match = m.Groups[2].Index;
                int eToken = m.Groups[3].Index;
                int endStr = m.Groups[3].Index + m.Groups[3].Length;

                Tokens.add(new Range<int> { Minimum = sToken, Maximum = match - 1 });

                matches.add(new Range<int> { Minimum = match, Maximum = eToken - 1 });

                Tokens.add(new Range<int> { Minimum = eToken, Maximum = endStr - 1 });
            }
        }
        else if (pattern == Pattern.Tab)
        {
            MatchCollection mc = regex.Matches(md);

            foreach (Match m in mc)
            {
                matches.add(new Range<int> { Minimum = m.Index, Maximum = m.Index });

                Tokens.add(new Range<int> { Minimum = m.Index, Maximum = m.Index + m.Length - 1 });
            }
        }
        else if (pattern == Pattern.Hyperlink_Text || pattern == Pattern.Hyperlink)
        {
        }
        else
        {
            throw new InvalidOperationException("FindMatches not implemented for this pattern");
        }
    }

    public bool HasMatches()
    {
        return matches.Count() > 0;
    }

    public bool ContainsValue(int num)
    {
        return matches.ContainsValue(num);
    }

    public void SetFlagFor(int pos)
    {
        if (ContainsValue(pos))
        {
            Flag = true;
        }
        else
        {
            Flag = false;
        }
    }
}