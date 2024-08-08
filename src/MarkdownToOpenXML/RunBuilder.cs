using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML
{
    internal class RunBuilder
    {
        private readonly PatternMatcher Bold;
        private readonly PatternMatcher Hyperlinks;
        private readonly PatternMatcher Hyperlinks_Text;
        private readonly PatternMatcher Italic;
        private readonly string md;
        public Paragraph para;
        private Run run;
        private readonly PatternMatcher Tab;

        private readonly Ranges<int> Tokens = new Ranges<int>();
        private readonly PatternMatcher Underline;

        public RunBuilder(string md, Paragraph para)
        {
            this.para = para;
            this.md = md;

            PatternMatcher.Pattern it = MD2OXML.ExtendedMode ? PatternMatcher.Pattern.Grave : PatternMatcher.Pattern.Asterisk;
            Italic = new PatternMatcher(it);
            Italic.FindMatches(md, ref Tokens);

            Bold = new PatternMatcher(PatternMatcher.Pattern.DblAsterisk);
            Bold.FindMatches(md, ref Tokens);

            if (MD2OXML.ExtendedMode)
            {
                Underline = new PatternMatcher(PatternMatcher.Pattern.Underscore);
                Underline.FindMatches(md, ref Tokens);
            }

            Tab = new PatternMatcher(PatternMatcher.Pattern.Tab);
            Tab.FindMatches(md, ref Tokens);

            Hyperlinks = new PatternMatcher(PatternMatcher.Pattern.Hyperlink);
            Hyperlinks.FindMatches(md, ref Tokens);

            Hyperlinks_Text = new PatternMatcher(PatternMatcher.Pattern.Hyperlink_Text);
            Hyperlinks_Text.FindMatches(md, ref Tokens);

            GenerateRuns();
        }

        private bool PatternsHaveMatches()
        {
            return Bold.HasMatches() || Italic.HasMatches() || Underline.HasMatches() || Hyperlinks.HasMatches() ||
                   Hyperlinks_Text.HasMatches() || Tab.HasMatches();
        }

        private void GenerateRuns()
        {
            // Calculate positions of all tokens and use this to set 
            // run styles when iterating through the string

            // in the same calculation note down location of tokens 
            // so they can be ignored when loading strings into the buffer

            if (!PatternsHaveMatches())
            {
                run = new Run();
                run.Append(new Text(md) { Space = SpaceProcessingModeValues.Preserve });
                para.Append(run);
            }
            else
            {
                int pos = 0;
                string buffer = "";

                // This needs optimizing so it builds a string buffer before adding the run itself
                while (pos < md.Length)
                {
                    if (!Tokens.ContainsValue(pos))
                    {
                        buffer += md.Substring(pos, 1);
                    }
                    else if (buffer.Length > 0)
                    {
                        run = new Run();
                        RunProperties rPr = new RunProperties();

                        Bold.SetFlagFor(pos - 1);
                        Italic.SetFlagFor(pos - 1);
                        Underline.SetFlagFor(pos - 1);

                        if (Bold.Flag)
                        {
                            rPr.Append(new Bold { Val = new OnOffValue(true) });
                        }

                        if (Italic.Flag)
                        {
                            rPr.Append(new Italic());
                        }

                        if (Underline.Flag)
                        {
                            rPr.Append(new Underline { Val = UnderlineValues.Single });
                        }

                        run.Append(rPr);
                        run.Append(new Text(buffer) { Space = SpaceProcessingModeValues.Preserve });

                        if (Tab.ContainsValue(pos))
                        {
                            run.Append(new TabChar());
                        }

                        para.Append(run);
                        buffer = "";
                    }

                    pos++;
                }
            }
        }
    }
}