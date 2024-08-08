using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML
{
    internal class ParagraphBuilder
    {
        private string current;
        private readonly string next;

        private readonly Paragraph para = new Paragraph();
        private readonly ParagraphProperties prop = new ParagraphProperties();
        public bool SkipNextLine;

        public ParagraphBuilder(string current, string next)
        {
            this.current = current;
            this.next = next;
        }

        public Paragraph Build()
        {
            if (MD2OXML.ExtendedMode)
            {
                DoAlignment();
            }

            DoHeaders();
            DoNumberedLists();

            para.Append(prop);
            RunBuilder run = new RunBuilder(current, para);
            return run.para;
        }

        private void DoAlignment()
        {
            Dictionary<JustificationValues, Match> Alignment = new Dictionary<JustificationValues, Match>();
            Alignment.Add(JustificationValues.Center, Regex.Match(current, @"^><"));
            Alignment.Add(JustificationValues.Left, Regex.Match(current, @"^<<"));
            Alignment.Add(JustificationValues.Right, Regex.Match(current, @"^>>"));
            Alignment.Add(JustificationValues.Distribute, Regex.Match(current, @"^<>"));

            foreach (KeyValuePair<JustificationValues, Match> match in Alignment)
            {
                if (match.Value.Success)
                {
                    prop.Append(new Justification { Val = match.Key });
                    current = current.Substring(2);
                    break;
                }
            }
        }

        private void DoHeaders()
        {
            int headerLevel = current.TakeWhile(x => x == '#').Count();

            if (headerLevel > 0)
            {
                current = current.TrimStart('#').TrimEnd('#').Trim();
            }
            else
            {
                String sTest = Regex.Replace(next, @"\w", "");
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
                prop.Append(new ParagraphStyleId { Val = "Heading" + headerLevel });
            }
        }

        private void DoNumberedLists()
        {
            Match numberedList = Regex.Match(current, @"^\\d\\.");

            // Set Paragraph Styles
            if (numberedList.Success)
            {
                // Doesnt work currently, needs NumberingDefinitions adding in filecreation.cs
                current = current.Substring(2);
                NumberingProperties nPr = new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }
                );

                prop.Append(nPr);
            }
        }
    }
}