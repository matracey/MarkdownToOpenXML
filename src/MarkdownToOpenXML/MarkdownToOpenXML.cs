using System;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML
{
    public class MD2OXML
    {
        public static bool ExtendedMode = true;
        private int lineCount;
        private string[] lines;
        private readonly string md;
        private readonly string path;

        private bool SkipNextLine;

        public MD2OXML(string md, string path)
        {
            this.md = md;
            this.path = path;
        }

        public static void CreateDocX(string md, string path)
        {
            MD2OXML inst = new MD2OXML(md, path);
            inst.run();
        }

        public void run()
        {
            Body body = new Body();
            int index = 0;

            lines = md.Split(
                new[]
                {
                    "\r\n",
                    "\n"
                },
                StringSplitOptions.None);
            lineCount = lines.Count();

            foreach (string line in lines)
            {
                if (SkipNextLine)
                {
                    index += 1;
                    SkipNextLine = !SkipNextLine;
                    continue;
                }

                ParagraphBuilder paragraph = new ParagraphBuilder(line, GetLine(index - 1), GetLine(index + 1));
                SkipNextLine = paragraph.SkipNextLine;
                body.Append(paragraph.Build());
                index += 1;
            }

            DocumentBuilder file = new DocumentBuilder(body);
            file.SaveTo(path);
        }

        private string GetLine(int n)
        {
            return inRange(n) ? lines[n] : "";
        }

        private bool inRange(int n)
        {
            return n >= 0 && n < lineCount;
        }
    }
}