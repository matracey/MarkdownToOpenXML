using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MarkdownToOpenXML;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string markdown = File.ReadAllText(@".\demo.txt");
            string saveTo = @".\demo.docx";

            MD2OXML.CreateDocX(markdown, saveTo);
            Process.Start(saveTo);
        }
    }
}