﻿using System.Diagnostics;
using System.IO;

using MarkdownToOpenXML;

namespace Test
{
    internal class Program
    {
        private static void Main()
        {
            string markdown = File.ReadAllText(@".\demo.txt");
            string saveTo = @".\demo.docx";

            MD2OXML.CreateDocX(markdown, saveTo);
            Process.Start(saveTo);
        }
    }
}