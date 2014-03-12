using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using MarkdownToOpenXML;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string markdown = File.ReadAllText(@".\test.txt");
            string saveTo = @".\Test.docx";
            
            MD2OXML.CreateDocX(markdown, saveTo);
            Process.Start(saveTo);
        }
    }
}
