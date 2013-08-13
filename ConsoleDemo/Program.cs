
namespace ConsoleDemo
{
    using System;
    using System.IO;
    using MarkdownToOpenXML;

    class ConsoleDemo
    {
        static void Main(string[] args)
        {
            string mdPath = "";

            do
            {
                Console.WriteLine("Please provide a path to an existing Markdown file:");
            }
            while (!File.Exists(mdPath = Console.ReadLine()));

            string md = System.IO.File.ReadAllText(mdPath);
            
            Console.WriteLine("");
            Console.WriteLine("Please provide the path for the output destination:");
            string outPath = Console.ReadLine();
            MD2OXML.CreateDocX(md, outPath);
        }
    }
}
