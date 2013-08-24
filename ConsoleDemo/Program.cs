
namespace ConsoleDemo
{
    using System;
    using System.IO;
    using MarkdownToOpenXML;

    class ConsoleDemo
    {
        static string inputFile = null;
        static string outputFile = null;
        static void Main(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-i":
                        if (i + 1 >= args.Length || args[i + 1].StartsWith("-"))
                        {
                            Console.WriteLine("Argument required for -i");
                            return;
                        }
                        inputFile = args[++i];
                        break;
                    case "-o":
                        if (i + 1 >= args.Length || args[i + 1].StartsWith("-"))
                        {
                            Console.WriteLine("Argument required for -o");
                            return;
                        }
                        outputFile = args[++i];
                        break;
                    default:
                        Console.WriteLine("argument {0} not understood", args[i]);
                        break;
                }
            }

            while (!File.Exists(inputFile))
            {
                Console.WriteLine("Please provide a path to an existing Markdown file:");
                inputFile = Console.ReadLine();
            }

            string md = System.IO.File.ReadAllText(inputFile);

            if (outputFile == null)
            {
                Console.WriteLine("");
                Console.WriteLine("Please provide the path for the output destination:");
                outputFile = Console.ReadLine();
            }
            MD2OXML.CreateDocX(md, outputFile);
        }
    }
}
