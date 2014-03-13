
namespace MarkdownToOpenXML
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    using MarkdownToOpenXML;

    public class MD2OXML
    {
        public static bool ExtendedMode = true;

        private bool SkipNextLine = false;
        private string[] lines;
        private int lineCount;
        private string md;
        private string path;

        public static void CreateDocX(string md, string path)
        {
            MD2OXML inst = new MD2OXML(md, path);
            inst.run();
        }

        public MD2OXML(string md, string path)
        {
            this.md = md;
            this.path = path;
        }

        public void run()
        {
            Body body = new Body();
            int index = 0;
            
            lines = md.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
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
            return (n >= 0 && n < lineCount);
        }
        
    }
}
