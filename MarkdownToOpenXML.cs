
namespace MarkdownToOpenXML
{
    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Threading.Tasks;

    using MarkdownToOpenXML;

    class MD2OXML
    {
        private static bool SkipNextLine = false;
        private static string Buffer = "";

        private static bool CustomMode;
        public static bool customMode
        {
            set { CustomMode = value; }
        }

        private enum rExKey
        {
            DblAsterisk, Asterisk, Grave, Underscore
        }

        private static Dictionary<rExKey, Regex> Patterns = new Dictionary<rExKey, Regex>()
        {
            { rExKey.DblAsterisk, new Regex("(?<!\\*)(\\*\\*)([^\\ +].+?)(\\*\\*)") },
            { rExKey.Asterisk, new Regex("(?<!\\*)[^\\*]?(\\*)([^\\*].+?)(\\*)[^\\*]") },
            { rExKey.Grave, new Regex("(?<!`)[^`]?(`)([^`].+?)(\\`)[^`]?") },
            { rExKey.Underscore, new Regex("(?<!_)[^_]?(_)([^_].+?)(_)[^_]?") },
        };

        public static void CreateDocX(string content, string docName)
        {
            List<string> md = StripDoubleLines(content);
            Body body = Convert(md);
            SaveDocX(body, docName);
        }
        
        private static List<string> StripDoubleLines(string text)
        {
            List<string> md = new List<string>();

            using (StringReader reader = new StringReader(text))
            {
                string line;
                string tmp = "";

                while ((line = reader.ReadLine()) != null)
                {
                    if (Regex.Match(line, @"^$").Success)
                    {
                        if (tmp + line != "") md.Add(tmp + line);
                        tmp = "";
                    }
                    else if (Regex.Match(line, @"^={2,}$").Success)
                    {
                        md.Add(tmp);
                        md.Add(line);
                        tmp = "";
                    }
                    else
                    {
                        tmp = tmp + line;
                    }
                }

                md.Add(tmp);
            }

            return md;
        }

        private static Body Convert(List<string> md)
        {
            Body body = new Body();
            SkipNextLine = false;
            Buffer = @"";
            int index = 0;
            
            foreach (string line in md)
            {
                if (SkipNextLine)
                {
                    index += 1;
                    SkipNextLine = !SkipNextLine;
                    continue;
                }

                Buffer = line;
                string lookahead = index + 1 != md.Count ? md[index + 1] : "";

                ParagraphProperties pPr = GenerateParagraphProperties(lookahead);
                Paragraph p = new Paragraph();
                p.Append(pPr);
                
                p = GenerateRuns(p, Buffer);
                body.Append(p);
                index += 1;
            }

            return body;
        }

        private static Ranges<int> FindMatches(Regex rx, string s, ref Ranges<int> TokenRange)
        {
            Ranges<int> ranges = new Ranges<int>();
            MatchCollection mc = rx.Matches(s);
            
            foreach (Match m in mc)
            {
                int sToken = m.Groups[1].Index;
                int match = m.Groups[2].Index;
                int eToken = m.Groups[3].Index;
                int endStr = m.Groups[3].Index + m.Groups[3].Length;

                Range<int> StartToken = new Range<int>();
                StartToken.Minimum = sToken;
                StartToken.Maximum = match - 1;
                TokenRange.add(StartToken);

                Range<int> MatchRange = new Range<int>();
                MatchRange.Minimum = match;
                MatchRange.Maximum = eToken - 1;
                ranges.add(MatchRange);

                Range<int> EndToken = new Range<int>();
                EndToken.Minimum = eToken;
                EndToken.Maximum = endStr - 1;
                TokenRange.add(EndToken);
            }

            return ranges;
        }

        private static ParagraphProperties GenerateParagraphProperties(string lookahead)
        {
            ParagraphProperties pPr = new ParagraphProperties();
            Match isHeader1 = Regex.Match(Buffer, @"^#");
            String sTest = Regex.Replace(lookahead, @"\w", "");
            Match isSetextHeader1 = Regex.Match(sTest, @"[=]{2,}");

            Match numberedList = Regex.Match(Buffer, @"^\\d\\.");

            // Set Paragraph Styles
            if (numberedList.Success)
            {
                // Doesnt work currently, needs NumberingDefinitions adding in filecreation
                Buffer = Buffer.Substring(2);
                NumberingProperties nPr = new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = 1 }
                );

                pPr.Append(nPr);
            }
            
            if (isHeader1.Success || isSetextHeader1.Success)
            {
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading1" };
                pPr.Append(paragraphStyleId1);

                if (isHeader1.Success)
                {
                    Buffer = Buffer.Substring(1);

                    // Strip the # from the end of the line if there is one
                    if (Regex.Match(Buffer, @"#$").Success) Buffer = Buffer.Substring(0, Buffer.Length - 1);

                    Buffer = Buffer.Trim();
                }

                if (isSetextHeader1.Success) SkipNextLine = true;
            }

            return pPr;
        }

        private static Paragraph GenerateRuns(Paragraph p, string Buffer)
        {
            // Calculate positions of all tokens and use this to set 
            // run styles when iterating through the string

            // in the same calculation note down location of tokens 
            // so they can be ignored when loading strings into the buffer
            
            Regex pBold, pItalic, pUnderline;

            if (!CustomMode)
            {
                pBold = Patterns[rExKey.DblAsterisk];
                pItalic = Patterns[rExKey.Asterisk];
                pUnderline = Patterns[rExKey.Underscore];
            }
            else
            {
                pBold = Patterns[rExKey.DblAsterisk];
                pItalic = Patterns[rExKey.Grave];
                pUnderline = Patterns[rExKey.Underscore];
            }

            Ranges<int> Tokens = new Ranges<int>();
            Ranges<int> Bold = FindMatches(pBold, Buffer, ref Tokens);
            Ranges<int> Italic = FindMatches(pItalic, Buffer, ref Tokens);
            Ranges<int> Underline = FindMatches(pUnderline, Buffer, ref Tokens);

            if ((Bold.Count() + Italic.Count() + Underline.Count()) == 0)
            {
                Run run = new Run();
                run.Append(new Text(Buffer)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
                p.Append(run);
            }
            else
            {
                int CurrentPosition = 0;
                Run run;

                // This needs optimizing so it builds a string buffer before adding the run itself

                while (CurrentPosition < Buffer.Length)
                {
                    if (!Tokens.ContainsValue(CurrentPosition))
                    {
                        run = new Run();

                        RunProperties rPr = new RunProperties();
                        if (Bold.ContainsValue(CurrentPosition)) rPr.Append(new Bold() { Val = new OnOffValue(true) });
                        if (Italic.ContainsValue(CurrentPosition)) rPr.Append(new Italic());
                        if (Underline.ContainsValue(CurrentPosition)) rPr.Append(new Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
                        run.Append(rPr);

                        string TextBuffer = Buffer.Substring(CurrentPosition, 1);
                        run.Append(new Text(TextBuffer)
                        {
                            Space = SpaceProcessingModeValues.Preserve
                        });
                        p.Append(run);
                    }

                    CurrentPosition++;
                }
            }
            return p;
        }

        private static void SaveDocX(Body body, String docName)
        {
            // Create a Wordprocessing document. 
            using (WordprocessingDocument package = WordprocessingDocument.Create(docName, WordprocessingDocumentType.Document))
            {
                package.AddMainDocumentPart();
                package.MainDocumentPart.Document = new Document();

                StyleDefinitionsPart styleDefinitionsPart1 = package.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
                MD2OXMLFile file = new MD2OXMLFile();
                file.GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

                package.MainDocumentPart.Document.AppendChild(body);
                package.MainDocumentPart.Document.Save();
            }
        }
        
        public class Ranges<T> where T : IComparable<T>
        {
            private List<Range<T>> rangelist = new List<Range<T>>();

            public void add(Range<T> range)
            {
                rangelist.Add(range);
            }

            public int Count()
            {
                return rangelist.Count;
            }

            public Boolean ContainsValue(T value)
            {
                foreach (Range<T> range in rangelist)
                {
                    if (range.ContainsValue(value)) return true;
                }

                return false;
            }
        }
        
        public class Range<T> where T : IComparable<T>
        {
            public T Minimum { get; set; }
            public T Maximum { get; set; }
            
            public override string ToString() { return String.Format("[{0} - {1}]", Minimum, Maximum); }

            public Boolean IsValid() { return Minimum.CompareTo(Maximum) <= 0; }

            public Boolean ContainsValue(T value)
            {
                return (Minimum.CompareTo(value) <= 0) && (value.CompareTo(Maximum) <= 0);
            }
        }
    }
}
