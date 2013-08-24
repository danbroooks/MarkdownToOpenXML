
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
    using MarkdownToOpenXML.Classes;

    public class MD2OXML
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

        public static void CreateDocX(string md, string docName)
        {
            Body body = MarkdownToDocBody(md);
            SaveDocX(body, docName);
        }

        private static Body MarkdownToDocBody(string RawMarkdown)
        {
            Body body = new Body();
            SkipNextLine = false;
            Buffer = @"";
            int index = 0;

            List<string> md = StripDoubleLines(RawMarkdown);

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
        /*
      A space for playing with OpenXML elements
         *
            Paragraph para = new Paragraph();
            Run r = new Run();
            Text t = new Text("A Test");
            r.Append(t);
            para.Append(r);
            body.Append(para);
        */
            return body;
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

                TokenRange.add(new Range<int>()
                { 
                    Minimum = sToken, Maximum = match - 1
                });
                
                ranges.add(new Range<int>()
                {
                    Minimum = match, Maximum = eToken - 1
                });
                
                TokenRange.add(new Range<int>()
                {
                    Minimum = eToken, Maximum = endStr - 1
                });
            }

            return ranges;
        }
        
        private static Ranges<int> FindTabs(Regex rx, string s, ref Ranges<int> TokenRange)
        {
            Ranges<int> ranges = new Ranges<int>();
            MatchCollection mc = rx.Matches(s);

            foreach (Match m in mc)
            {
                ranges.add(new Range<int>()
                {
                    Minimum = m.Index,
                    Maximum = m.Index
                });

                TokenRange.add(new Range<int>()
                {
                    Minimum = m.Index,
                    Maximum = m.Index + m.Length-1
                });
            }
            return ranges;
        }

        private static ParagraphProperties GenerateParagraphProperties(string lookahead)
        {
            ParagraphProperties pPr = new ParagraphProperties();

            int headerLevel = Buffer.TakeWhile((x) => x == '#').Count();

            if (headerLevel > 0)
                Buffer = Buffer.TrimStart('#').TrimEnd('#').Trim();
            else
            {
                String sTest = Regex.Replace(lookahead, @"\w", "");
                Match isSetextHeader1 = Regex.Match(sTest, @"[=]{2,}");
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
                pPr.Append(new ParagraphStyleId()
                {
                    Val = "Heading" + headerLevel
                });

            Dictionary<JustificationValues, Match> Alignment = new Dictionary<JustificationValues, Match>();
            
            Alignment.Add(JustificationValues.Center, Regex.Match(Buffer, @"^><"));
            Alignment.Add(JustificationValues.Left, Regex.Match(Buffer, @"^<<"));
            Alignment.Add(JustificationValues.Right, Regex.Match(Buffer, @"^>>"));
            Alignment.Add(JustificationValues.Distribute, Regex.Match(Buffer, @"^<>"));

            foreach (KeyValuePair<JustificationValues, Match> match in Alignment)
            {
                if (match.Value.Success)
                {
                    pPr.Append(new Justification() { Val = match.Key });
                    Buffer = Buffer.Substring(2);
                    break;
                }
            }
            
            Match numberedList = Regex.Match(Buffer, @"^\\d\\.");

            // Set Paragraph Styles
            if (numberedList.Success)
            {
                // Doesnt work currently, needs NumberingDefinitions adding in filecreation.cs
                Buffer = Buffer.Substring(2);
                NumberingProperties nPr = new NumberingProperties(
                    new NumberingLevelReference() { Val = 0 },
                    new NumberingId() { Val = 1 }
                );

                pPr.Append(nPr);
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

            Ranges<int> Tabs = FindTabs(new Regex(@"(\\t|[\\ ]{4})"), Buffer, ref Tokens);

            if ((Bold.Count() + Italic.Count() + Underline.Count() + Tabs.Count()) == 0)
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
                    if (!Tokens.ContainsValue(CurrentPosition) || Tabs.ContainsValue(CurrentPosition))
                    {
                        run = new Run();

                        if (Tabs.ContainsValue(CurrentPosition))
                        {
                            run.Append(new TabChar());
                        }
                        else
                        {
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
                        }
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
    }
}
