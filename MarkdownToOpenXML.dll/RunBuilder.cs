using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace MarkdownToOpenXML
{
    class RunBuilder
    {
        private string md;
        public Paragraph para;
        private Run run;
        Ranges<int> Tokens = new Ranges<int>();
        Ranges<int> Bold;
        Ranges<int> Italic;
        Ranges<int> Underline;
        Ranges<int> Tabs;

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

        public RunBuilder(string md, Paragraph para)
        {
            this.para = para;
            this.md = md;
            FindRanges();
            GenerateRuns();
        }

        private Ranges<int> FindMatches(Regex rx)
        {
            Ranges<int> ranges = new Ranges<int>();
            MatchCollection mc = rx.Matches(md);

            foreach (Match m in mc)
            {
                int sToken = m.Groups[1].Index;
                int match = m.Groups[2].Index;
                int eToken = m.Groups[3].Index;
                int endStr = m.Groups[3].Index + m.Groups[3].Length;

                Tokens.add(new Range<int>()
                {
                    Minimum = sToken,
                    Maximum = match - 1
                });

                ranges.add(new Range<int>()
                {
                    Minimum = match,
                    Maximum = eToken - 1
                });

                Tokens.add(new Range<int>()
                {
                    Minimum = eToken,
                    Maximum = endStr - 1
                });
            }

            return ranges;
        }

        private Ranges<int> FindTabs(Regex rx)
        {
            Ranges<int> ranges = new Ranges<int>();
            MatchCollection mc = rx.Matches(md);

            foreach (Match m in mc)
            {
                ranges.add(new Range<int>()
                {
                    Minimum = m.Index,
                    Maximum = m.Index
                });

                Tokens.add(new Range<int>()
                {
                    Minimum = m.Index,
                    Maximum = m.Index + m.Length - 1
                });
            }
            return ranges;
        }

        private void FindRanges()
        {
            Regex pBold, pItalic, pUnderline;
            AssignPatterns(out pBold, out pItalic, out pUnderline);
            Bold = FindMatches(pBold);
            Italic = FindMatches(pItalic);
            Underline = FindMatches(pUnderline);
            Tabs = FindTabs(new Regex(@"(\\t|[\\ ]{4})"));
        }

        private void AssignPatterns(out Regex pBold, out Regex pItalic, out Regex pUnderline)
        {
            if (!MD2OXML.CustomMode)
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
        }

        private bool NoMatches()
        {
            return ( Bold.Count() + Italic.Count() + Underline.Count() + Tabs.Count() ) == 0;  
        }

        private void GenerateRuns()
        {
            // Calculate positions of all tokens and use this to set 
            // run styles when iterating through the string

            // in the same calculation note down location of tokens 
            // so they can be ignored when loading strings into the buffer

            if (NoMatches())
            {
                run = new Run();
                run.Append(new Text(md)
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
                para.Append(run);
            }
            else
            {
                int CurrentPosition = 0;
                
                // This needs optimizing so it builds a string buffer before adding the run itself

                while (CurrentPosition < md.Length)
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
                            
                            if (Bold.ContainsValue(CurrentPosition))
                            {
                                rPr.Append(new Bold() { Val = new OnOffValue(true) });
                            }

                            if (Italic.ContainsValue(CurrentPosition))
                            {
                                rPr.Append(new Italic());
                            }

                            if (Underline.ContainsValue(CurrentPosition))
                            {
                                rPr.Append(new Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
                            }
                            
                            run.Append(rPr);

                            string TextBuffer = md.Substring(CurrentPosition, 1);
                            run.Append(new Text(TextBuffer)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            });
                        }
                        para.Append(run);
                    }
                    CurrentPosition++;
                };
            }
        }
    }
}
