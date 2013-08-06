
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
    using System.Diagnostics;

    using MarkdownToOpenXML;

    class MD2OXML
    {
        static private bool CustomMode;

        static public bool customMode
        {
            set { CustomMode = value; }
        }

        static public void CreateDocX(string content, string docName)
        {
            List<string> md = StripDoubleLines(content);

            Body body = new Body();
            int index = 0;
            bool SkipNextLine = false;
            foreach (string line in md)
            {
                if (SkipNextLine)
                {
                    index += 1;
                    SkipNextLine = !SkipNextLine;
                    continue;
                }

                string lookahead = index + 1 != md.Count ? md[index + 1] : "";
                string output = line;
                
                Match isHeader1 = Regex.Match(output, @"^#");
                String sTest = Regex.Replace(lookahead, @"\w", "");
                Match isSetextHeader1 = Regex.Match(sTest, @"[=]{2,}");
                ParagraphProperties pPr = new ParagraphProperties();

                // Set Paragraph Styles
                if (isHeader1.Success || isSetextHeader1.Success)
                {
                    ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Heading1" };
                    pPr.Append(paragraphStyleId1);

                    if (isHeader1.Success) output = output.Substring(1);

                    if (isSetextHeader1.Success) SkipNextLine = true;
                }
                
                Paragraph p = new Paragraph();
                p.Append(pPr);

                // Set Run Styles :-
                // Here it needs to weigh up what combo of formatting
                // each run is going to have before 'Committing' the
                // run to the paragraph
                bool bOn = false;
                bool iOn = false;
                bool uOn = false;

                int CurrentPosition = 0;
                int cropSymbol = 0;
                
                Regex pattern = CustomMode 
                    ? new Regex(@"(\*\*|`|_)")
                    : new Regex(@"(\*\*|\*)");

                Match m = pattern.Match(output);
                Run run = new Run();

                if (m.Success)
                {
                    RunProperties rPr = new RunProperties();
                    run.Append(new Text(output.Substring(CurrentPosition, m.Index)));

                    while (m.Success)
                    {
                        if (!CustomMode)
                        {
                            if (m.ToString() == "**") bOn = !bOn;
                            if (m.ToString() == "*") iOn = !iOn;
                        }
                        else
                        {
                            if (m.ToString() == "**") bOn = !bOn;
                            if (m.ToString() == "`") iOn = !iOn;
                            if (m.ToString() == "_") uOn = !uOn;
                        }

                        rPr = new RunProperties();

                        rPr.Append(new Bold() { Val = new OnOffValue(bOn) });
                        if (iOn) rPr.Append(new Italic());
                        if (uOn) rPr.Append(new Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
                        run.Append(rPr);

                        CurrentPosition = m.Index;
                        cropSymbol = CurrentPosition + m.Length;
                        m = m.NextMatch();

                        if (m.Index == 0)
                        {
                            run.Append(new Text(output.Substring(cropSymbol)));
                        }
                        else
                        {
                            run.Append(new Text(output.Substring(cropSymbol, m.Index - cropSymbol)));
                        }

                        p.Append(run);
                        run = new Run();
                    }
                }
                else
                {
                    run.Append(new Text(output));
                    p.Append(run);
                }

                body.Append(p);

                index += 1;
            }

            SaveDocX(body, docName);
        }
        
        static private List<string> StripDoubleLines(string text)
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
