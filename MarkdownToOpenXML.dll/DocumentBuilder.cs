
namespace MarkdownToOpenXML
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    using DocumentFormat.OpenXml;
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml.Packaging;

    using Ap = DocumentFormat.OpenXml.ExtendedProperties;
    using M = DocumentFormat.OpenXml.Math;
    using Ovml = DocumentFormat.OpenXml.Vml.Office;
    using V = DocumentFormat.OpenXml.Vml;
    using A = DocumentFormat.OpenXml.Drawing;
    using Sty = DocumentFormat.OpenXml.Wordprocessing.Style;
    using Clr = DocumentFormat.OpenXml.Wordprocessing.Color;

    public class DocumentBuilder
    {
        private Document document;

        public DocumentBuilder(Body body)
        {
            document = new Document();
            document.AppendChild(body);
        }
        
        public void SaveTo(string path)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                package.AddMainDocumentPart();
                package.MainDocumentPart.Document = document;

                StyleDefinitionsPart styleDefinitionsPart1 = package.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
                GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);
                
                package.MainDocumentPart.Document.Save();
            }
        }

        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart part)
        {
            Styles doc_styles = GenerateDocumentStyles();
            DocDefaults document_defaults = new DocDefaults();
            RunPropertiesDefault defaultRunProperties = new RunPropertiesDefault(CreateRunBaseStyle());
            document_defaults.Append(defaultRunProperties);

            ParagraphPropertiesBaseStyle paragraphBaseStyle = new ParagraphPropertiesBaseStyle();
            paragraphBaseStyle.Append(new SpacingBetweenLines()
            {
                After = "200",
                Line = "276",
                LineRule = LineSpacingRuleValues.Auto
            });

            ParagraphPropertiesDefault default_ParagraphProperties = new ParagraphPropertiesDefault();
            default_ParagraphProperties.Append(paragraphBaseStyle);
            document_defaults.Append(default_ParagraphProperties);

            LatentStyles latentStyles1 = new LatentStyles()
            {
                DefaultLockedState = false,
                DefaultUiPriority = 99,
                DefaultSemiHidden = true,
                DefaultUnhideWhenUsed = true,
                DefaultPrimaryStyle = false,
                Count = 267
            };

            latentStyles1.Append(new LatentStyleExceptionInfo()
            {
                Name = "Normal",
                UiPriority = 0,
                SemiHidden = false,
                UnhideWhenUsed = false,
                PrimaryStyle = true
            });
            Style Normal = GenerateNormal();
            
            doc_styles.Append(document_defaults);
            doc_styles.Append(latentStyles1);
            doc_styles.Append(Normal);

            for (int i = 1; i <= 7; i++)
            {
                latentStyles1.Append(new LatentStyleExceptionInfo()
                {
                    Name = "Heading "+i,
                    UiPriority = 9,
                    SemiHidden = false,
                    UnhideWhenUsed = false,
                    PrimaryStyle = true
                });
                Style Header = GenerateHeader(i);

                doc_styles.Append(Header);
            }

            part.Styles = doc_styles;
        }

        private RunPropertiesBaseStyle CreateRunBaseStyle()
        {
            RunPropertiesBaseStyle runBaseStyle = new RunPropertiesBaseStyle();

            RunFonts font = new RunFonts();
            font.Ascii = "Arial";
            runBaseStyle.Append(font);
            runBaseStyle.Append(new FontSize() { Val = "20" });
            runBaseStyle.Append(new FontSizeComplexScript() { Val = "20" });
            runBaseStyle.Append(new Languages()
            {
                Val = "en-GB",
                EastAsia = "en-GB",
                Bidi = "ar-SA"
            });

            return runBaseStyle;
        }

        private Styles GenerateDocumentStyles()
        {
            Styles doc_styles = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14" } };
            doc_styles.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            doc_styles.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            doc_styles.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            doc_styles.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            return doc_styles;
        }

        private Style GenerateNormal()
        {
            Style Style_Normal = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            };
            Style_Normal.Append(new StyleName() { Val = "Normal" });
            Style_Normal.Append(new PrimaryStyle());
            
            return Style_Normal;
        }

        private Style GenerateHeader(int id)
        {
            Style Style_Header1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading" + id };

            StyleParagraphProperties ParagraphProperties_Header = new StyleParagraphProperties();
            ParagraphProperties_Header.Append(new KeepNext());
            ParagraphProperties_Header.Append(new KeepLines());
            ParagraphProperties_Header.Append(new SpacingBetweenLines() { Before = "480", After = "0" });
            ParagraphProperties_Header.Append(new OutlineLevel() { Val = 0 });
            Style_Header1.Append(ParagraphProperties_Header);

            StyleRunProperties RunProperties_Header = new StyleRunProperties();
            RunFonts font = new RunFonts();
            font.Ascii = "Arial";
            RunProperties_Header.Append(font);
            RunProperties_Header.Append(new Bold());
            RunProperties_Header.Append(new BoldComplexScript());
            string size = (26 - (id * 2)).ToString();
            RunProperties_Header.Append(new FontSize() { Val = size });
            RunProperties_Header.Append(new FontSizeComplexScript() { Val = size });
            Style_Header1.Append(RunProperties_Header);

            Style_Header1.Append(new StyleName() { Val = "Heading " + id });
            Style_Header1.Append(new BasedOn() { Val = "Normal" });
            Style_Header1.Append(new NextParagraphStyle() { Val = "Normal" });
            Style_Header1.Append(new LinkedStyle() { Val = "Heading" + id + "Char" });
            Style_Header1.Append(new UIPriority() { Val = 9 });
            Style_Header1.Append(new PrimaryStyle());
            Style_Header1.Append(new Rsid() { Val = "00AF6F24" });

            return Style_Header1;
        }

    }
}