
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

    public class MD2OXMLFile
    {        
        public void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart part)
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

            latentStyles1.Append(new LatentStyleExceptionInfo()
            {
                Name = "Heading 1",
                UiPriority = 9,
                SemiHidden = false,
                UnhideWhenUsed = false,
                PrimaryStyle = true
            });
            Style Header1 = GenerateHeader1();

            doc_styles.Append(document_defaults);
            doc_styles.Append(latentStyles1);
            doc_styles.Append(Normal);
            doc_styles.Append(Header1);

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

        private Style GenerateHeader1()
        {
            Style Style_Header1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Heading1" };

            StyleParagraphProperties ParagraphProperties_Header1 = new StyleParagraphProperties();
            ParagraphProperties_Header1.Append(new KeepNext());
            ParagraphProperties_Header1.Append(new KeepLines());
            ParagraphProperties_Header1.Append(new SpacingBetweenLines() { Before = "480", After = "0" });
            ParagraphProperties_Header1.Append(new OutlineLevel() { Val = 0 });
            Style_Header1.Append(ParagraphProperties_Header1);

            StyleRunProperties RunProperties_Header1 = new StyleRunProperties();
            RunFonts font = new RunFonts();
            font.Ascii = "Arial";
            RunProperties_Header1.Append(font);
            RunProperties_Header1.Append(new Bold());
            RunProperties_Header1.Append(new BoldComplexScript());
            RunProperties_Header1.Append(new FontSize() { Val = "24" });
            RunProperties_Header1.Append(new FontSizeComplexScript() { Val = "24" });
            Style_Header1.Append(RunProperties_Header1);

            Style_Header1.Append(new StyleName() { Val = "Heading 1" });
            Style_Header1.Append(new BasedOn() { Val = "Normal" });
            Style_Header1.Append(new NextParagraphStyle() { Val = "Normal" });
            Style_Header1.Append(new LinkedStyle() { Val = "Heading1Char" });
            Style_Header1.Append(new UIPriority() { Val = 9 });
            Style_Header1.Append(new PrimaryStyle());
            Style_Header1.Append(new Rsid() { Val = "00AF6F24" });

            return Style_Header1;
        }

    }
}