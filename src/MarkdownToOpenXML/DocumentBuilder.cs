using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MarkdownToOpenXML;

public class DocumentBuilder
{
    private readonly Document _document;

    public DocumentBuilder(Body body)
    {
        _document = new Document();
        _document.AppendChild(body);
    }

    public void SaveTo(string path)
    {
        using WordprocessingDocument package = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        package.AddMainDocumentPart();
        if (package.MainDocumentPart == null)
        {
            return;
        }

        package.MainDocumentPart.Document = _document;

        StyleDefinitionsPart styleDefinitionsPart1 = package.MainDocumentPart.AddNewPart<StyleDefinitionsPart>("rId1");
        GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

        package.MainDocumentPart.Document.Save();
    }

    private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart part)
    {
        Styles docStyles = GenerateDocumentStyles();
        DocDefaults documentDefaults = new();
        RunPropertiesDefault defaultRunProperties = new(CreateRunBaseStyle());
        documentDefaults.Append(defaultRunProperties);

        ParagraphPropertiesBaseStyle paragraphBaseStyle = new();
        paragraphBaseStyle.Append(new SpacingBetweenLines { After = "200", Line = "276", LineRule = LineSpacingRuleValues.Auto });

        ParagraphPropertiesDefault defaultParagraphProperties = new();
        defaultParagraphProperties.Append(paragraphBaseStyle);
        documentDefaults.Append(defaultParagraphProperties);

        LatentStyles latentStyles1 = new()
        {
            DefaultLockedState = false,
            DefaultUiPriority = 99,
            DefaultSemiHidden = true,
            DefaultUnhideWhenUsed = true,
            DefaultPrimaryStyle = false,
            Count = 267
        };

        latentStyles1.Append(
            new LatentStyleExceptionInfo
            {
                Name = "Normal",
                UiPriority = 0,
                SemiHidden = false,
                UnhideWhenUsed = false,
                PrimaryStyle = true
            });
        Style normal = GenerateNormal();

        docStyles.Append(documentDefaults);
        docStyles.Append(latentStyles1);
        docStyles.Append(normal);

        for (int i = 1; i <= 7; i++)
        {
            latentStyles1.Append(
                new LatentStyleExceptionInfo
                {
                    Name = "Heading " + i,
                    UiPriority = 9,
                    SemiHidden = false,
                    UnhideWhenUsed = false,
                    PrimaryStyle = true
                });
            Style header = GenerateHeader(i);

            docStyles.Append(header);
        }

        part.Styles = docStyles;
    }

    private static RunPropertiesBaseStyle CreateRunBaseStyle()
    {
        RunPropertiesBaseStyle runBaseStyle = new();

        RunFonts font = new() { Ascii = "Arial" };
        runBaseStyle.Append(font);
        runBaseStyle.Append(new FontSize { Val = "20" });
        runBaseStyle.Append(new FontSizeComplexScript { Val = "20" });
        runBaseStyle.Append(new Languages { Val = "en-GB", EastAsia = "en-GB", Bidi = "ar-SA" });

        return runBaseStyle;
    }

    private static Styles GenerateDocumentStyles()
    {
        Styles docStyles = new() { MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" } };
        docStyles.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
        docStyles.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        docStyles.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        docStyles.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        return docStyles;
    }

    private static Style GenerateNormal()
    {
        Style styleNormal = new() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
        styleNormal.Append(new StyleName { Val = "Normal" });
        styleNormal.Append(new PrimaryStyle());

        return styleNormal;
    }

    private static Style GenerateHeader(int id)
    {
        Style styleHeader1 = new() { Type = StyleValues.Paragraph, StyleId = "Heading" + id };

        StyleParagraphProperties paragraphPropertiesHeader = new();
        paragraphPropertiesHeader.Append(new KeepNext());
        paragraphPropertiesHeader.Append(new KeepLines());
        paragraphPropertiesHeader.Append(new SpacingBetweenLines { Before = "480", After = "0" });
        paragraphPropertiesHeader.Append(new OutlineLevel { Val = 0 });
        styleHeader1.Append(paragraphPropertiesHeader);

        StyleRunProperties runPropertiesHeader = new();
        RunFonts font = new() { Ascii = "Arial" };
        runPropertiesHeader.Append(font);
        runPropertiesHeader.Append(new Bold());
        runPropertiesHeader.Append(new BoldComplexScript());
        string size = (26 - (id * 2)).ToString();
        runPropertiesHeader.Append(new FontSize { Val = size });
        runPropertiesHeader.Append(new FontSizeComplexScript { Val = size });
        styleHeader1.Append(runPropertiesHeader);

        styleHeader1.Append(new StyleName { Val = "Heading " + id });
        styleHeader1.Append(new BasedOn { Val = "Normal" });
        styleHeader1.Append(new NextParagraphStyle { Val = "Normal" });
        styleHeader1.Append(new LinkedStyle { Val = "Heading" + id + "Char" });
        styleHeader1.Append(new UIPriority { Val = 9 });
        styleHeader1.Append(new PrimaryStyle());
        styleHeader1.Append(new Rsid { Val = "00AF6F24" });

        return styleHeader1;
    }
}