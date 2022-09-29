namespace docx_lib;
using Markdig;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.Text;

public class DgDocx
{

    public async static Task md_to_docx(Stream input, Stream generatedDocument, Stream? template) //String mdFile, String docxFile, String template)
    {
        if (template != null)
            template.CopyTo(generatedDocument);

        var md = new StreamReader(input).ReadToEnd();
        var html = Markdown.ToHtml(md);


        generatedDocument.Position = 0L;
        WordprocessingDocument package = WordprocessingDocument.Open(generatedDocument, true);

        MainDocumentPart mainPart = package.MainDocumentPart;
        if (mainPart == null)
        {
            mainPart = package.AddMainDocumentPart();
            new Document(new Body()).Save(mainPart);
        }

        HtmlConverter converter = new HtmlConverter(mainPart);
        Body body = mainPart.Document.Body;
        converter.ParseHtml(html);
        mainPart.Document.Save();

        AssertThatOpenXmlDocumentIsValid(package);
    }

    public async static Task docx_to_md(Stream infile, Stream outfile)
    {
        WordprocessingDocument wordDoc = WordprocessingDocument.Open(infile, false);
        DocumentFormat.OpenXml.Wordprocessing.Body body
        = wordDoc.MainDocumentPart.Document.Body;

        StringBuilder textBuilder = new StringBuilder();
        var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (parts != null)
        {
            foreach (var node in parts.ChildElements)
            {
                if (node is Paragraph)
                {
                    ProcessParagraph((Paragraph)node, textBuilder);
                    textBuilder.AppendLine("");
                }

                if (node is Table)
                {
                    ProcessTable((Table)node, textBuilder);
                }
            }
        }
        String s = textBuilder.ToString();
        new StreamWriter(outfile).Write(s);
    }

    private static void ProcessTable(Table node, StringBuilder textBuilder)
    {
        foreach (var row in node.Descendants<TableRow>())
        {
            textBuilder.Append("| ");
            foreach (var cell in row.Descendants<TableCell>())
            {
                foreach (var para in cell.Descendants<Paragraph>())
                {
                    ProcessParagraph(para, textBuilder);
                }
                textBuilder.Append(" | ");
            }
            textBuilder.AppendLine("");
        }
    }

    private static void ProcessParagraph(Paragraph node, StringBuilder textBuilder)
    {

        foreach (var run in node.Descendants<Run>())
        {
            String prefix = "";
            if (run.RunProperties != null)
            {
                if (run.RunProperties.Bold != null)
                {
                    prefix += "*";
                }
                if (run.RunProperties.Italic != null)
                {
                    prefix += "_";
                }
            }
            textBuilder.Append(prefix + run.InnerText + prefix + " ");
            prefix = "";
            //text.GetAttributes();

        }
        textBuilder.Append("\n\n");
    }



    static void AssertThatOpenXmlDocumentIsValid(WordprocessingDocument wpDoc)
    {

        var validator = new OpenXmlValidator(FileFormatVersions.Office2010);
        var errors = validator.Validate(wpDoc);

        if (!errors.GetEnumerator().MoveNext())
            return;

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("The document doesn't look 100% compatible with Office 2010.\n");

        Console.ForegroundColor = ConsoleColor.Gray;
        foreach (ValidationErrorInfo error in errors)
        {
            Console.Write("{0}\n\t{1}", error.Path.XPath, error.Description);
            Console.WriteLine();
        }

        Console.ReadLine();
    }
}