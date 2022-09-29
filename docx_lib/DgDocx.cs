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
using System.IO.Compression;

public class DgDocx
{

    public static void createWordprocessingDocument(string filepath)
    {
        // Create a document by supplying the filepath. 
        using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Create(filepath, WordprocessingDocumentType.Document))
        {
            // Add a main document part. 
            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));
        }

    }

    // stream here because anticipating zip.
    public async static Task md_to_docx(String md, Stream generatedDocument, bool debug = false) //String mdFile, String docxFile, String template)
    {
        var html = Markdown.ToHtml(md);

        // if (template != null)
        //     template.CopyTo(generatedDocument);
        generatedDocument.Position = 0L;
        using (WordprocessingDocument doc = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document, true))
        {
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            if (debug) {
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Markdown: " + md));
                }

                {
                    Paragraph para2 = body.AppendChild(new Paragraph());
                    Run run2 = para2.AppendChild(new Run());
                    run2.AppendChild(new Text("Html: " + html));
                }
            }



            //new Document(new Body()).Save(mainPart);
            HtmlConverter converter = new HtmlConverter(mainPart);
            converter.ParseHtml(html);
            mainPart.Document.Save();

        }
    }

    public async static Task docx_to_md(Stream infile, Stream outfile, String name)
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



        using (var archive = new ZipArchive(outfile, ZipArchiveMode.Create, true))
        {
            var demoFile = archive.CreateEntry(name);
            using (var entryStream = demoFile.Open())
            {
                using (var streamWriter = new StreamWriter(entryStream))
                {
                    String s = textBuilder.ToString();
                    streamWriter.Write(s);
                }
            }
        }
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




}