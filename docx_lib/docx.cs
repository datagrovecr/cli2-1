namespace Datagrove;
using Markdig;
using System.Text;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO.Compression;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.EMMA;

public partial class ConversionJob
{

    public async Task docx()
    {
        using (Stream infile = await inputStream()){
        WordprocessingDocument wordDoc = WordprocessingDocument.Open(infile, false);
        DocumentFormat.OpenXml.Wordprocessing.Body body
        = wordDoc.MainDocumentPart.Document.Body;

        StringBuilder textBuilder = new StringBuilder();
        var parts = wordDoc.MainDocumentPart.Document.Descendants().FirstOrDefault();
        StyleDefinitionsPart styleDefinitionsPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
        if (parts != null)
        {
            foreach (var block in parts.ChildElements)
            {
                if (block is Paragraph)
                {
                    //This method is for manipulating the style of Paragraphs and text inside
                    ProcessParagraph((Paragraph)block, textBuilder);
                }

                if (block is Table)
                {
                    ProcessTable((Table)block, textBuilder);
                }
            }
        }

        //This code is replacing the below one because I need to check the .md file faster
        //writing the .md file in test_result folder

            var b2= Encoding.UTF8.GetBytes(textBuilder.ToString());
            await fs.Write("", b2);
          }
    }

    private static void ProcessTable(Table node, StringBuilder textBuilder)
    {
        List<string> headerDivision = new List<string>();
        int rowNumber = 0;

        foreach (var row in node.Descendants<TableRow>())
        {
            rowNumber++;
            
            if(rowNumber == 2)
            {
                headerDivider(headerDivision, textBuilder);
            }

            textBuilder.Append("| ");
            foreach (var cell in row.Descendants<TableCell>())
            {
                foreach (var para in cell.Descendants<Paragraph>())
                {
                    if(para.ParagraphProperties != null)
                    {
                        headerDivision.Add(para.ParagraphProperties.Justification.Val);
                    }
                    else
                    {
                        headerDivision.Add("normal");
                    }
                    textBuilder.Append(para.InnerText);
                }
                textBuilder.Append(" | ");
            }
            textBuilder.AppendLine("");
        }
    }

    private static String ProcessParagraphElements(Paragraph block)
    {
        String style = block.ParagraphProperties.ParagraphStyleId.Val;
        int num;
        String prefix = "";

        //to find Heading Paragraphs
        if (style.Contains("Heading"))
        {
            num = int.Parse(style.Substring(style.Length - 1));
            

            for(int i = 0; i<num; i++)
            {
                prefix += "#";
            }
            return prefix;
        }

        //to find List Paragraphs
        if (style == "ListParagraph")
        {
            return prefix = "-";
        }

        //to find quotes Paragraphs
        if (style == "IntenseQuote")
        {
            return prefix = ">";
        }

        return null;
    }

    private static String ProcessRunElements(Run run)
    {
        //extract the child element of the text (Bold or Italic)
        OpenXmlElement expression = run.RunProperties.ChildElements.ElementAtOrDefault(0);

        String prefix = "";

        //to know if the propertie is Bold, Italic or both
        switch (expression)
        {
            case Bold:
                if (run.RunProperties.ChildElements.Count == 2)
                {
                    prefix = "***";
                    break;
                }
                prefix = "**";
                break;
            case Italic:
                prefix = "*";
                break;
        }
        return prefix;
    }

    private static String ProcessBlockQuote(Run block)
    {
        String text = block.InnerText;
        String[] textSliced = text.Split("\n");
        String textBack = "";

        foreach(String n in textSliced)
        {
            textBack += "> "+n+"\n";
        }

        return textBack;
    }

    private static void ProcessParagraph(Paragraph block, StringBuilder textBuilder)
    {
        String constructorBase = "";

        //iterate along every element in the Paragraphs and childrens
        foreach (var run in block.Descendants<Run>())
        {
            String prefix = "";

            if (run.InnerText != "")
            {
                constructorBase = run.InnerText;
            }

            // fonts, size letter, links
            if (run.RunProperties != null)
            {
                prefix = ProcessRunElements(run);
                constructorBase = prefix + constructorBase + prefix;
            }

            //general style, lists, aligment, spacing
            if (block.ParagraphProperties != null)
            {
                prefix = ProcessParagraphElements(block);

                if (prefix.Contains("#") || prefix.Contains("-"))
                {
                    constructorBase = prefix + " " + constructorBase;
                }
                if (prefix.Contains(">"))
                {
                    constructorBase = ProcessBlockQuote(run);
                }

            }
            textBuilder.Append(constructorBase);
            constructorBase = "";
        }
        
        textBuilder.Append("\n\n");
    }

    private static void headerDivider(List<String> align, StringBuilder textBuilder)
    {
        textBuilder.Append("|");
        foreach (var column in align)
        {
            switch (column)
            {
                case "left":
                    textBuilder.Append(":---|");
                    break;

                case "center":
                    textBuilder.Append(":---:|");
                    break;

                case "right":
                    textBuilder.Append("---:|");
                    break;

                case "normal":
                    textBuilder.Append("---|");
                    break;
            }
        }
        textBuilder.AppendLine("");

    }


}


