namespace Datagrove;
using Markdig;
using Markdig.Extensions.Yaml;
using Markdig.Syntax;

using System.Text;

using DocumentFormat.OpenXml;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using System.IO.Compression;
using System.Linq.Expressions;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Presentation;
using YamlDotNet;
using YamlDotNet.Helpers;
using YamlDotNet.Serialization;


using YamlDotNet.Serialization.NamingConventions;

public class Style
{

}
public class ComputedStyle
{
    Chunk? root;

    // this might not be necessary? just get computed style from anglesharp?

    public Style style()
    {
        return new Style();
    }
    public Chunk read1(Vnode e)
    {
        return new Chunk(e, style());
    }
    public void read(Vnode doc)
    {

    }
}

class WordStyle : ComputedStyle
{
    static public void Write(Vnode v, Stream ts)
    {
        var ws = new WordStyle();
        ws.read(v);
        var doc = WordprocessingDocument.Create(ts, WordprocessingDocumentType.Document, true);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();


        Run run = new Run();
        RunProperties runProperties = new RunProperties();

        runProperties.AppendChild<Underline>(new Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
        runProperties.AppendChild<Bold>(new Bold());
        run.AppendChild<RunProperties>(runProperties);
        run.AppendChild(new Text("test"));

        //Note: I had to create a paragraph element to place the run into.
        Paragraph p = new Paragraph();
        p.AppendChild(run);
        doc.MainDocumentPart.Document.Body.AppendChild(p);




        mainPart.Document.Save();


    }
}

public class Span
{
    String text;
    Style style;
    Span(String text, Style style)
    {
        this.text = text;
        this.style = style;
    }
}
public class Chunk
{
    public Vnode e;
    public Style style;
    public Chunk(Vnode e, Style style)
    {
        this.e = e;
        this.style = style;

        if (e.children != null)
            foreach (var c in e.children)
            {

            }
    }
    public List<Span> span = new List<Span>();
    public List<Chunk> block = new List<Chunk>();

}
