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

using AngleSharp;
using AngleSharp.Html;
using YamlDotNet.Serialization.NamingConventions;
using AngleSharp.Html.Dom;
using AngleSharp.Dom;
using AngleSharp.Css;

public partial class ConversionJob
{

    AngleSharp.Dom.IDocument parseHtml(string html)
    {
        var config = Configuration.Default.WithCss();
        var context = BrowsingContext.New(config);

        //var parser = new AngleSharp.Html.Parser.HtmlParser();
        var document = parser.ParseDocument(html);
        return document;
    }
    Chunk h2c(IHtmlElement e, ComputedStyle st)
    {
        string tag = e.TagName;
        return new Chunk(e, st.style());
    }

    public async Task md_pptx(string html)
    {
        using (Stream ts = fs.outputStream())
        {


            await fs.commit(System.IO.Path.ChangeExtension(input, ".docx"), ts);
        }
    }

    public async Task md_docx(string html)
    {
        using (var dm = parseHtml(html))
            if (dm.Body != null)
            {
                using (Stream ts = fs.outputStream())
                {
                    var ws = new WordStyle(ts);
                    ws.read(dm);
                    ws.write();

                    // HtmlConverter converter = new HtmlConverter(mainPart);
                    // converter.ParseHtml(html);

                    await fs.commit(Path.ChangeExtension(input, ".docx"), ts);
                }

            }
    }
}



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
    public Chunk read(IHtmlElement e){
        return new Chunk(e,style());
    }
    public void read(IDocument doc)
    {
        var body = doc.Body;
    
        if (body!=null) {
            root = read(body);
        }
    }
}
class PowerpointStyle : ComputedStyle
{
    PresentationDocument doc;


    public PowerpointStyle(Stream ts)
    {
        using (var docx = PresentationDocument.Create(ts, PresentationDocumentType.Presentation))
        {
            var mainPart = docx.AddPresentationPart();
            mainPart.Presentation = new Presentation();
            mainPart.Presentation.Save();
        }
    }

    public void write()
    {


    }
}
class WordStyle : ComputedStyle
{
    WordprocessingDocument doc;
    MainDocumentPart mainPart;
    Stream ts;

    public WordStyle(Stream ts)
    {
        this.ts = ts;
        doc = WordprocessingDocument.Create(ts, WordprocessingDocumentType.Document, true);
        mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
    }
    public void write()
    {
        mainPart.Document.Save();
    }
}

public class Span
{

}
public class Chunk
{
    public IHtmlElement e;
    public Style style;
    public Chunk(IHtmlElement e, Style style)
    {
        this.e = e;
        this.style = style;

        e.com
        foreach (var c in e.Children){

        }
    }
    public List<Span> span = new List<Span>();
    public List<Chunk> block = new List<Chunk>();

}
