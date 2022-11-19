namespace Datagrove;
using DocumentFormat.OpenXml;
using System;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;


// unlike the pdf module; we need to flatten the tree to get this into word
// basically  tables|paragraph > text runs.
public class Span
{
    public String text;
    public Style style;
    public Span(String text, Style style)
    {
        this.text = text;
        this.style = style;
    }
}


// a chunk is a table or paragraph.
public class Chunk {
    public List<Span> span = new List<Span>();
}

public class Style
{
    public  void merge(Style s){

    }
}

public class Section {
    public List<Chunk> chunk = new List<Chunk>(); 
}
public class ComputedStyle
{
    public List<Style> stack = new List<Style>();


    // this might not be necessary? just get computed style from chrome?
    // for pptx it might be hard to do another way?
    public Style style()
    {
        return new Style();
    }

    public void openStyle(Style s){
        s.merge(stack.Last());
        stack.Add(s);
    }
    public void closeStyle(){

    }
}

public enum NodeType {
    Section,
    Block,
    Span,
    Text
}

class WordStyle : ComputedStyle
{
   // word has sections with dramatic changes in formatting. for now just one though.
    List<Section> section = new List<Section>();

    void newSection() {
        section.Add(new Section());
    }
    void newChunk(){
        section.Last().chunk.Add(new Chunk());
    }

    NodeType nodeType(Vnode v) {
        switch(v.name){

        default: 
            return NodeType.Text;
        }
    }

    void visit1(Vnode v, Style s){
        openStyle(s);
        visit(v.children);
        closeStyle();
    }
    void visit(List<Vnode>? children){
        // push the style, 
        if (children==null) {
            return;
        }

        foreach (var v in children) {

            switch(nodeType(v)){
            case NodeType.Section:
                // clear 
            case NodeType.Block:
                // clear off open blocks and spans. This doesn't clear the style, styles track
                // the html nesting.


                visit1(v, new Style(

                ));
                break;
            case NodeType.Span:
                break;
            case NodeType.Text:
                break;
            }

            visit(v.children);
            closeStyle();
        }

    }
    static public void Write(Vnode v, Stream ts)
    {
        var ws = new WordStyle();
        var doc = WordprocessingDocument.Create(ts, WordprocessingDocumentType.Document, true);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

        // walk the vdom and flatten it into a chunk list.

        


        Run run = new Run();
        RunProperties runProperties = new RunProperties();

        runProperties.AppendChild<Underline>(new Underline() { Val = DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single });
        runProperties.AppendChild<Bold>(new Bold());
        run.AppendChild<RunProperties>(runProperties);
        //run.AppendChild(new Text("test"));

        //Note: I had to create a paragraph element to place the run into.
        Paragraph p = new Paragraph();
        p.AppendChild(run);
        doc.MainDocumentPart.Document.Body.AppendChild(p);




        mainPart.Document.Save();


    }
}

