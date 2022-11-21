using AngleSharp.Html.Dom;
using AngleSharp.Dom;
using AngleSharp;
using AngleSharp.Html;

using Datagrove;
// thoughts:
// in the browser we can use our own markdown & prosemirror translations
// on the command line, we can use puppeteer.

class MyHtmlParser
{

    Vnode h(IElement e)
    {
        var v = new Vnode(e.TagName);
        foreach (IAttr o in e.Attributes)
        {
            if (v.attribute == null)
            {
                v.attribute = new Dictionary<string, string>();
            }
            v.attribute.Add(o.Name, o.Value);
            if (o.Name == "style")
            {
                var h = (IHtmlElement)(e);
                h.ComputeCurrentStyle();
            }
        }
        if (e.HasChildNodes)
        {
            v.children = new List<Vnode>();
            foreach (var o in e.Children)
            {
                v.children.Add(h(o));
            }
        }

        return v;
    }
    public Vnode parse(string html)
    {
        var parser = new AngleSharp.Html.Parser.HtmlParser();
        var document = parser.ParseDocument(html);
        if (document == null || document.Body == null)
        {
            return new Vnode("body");
        }
        return h(document.Body);
    }
}
