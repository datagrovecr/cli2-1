namespace Datagrove;
using Markdig;
using Markdig.Extensions.Yaml;
using Markdig.Syntax;
using System.Text;
using System.Linq;
using YamlDotNet.Serialization;


public partial class ConversionJob


    // this will 
    public async Task md()
    {
        MarkdownPipeline pipeline = new MarkdownPipelineBuilder()
            .UseYamlFrontMatter()
            .UseAdvancedExtensions()
            .UseCustomContainers()
            .UseEmphasisExtras()
            .UseGridTables()
            .UseMediaLinks()
            .UsePipeTables()
            .UseGenericAttributes()
            .Build();
        var mdb = await fs.Read(input);
        var md = Encoding.UTF8.GetString(mdb    );
        var x = md.GetFrontMatter<FrontMatter>();
        bool marp= x?.Marp??false;   // marp: true

        // how do I get the front matter here?
        var document = Markdown.Parse(md, pipeline);
        var html = Markdown.ToHtml(md, pipeline);
        byte[] bytes = Encoding.UTF8.GetBytes(html);
        if (marp){
            await md_pptx(html);
        } else {
            await md_docx(html);
        }

        //All the document is being saved in the stream

    }

}

// https://khalidabuhakmeh.com/parse-markdown-front-matter-with-csharp
public class FrontMatter
{
    [YamlMember(Alias = "marp")]
    public bool Marp { get; set; }
}

public static class MarkdownExtensions
{
    private static readonly IDeserializer YamlDeserializer = 
        new DeserializerBuilder()
        .IgnoreUnmatchedProperties()
        .Build();
    
    private static readonly MarkdownPipeline Pipeline 
        = new MarkdownPipelineBuilder()
        .UseYamlFrontMatter()
        .Build();

    public static T? GetFrontMatter<T>(this string markdown)
    {
        var document = Markdown.Parse(markdown, Pipeline);
        var yaml =
            document?.Descendants<YamlFrontMatterBlock>()
            ?.FirstOrDefault()?
            // this is not a mistake
            // we have to call .Lines 2x
            .Lines // StringLineGroup[]
            .Lines // StringLine[]
            .OrderByDescending(x => x.Line)
            .Select(x => $"{x}\n")
            .ToList()
            .Select(x => x.Replace("---", string.Empty))
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Aggregate((s, agg) => agg + s);
        if (yaml==null) return default;
        return YamlDeserializer.Deserialize<T>(yaml);
    }
}


/*
        MarkdownObject yamlBlock = document.Descendants().FirstOrDefault();

        var yaml = yamlBlock.Lines.ToString();
        var deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();

        var metadata = deserializer.Deserialize(yaml);
        */