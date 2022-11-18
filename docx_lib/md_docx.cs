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
using YamlDotNet.Serialization.NamingConventions;


public partial class ConversionJob
{
    public async Task md_docx(string html) {

        using (Stream ts = fs.outputStream()) {
            using (WordprocessingDocument docx = WordprocessingDocument.Create(ts, WordprocessingDocumentType.Document, true)) {
                    MainDocumentPart mainPart = docx.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(html);
                    mainPart.Document.Save();
                }

               await fs.commit(Path.ChangeExtension(input,".docx"),ts);          
            }
        

    }


}
