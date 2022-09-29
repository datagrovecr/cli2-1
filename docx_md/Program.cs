using docx_lib;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

internal class Program
{
    static async Task Main(string[] args)
    {
        var outdir = "../test_results/";
        var template = "../template.docx";

        string[] files = Directory.GetFiles("../test_md", "*.md", SearchOption.AllDirectories);
        DgDocx.createWordprocessingDocument("../test_results/hello.docx");
        foreach (var mdFile in files)
        {
            string fn = Path.GetFileNameWithoutExtension(mdFile);
            string root = outdir + fn;
            var docxFile = root + ".docx";
            {
                try
                {
                   // markdown to docx
                    using (var mdfile = File.Open(mdFile, FileMode.Open)){
                        var doc = new MemoryStream();
                        var buffer = new FileStream(template, FileMode.Open, FileAccess.Read);
                        await DgDocx.md_to_docx(mdfile, doc, buffer); // template);
                        File.WriteAllBytes(docxFile, doc.ToArray());                       
                    }
                                
                    // convert the docx back to markdown.
                    using (var instream = File.Open(docxFile, FileMode.Open)){
                        var outstream = new MemoryStream();
                        await DgDocx.docx_to_md(instream, outstream,fn+".md");
                        using (var fileStream = new FileStream(root+".md.zip", FileMode.Create))
                        {
                            outstream.Seek(0, SeekOrigin.Begin);
                            outstream.CopyTo(fileStream);
                        }                        
                    }
                }catch (Exception e)
                {
                    Console.WriteLine($"{mdFile} failed {e}");
                }
            }
        }
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