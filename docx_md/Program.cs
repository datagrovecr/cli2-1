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

        foreach (var mdFile in files)
        {
            string root = outdir + Path.GetFileNameWithoutExtension(mdFile);
            var docxFile = root + ".docx";
            {
                var fileStream = File.Open(mdFile, FileMode.Open);
                Stream outStream = null;
                try {
                 outStream = File.Open(docxFile, FileMode.Create);

                }catch(Exception e){
                    
                }
                if (outStream!=null){
                    var buffer = new FileStream(template, FileMode.Open, FileAccess.Read);
                    await DgDocx.md_to_docx(fileStream, outStream, buffer); // template);
                    outStream.Close();                    
                }
            }

            {
                // convert the docx back to markdown.
                var instream = File.Open(docxFile, FileMode.Open);
                var outstream = File.Open(root + ".md", FileMode.Create);
                await DgDocx.docx_to_md(instream, outstream);
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