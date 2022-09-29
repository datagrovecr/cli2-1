using docx_lib;

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
                var outStream = File.Open(docxFile, FileMode.Create);
                var buffer = new FileStream(template, FileMode.Open, FileAccess.Read);
                await DgDocx.md_to_docx(fileStream, outStream, buffer); // template);
                outStream.Close();
            }

            {
                // convert the docx back to markdown.
                var instream = File.Open(docxFile, FileMode.Open);
                var outstream = File.Open(root + ".md", FileMode.Create);
                await DgDocx.docx_to_md(instream, outstream);
            }
        }
    }
}