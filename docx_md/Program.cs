using docx_lib;

internal class Program
{


    private static void Main(string[] args)
    {
        var mdFile = "../test.md";
        var mdFile2 = "../test2.md";
        var docxFile = mdFile + ".docx";
        if (File.Exists(docxFile)) File.Delete(docxFile);

        // var buffer = new FileStream(template, FileMode.Open, FileAccess.Read))

        {
            var fileStream = File.Open(mdFile, FileMode.Open);
            var outStream = File.Open(mdFile + ".docx", FileMode.Create);
            if (fileStream == null || outStream == null)
            {
                return;
            }
            DgDocx.md_to_docx(fileStream, outStream, null); // template);
            outStream.Close();
        }

        var instream = File.Open(docxFile, FileMode.Open);
        var outstream = File.Open(mdFile2, FileMode.Create);
        DgDocx.docx_to_md(instream, outstream);
        // docx_lib.DgDocx.roundTrip(, "../template.docx");
    }
}