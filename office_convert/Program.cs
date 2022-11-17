using System;
using CommandLine;
using System.IO.Compression;
using docx_lib;

namespace OfficeConvert {
    
     class Options
    {
        [Option('o', "output", Required = true, HelpText = "Output directory")]
        [Option('i', "input", Required = true, HelpText = "Input directory")]
        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; }
    }
    class Program
    {



        // folder in - defaults to current
        // folder out - defaults to ../build
        static async Task Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {

                });
        }


}




 class ImageReader : DgConvert.DgImageReader {
    byte[] read(string link) {

    }
}
 class ImageWriter : DgConvert.DgImageWriter {
    string write(byte[] data, string type){

    }
}
// converting a single docx or pptx file could result in a tree of markdown and image files.
// 
 static class Lib
{
    static public string outpath(string f, string dir, string ext)
    {
        string fn = Path.GetFileNameWithoutExtension(mdFile);
        return dir + fn + ext;
    }
    static public convertZip(string inpath, string outpath, string roundtrip="")
    {
        // future work: take a zip file and produce a zip file.
        // extract zip file into temp directory
        string tempin = Path.GetTempPath() + "/dgin";
        string tempout = Path.GetTempPath() + "/dgout";
        string tempout = Path.GetTempPath() + "/dground";
        ZipFile.ExtractToDirectory(inpath, tempdir);
        convert(tempin, tempout, roundtrip==""?tempout:"");
        ZipFile.CreateFromDirectory(tempdir,outpath);
    }

    static public convert(string input, string outdir, string roundtrip, bool unzip)
    {
        string[] files = Directory.GetFiles(input, "*.md,*.parquet,*.docx,*.xlsx,*.pptx", SearchOption.AllDirectories);


        foreach (var mdFile in files)
        {
            var tempStream = new MemoryStream();
            var writeStream = delegate (string ext, s MemoryStream)
            {
                File.WriteAllBytes(outpath(tempStream, outdir, ".docx"), inputStream.ToArray());
                if (unzip)
                {
                    using (ZipArchive archive = ZipFile.OpenRead(outdir + "test.docx"))
                    {
                        archive.ExtractToDirectory(outdir + "test.unzipped", true);
                    }
                }
            };
            try
            {
                string ext = Path.GetExtension(mdFile);
                switch (ext)
                {
                    case ".md":
                        //await DgDocx.md_to_docx(tempStream, File.ReadAllText(inpath));
                        writeStream(".docx");
                        break;
                    case ".parquet":
                        break;
                    case ".docx":
                        using (var instream = File.Open(docxFile, FileMode.Open))
                        {
                            var outstream = new MemoryStream();
                            // await DgDocx.docx_to_md(instream, outstream, root);
                        }
                    case ".xlsx":
                        break;
                    case ".pptx":
                        break;
                }
            }
            catch (e)
            {
                Console.WriteLine($"{mdFile} failed {e}");
            }
        }
        if (roundtrip != "")
        {
            convert(outdir, roundtrip, "");
        }
    }


}

}