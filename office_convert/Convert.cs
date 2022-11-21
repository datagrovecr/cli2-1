
using System;
using System.IO.Compression;
using CommandLine;

using Datagrove;

namespace OfficeConvert
{

    // we need the reader to find files that potentially use .. and /
    // so we need to establish a 
    class ReadWriter : MyHtmlParser, Datagrove.DgReadWriter, HtmlParser
    {
        string readRoot;
        string writeRoot;

        public ReadWriter(string read, string write, bool unzip = false)
        {
            readRoot = read;
            writeRoot = write;
        }




        public async Task<byte[]> Read(string path)
        {
            return await Task.Run(() => File.ReadAllBytes(Path.Join(readRoot, path)));
        }
        public async Task Write(string path, byte[] data)
        {
            await Task.Run(() => File.WriteAllBytes(Path.Join(writeRoot, path), data));
        }
    }
    // converting a single docx or pptx file could result in a tree of markdown and image files.
    // 
    public class OfficeConvert
    {

        static public async Task Convert(string inpath, string outpath, string roundtrip = "", bool unzip = false)
        {
            // future work: take a zip file and produce a zip file.
            // extract zip file into temp directory
            string ext = Path.GetExtension(inpath);
            if (ext == ".zip")
            {
                string tempin = Path.GetTempPath() + "/dgin";
                string tempout = Path.GetTempPath() + "/dgout";
                string tempround = Path.GetTempPath() + "/dground";
                ZipFile.ExtractToDirectory(inpath, tempin);
                await ConvertDir(tempin, tempout, roundtrip == "" ? tempround : "", false);
                ZipFile.CreateFromDirectory(tempout, outpath);
            }
            else if (ext == "")
            {
                await ConvertDir(inpath, outpath, roundtrip, unzip);
            }
        }

        static public async Task ConvertDir(string input, string outdir, string roundtrip = "", bool unzip = false)
        {
            string[] files = Directory.GetFiles(input, "*.*", SearchOption.AllDirectories);
            var w = new ReadWriter(input, outdir);

            Parallel.ForEach(files, async (mdFile) =>
            {
                try
                {
                    await ConversionJob.Convert(mdFile.Substring(input.Length), w, w);
                }
                catch (Exception e)
                {
                    Console.WriteLine($"{mdFile} failed {e}");
                }
            });

            if (roundtrip != "")
            {
                await ConvertDir(outdir, roundtrip, "", false);
            }
            if (unzip)
            {
                string[] files2 = Directory.GetFiles(outdir, "*.*", SearchOption.AllDirectories);
                Parallel.ForEach(files2, async (path) =>
                {
                    string ext = Path.GetExtension(path);
                    switch (ext)
                    {
                        case ".docx":
                        case ".xlsx":
                        case ".pptx":
                            if (unzip)
                            {
                                using (ZipArchive archive = ZipFile.OpenRead(path))
                                {
                                    archive.ExtractToDirectory(path.Substring(0, path.Length - ext.Length), true);
                                }
                            }
                            break;
                    }
                });

            }
        }


    }

}

/*
            var tempStream = new MemoryStream();
            var writeStream = delegate ( string new_ext,  MemoryStream s,bool unzip)
            {
                string ext = Path.GetExtension(mdFile);    
                int len = mdFile.Length - input.Length - ext.Length;
                string suffix = mdFile.Substring(input.Length,len);
                string pathbase = Path.Join(outdir, suffix);
                string path = pathbase + new_ext;
                File.WriteAllBytes(path, tempStream.ToArray());
            };
            */


//commented code is for .zip files

//using (var archive = new ZipArchive(outfile, ZipArchiveMode.Create, true))
//{
//    var demoFile = archive.CreateEntry(name);
//    using (var entryStream = demoFile.Open())
//    {
//        using (var streamWriter = new StreamWriter(entryStream))
//        {
//            String s = textBuilder.ToString();
//            streamWriter.Write(s);
//        }
//    }
//}