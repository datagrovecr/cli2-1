namespace Datagrove
{

    // we need something to create files. this might be different for web streams
    public interface DgWriter
    {
        public Task Write(string url, byte[] child);
    }

    public interface DgReader
    {
        public Task<byte[]> Read(string url);
    }
    public interface DgReadWriter : DgWriter, DgReader
    {
        // we might want to make this a disk stream
        Stream outputStream()
        {
            return new MemoryStream();
        }
        public async Task commit(string path, Stream s)
        {
            await Write(path, (s as MemoryStream).ToArray());
        }
    }



    // for web we can make a simple memory based reader and writer.

    public class MemoryFs : DgWriter, DgReader
    {
        public Dictionary<String, byte[]> data = new Dictionary<String, byte[]>();

        public async Task<byte[]> Read(string url)
        {
            return new byte[] { };
        }
        public async Task Write(string url, byte[] child)
        {

        }

        public static MemoryFs Convert(string input)
        {
            return new MemoryFs();
        }
    }


    // we'll need a context to pass arguments from the command line.
    // input bytes is not enough when converting markdown, we need a way to get images etc.
    public interface HtmlParser
    {
        Vnode parse(string html);
    }

    public class Vnode
    {
        public string name;
        public List<Vnode>? children;
        public Dictionary<string, string>? attribute;
        public Dictionary<string, string>? style;
        public string? text;
        public Vnode(string name)
        {
            this.name = name;
        }
    }


    public partial class ConversionJob
    {
        public String input; // name to read from fs. 

        public DgReadWriter fs;
        public HtmlParser ps;

        public ConversionJob(String input, DgReadWriter fs, HtmlParser p)
        {
            this.ps = p;
            this.input = input;
            this.fs = fs;
        }

        public static async Task Convert(String input, DgReadWriter fs, HtmlParser p)
        {
            await new ConversionJob(input, fs, p).exec();
        }

        public async Task exec()
        {
            string ext = Path.GetExtension(input);
            switch (ext)
            {
                case ".md":
                    // might produce pptx or docx depending on the front matter.
                    await md();
                    break;
                case ".parquet":
                    // potentially flag for csv vs excel? seems like most things that read csv will read parquet before long.
                    await parquet();
                    break;
                case ".docx":
                    await docx();
                    break;
                case ".xlsx":
                    await xlsx();
                    break;
                case ".pptx":
                    await pptx();
                    break;
            }
        }
    }

}