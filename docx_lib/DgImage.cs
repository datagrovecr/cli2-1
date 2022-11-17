
namespace DgConvert {


    // we need something to create files. this might be different for web streams
    public interface DgImageWriter {
        string write(string parent,byte[] child);
    }

    public class DgImageReader {
        byte[] read(string url);
    }
}