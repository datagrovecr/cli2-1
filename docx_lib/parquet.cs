namespace Datagrove;
using Parquet;
using Parquet.Data;

// https://github.com/aloneguid/parquet-dotnet
public partial class ConversionJob
{
    public async Task<Stream> inputStream(){
        Stream r = new MemoryStream();
        var b = await fs.Read(input);
        r.Write(b);
        return r;
    }
    public async Task parquet()
    {
        // open file stream
        using (Stream fileStream = await inputStream())
        {
            // open parquet file reader
            using (ParquetReader parquetReader = await ParquetReader.CreateAsync(fileStream))
            {
                // get file schema (available straight after opening parquet reader)
                // however, get only data fields as only they contain data values
                DataField[] dataFields = parquetReader.Schema.GetDataFields();

                // enumerate through row groups in this file
                for (int i = 0; i < parquetReader.RowGroupCount; i++)
                {
                    // create row group reader
                    using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                    {
                        // read all columns inside each row group (you have an option to read only
                        // required columns if you need to.
                        var columns = new DataColumn[dataFields.Length];
                        for (int c = 0; c < columns.Length; c++)
                        {
                            columns[c] = await groupReader.ReadColumnAsync(dataFields[c]);
                        }

                        // get first column, for instance
                        DataColumn firstColumn = columns[0];

                        // .Data member contains a typed array of column data you can cast to the type of the column
                        Array data = firstColumn.Data;
                        int[] ids = (int[])data;
                    }
                }
            }
        }
    }
}