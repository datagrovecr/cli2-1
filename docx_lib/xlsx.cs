namespace Datagrove;
using Parquet;
using Parquet.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Spreadsheet;

// https://github.com/aloneguid/parquet-dotnet

public partial class ConversionJob
{
    public async Task xlsx()
    {
        using (Stream infile = await inputStream())
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(infile, false);


            var idColumn = new DataColumn(
        new DataField<int>("id"),
        new int[] { 1, 2 });

            var cityColumn = new DataColumn(
                new DataField<string>("city"),
                new string[] { "London", "Derby" });

            // create file schema
            var schema = new Parquet.Data.Schema(idColumn.Field, cityColumn.Field);

            using (Stream fileStream = System.IO.File.OpenWrite("c:\\test.parquet"))
            {
                using (ParquetWriter parquetWriter = await ParquetWriter.CreateAsync(schema, fileStream))
                {
                    // create a new row group in the file
                    using (ParquetRowGroupWriter groupWriter = parquetWriter.CreateRowGroup())
                    {
                        await groupWriter.WriteColumnAsync(idColumn);
                        await groupWriter.WriteColumnAsync(cityColumn);
                    }
                }
            }
        }
    }


}