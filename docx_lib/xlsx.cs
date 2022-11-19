namespace Datagrove;
using Parquet;
using Parquet.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

// https://github.com/aloneguid/parquet-dotnet
// https://github.com/ClosedXML/ClosedXML/issues/1778

// schema is sheet, column, row, type, value
// sheet 0 is meta information.
// 
public partial class ConversionJob
{
    public async Task xlsx()
    {
        using (Stream infile = await inputStream())
        using (var workbook = new XLWorkbook(infile))
        using (Stream fileStream = System.IO.File.OpenWrite(fs.outputStream()))
        {


            int count = 0;
            foreach (var sh in workbook.Worksheets)
            {
                for (int x=0; x<sh.ColumnCount(); x++){
                    for (int y=0; y<sh.RowCount(); y++){
                        var cell = sh.Cell(y,x);
                    }
                }

                var idColumn = new DataColumn(new DataField<int>("id"), new int[] { 1, 2 });

                var cityColumn = new DataColumn(
                    new DataField<string>("city"),
                    new string[] { "London", "Derby" });

                // create file schema
                var schema = new Parquet.Data.Schema(idColumn.Field, cityColumn.Field);

                
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

                sh++;
            }
        }
    }
}




