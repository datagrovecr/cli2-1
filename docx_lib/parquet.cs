namespace Datagrove;

// note that this version does not support web assembly
using ClosedXML.Excel;
using Parquet;
using Parquet.Data;

// https://github.com/aloneguid/parquet-dotnet
public partial class ConversionJob
{
    public async Task<Stream> inputStream()
    {
        Stream r = new MemoryStream();
        var b = await fs.Read(input);
        r.Write(b);
        return r;
    }
    public async Task parquet()
    {
        using (Stream fileStream = await inputStream())
        using (var workbook = new XLWorkbook())
        using (ParquetReader parquetReader = await ParquetReader.CreateAsync(fileStream))
        {
            IXLWorksheet? worksheet = null;
            // these should be sheet, column, row, type, value.
            // sheet 0 is meta information.
            DataField[] dataFields = parquetReader.Schema.GetDataFields();
 
            // thre should only be one of these? maybe we could use them as a way of logging updates.

            int shn=0;
            for (int i = 0; i < parquetReader.RowGroupCount; i++)
            {
                using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                {
                    DataColumn sheetx = await groupReader.ReadColumnAsync(dataFields[0]);
                    DataColumn colx = await groupReader.ReadColumnAsync(dataFields[1]);
                    DataColumn rowx = await groupReader.ReadColumnAsync(dataFields[2]);
                    DataColumn typex = await groupReader.ReadColumnAsync(dataFields[3]);                    
                    DataColumn valuex = await groupReader.ReadColumnAsync(dataFields[3]); 

                    int[] sheet = (int[])sheetx.Data;
                    int[] col = (int[])colx.Data;
                    int[] row = (int[])rowx.Data;
                    int[] type = (int[])typex.Data;
                    byte[][] value = (byte[][])typex.Data;
                        
                    for (int k=0; k<value.Length; k++){
                        if (sheet[k] > shn) {
                            worksheet=workbook.AddWorksheet($"${shn}");
                        }
                        if (worksheet!=null)
                         worksheet.Cell(row[k],col[k]).Value = value[k];
                    }
                }
            }

            using (Stream ts = fs.outputStream())
                workbook.SaveAs(ts);
        }
    }

    public async Task parquetRaw()
    {
        using (Stream fileStream = await inputStream())
        using (var workbook = new XLWorkbook())
        using (ParquetReader parquetReader = await ParquetReader.CreateAsync(fileStream))
        {
            var worksheet = workbook.Worksheets.Add("1");
            DataField[] dataFields = parquetReader.Schema.GetDataFields();

            long row = 1;
            for (int i = 0; i < parquetReader.RowGroupCount; i++)
            {
                using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                {
                    for (int c = 0; c < dataFields.Length; c++)
                    {
                       var colname = ConversionJob.ExcelColumnFromNumber(c);
                        DataColumn col = await groupReader.ReadColumnAsync(dataFields[c]);
                        for (long r = 0; r<col.Count; r++){
                            worksheet.Cell($"${colname}${row+r}").Value = col.Data.GetValue(r);  
                        }
                           
                    }
                    row += groupReader.RowCount;
                }
            }

            using (Stream ts = fs.outputStream())
                workbook.SaveAs(ts);
        }
    }

        public static string ExcelColumnFromNumber(int column)
        {
            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }
}


