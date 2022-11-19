namespace Datagrove;
using Parquet;
using Parquet.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using System.Text;
using System.Text.Json;

// https://github.com/aloneguid/parquet-dotnet
// https://github.com/ClosedXML/ClosedXML/issues/1778

// schema is sheet, column, row, type, value
// sheet 0 is meta information.
// 

class Workbook
{
    public List<int> sheet = new List<int>();
    public List<int> column = new List<int>();
    public List<int> row = new List<int>();
    public List<int> type = new List<int>();
    public List<byte[]> value = new List<byte[]>();

    public void add(int sheet, int row, int column, int type, byte[] value)
    {
        this.sheet.Add(sheet);
        this.row.Add(row);
        this.column.Add(column);
        this.type.Add(type);
        this.value.Add(value);
    }

    public async Task read(Stream s)
    {
    }
    public async Task Write(Stream ts)
    {
        var idColumn = new DataColumn[]{
            new DataColumn(new DataField<int>("sheet"), sheet.ToArray()),
            new DataColumn(new DataField<int>("column"), column.ToArray()),
            new DataColumn(new DataField<int>("row"), row.ToArray()),
            new DataColumn(new DataField<int>("type"), type.ToArray()),
            new DataColumn(new DataField<int>("value"), value.ToArray()),
        };
        Parquet.Data.DataField[] fld = idColumn.Select(e => e.Field).ToArray();
        var schema = new Parquet.Data.Schema();
        using (ParquetWriter parquetWriter = await ParquetWriter.CreateAsync(schema, ts))
        using (ParquetRowGroupWriter groupWriter = parquetWriter.CreateRowGroup())
        {
            foreach (var o in idColumn)
                await groupWriter.WriteColumnAsync(o);

        }
    }
}

public record Mark(
    String mark,
    String attribute
);

public  record Textmark (
    String fontFamily="",
    String fontName="",     
    String fontStyle="",
    String fontWeight="",
    String fontSize="",
    String color="", 
    String textDecoration="",
    String textShadow=""
);
public record Cellmark (
    String align=""
);

public class MarkedText
{
    String text;
    List<Mark>? mark;

    public MarkedText(String text) {
        this.text = text;
    }
    void add(String name, String value) {
        if (value!="") {
            this.mark!.Add(new Mark(name, value));
        }
    }
    public MarkedText(String text, Textmark m) {
        this.text = text;
        this.mark = new List<Mark>();
        add("fontFamily", m.fontFamily);
    }
}

// a cell value that's 
public partial class ConversionJob
{
   
    
    public async Task xlsx()
    {
        using (Stream infile = await inputStream())
        using (var workbook = new XLWorkbook(infile))
        using (Stream ts = fs.outputStream())
        {
            var ss = new Workbook();
            int sheetn = 0;
            foreach (var sh in workbook.Worksheets)
            {
                for (int x = 0; x < sh.ColumnCount(); x++)
                {
                    for (int y = 0; y < sh.RowCount(); y++)
                    {
                        IXLCell? v = sh.Cell(y, x);
                        if (v != null)
                        {
                            /*
                                public enum XLDataType
                                {
                                    Text = 0,
                                    Number = 1,
                                    Boolean = 2,
                                    DateTime = 3,
                                    TimeSpan = 4
                                }*/

                            int type = (int)v.DataType;
                            var str = "";
                            if (v.HasRichText)
                            {

                                type = 5;
                                var tx = new List<MarkedText>();
                                // convert rich text into prosemirror marks format
                                foreach (IXLRichString xx in v.GetRichText())
                                {
                                   String deco = xx.Strikethrough?"line-through":"";
                                   switch(xx.Underline){
                                    case XLFontUnderlineValues.Double:deco = "underline"; break;
                                    case XLFontUnderlineValues.DoubleAccounting:deco = "underline"; break;
                                    case  XLFontUnderlineValues.None: break;
                                    case  XLFontUnderlineValues.Single:deco = "underline"; break;
                                    case  XLFontUnderlineValues.SingleAccounting:deco = "underline"; break;
                                   }
                                   var c = System.Drawing.Color.FromArgb(xx.FontColor.Color.ToArgb());
                                   var a = System.Drawing.ColorTranslator.ToHtml(c);
                                   if (a=="#000000") a = "";

                                   var m = new Textmark(
                                        fontWeight: xx.Bold?"800":"",
                                        textDecoration: deco,
                                        fontStyle: xx.Italic?"italic":"",
                                        textShadow: xx.Shadow?"5px 10px":"",
                                        fontSize: xx.FontSize==0?"":$"${xx.FontSize}px",
                                        color: a,
                                        fontName: xx.FontName
                                        );
                                    tx.Add(new MarkedText(v.GetFormattedString(),m));
                                }
                                // now we need to serialize this to json.
                                str = JsonSerializer.Serialize(tx);
                            }
                            else
                            {
                                str = v.GetFormattedString();
                            }

                            var d = Encoding.UTF8.GetBytes(str);
                            ss.add(sheetn, y, x, type, d);
                        }
                    }
                }
                sheetn++;
            }
            await ss.Write(ts);

            await fs.commit(System.IO.Path.ChangeExtension(input, ".parquet"), ts);
        }

    }
}




