namespace Datagrove;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Presentation;

using ShapeCrawler;


// instead of marp, write these to a parquet file.
// use 

public partial class ConversionJob
{
    public async Task pptx()
    {
        using (Stream infile = await inputStream())
        {
            PresentationDocument doc = PresentationDocument.Open(infile, false);


            using var pres = SCPresentation.Open("some.pptx");

            // get number of slides
            var slidesCount = pres.Slides.Count;

            // get text of TextBox 
            var autoShape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
            var text = autoShape.TextFrame!.Text;
            
        }
    }
}