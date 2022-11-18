namespace Datagrove;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Presentation;


public partial class ConversionJob
{
    public async Task pptx()
    {
        using (Stream infile = await inputStream())
        {
            PresentationDocument doc = PresentationDocument.Open(infile, false);
            
        }
    }
}