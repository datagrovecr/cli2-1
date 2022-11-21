namespace Datagrove;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Presentation;

using ShapeCrawler;
using ShapeCrawler.SlideMasters;


// instead of marp, write these to a parquet file.
// <0 = master > 0 = slide
// create table (slide, shapenum, shapetype, serialized)
public partial class ConversionJob
{


    public async Task pptx()
    {
        using (Stream infile = await inputStream())
        {
            using var pres = SCPresentation.Open(infile);

            foreach (ISlideMaster v in pres.SlideMasters)
            {
                foreach (var sh in v.Shapes)
                {

                }
                var m = v.Background.MIME;
                var d = v.Background.BinaryData;
                foreach (ISlideLayout ly in v.SlideLayouts)
                {
                    foreach (var sh in ly.Shapes)
                    {

                    }
                }
            }
            foreach (ISlide v in pres.Slides)
            {
                var x = v.Background.BinaryData;
                var mt = v.Background.MIME;
                //var type = v.GetType(); just normal
                var baton = v.CustomData;
                var hidden = v.Hidden;
                var layout = v.SlideLayout;
                ISlideMaster master = layout.SlideMaster;
                foreach (var sh in layout.Shapes)
                {

                }
                SlidePart part = v.SDKSlidePart;

                foreach (ITextFrame tf in v.GetAllTextFrames())
                {
                    foreach (IParagraph p in tf.Paragraphs)
                    {
                        foreach (IPortion po in p.Portions)
                        {
                            string Text;
                            IFont fnt = po.Font;  // all the things
                            string? href = po.Hyperlink;
                            IField? fld = po.Field; // empty? no members.
                        }

                        SCBullet b = p.Bullet;
                        string emoji = b.Character;
                        string color = b.ColorHex;
                        string f = b.FontName;
                        int ps = b.Size;
                        int bt = (int)b.Type; /*
                             Numbered = 0,
                            Picture = 1,
                            Character = 2,
                            None = 3*/

                        /*
                                Left = 0,
                        Center = 1,
                        Right = 2,
                        Justify = 3*/
                        SCTextAlignment Alignment;
                        int indent = p.IndentLevel;
                    }
                    /*         None = 0,
            Shrink = 1,
            Resize = 2
            */
                    int autofit = (int)tf.AutofitType;
                    var margins = new double[] { tf.LeftMargin, tf.TopMargin, tf.RightMargin, tf.BottomMargin };
                    bool wrap = tf.TextWrapped;
                    bool isReadonly = !tf.CanChangeText();
                }

                //v.Shapes.AddNewAudio, video
                foreach (Shape sh in v.Shapes)
                {
                    if (sh is ITable)
                    {

                    }
                    // int X { get; set; }
                    // int Y { get; set; }
                    // int Width { get; set; }
                    // int Height { get; set; }
                    // int Id { get; }
                    // string Name { get; }
                    // bool Hidden { get; }
                    // IPlaceholder? Placeholder { get; }
                    // SCGeometry GeometryType { get; }
                    // string? CustomData { get; set; }
                    // SCShapeType ShapeType { get; }
                    // ISlideObject SlideObject { get; }
                }

            }




            // get number of slides
            var slidesCount = pres.Slides.Count;

            // get text of TextBox 
            var autoShape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
            var text = autoShape.TextFrame!.Text;

        }
    }
}

