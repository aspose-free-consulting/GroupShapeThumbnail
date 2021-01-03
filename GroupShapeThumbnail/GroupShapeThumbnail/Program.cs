using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupShapeThumbnail
{
    class Program
    {
        public static void LoadLicense()
        {

            License lic = new License();
            lic.SetLicense("Aspose.Total.lic");
        }

        public static IShape GroupShapes(List<IShape>shapes,ISlide slide)
        {
            //Shape to be added to group shape
            //Adding a group shape
            IGroupShape groupShape = slide.Shapes.AddGroupShape();
            foreach (IShape shape in shapes)
            {
                //Adding existing shape inside group shape
                groupShape.Shapes.AddClone(shape);

                //Removing the shape from slide
                slide.Shapes.Remove(shape);
            }
           
            return groupShape;
        }



        public static void GenerateShapeThumbnail(IShape shape, String ThumbnailName)
        {
            // Create a full scale image
       //     using (Bitmap bitmap = shape.GetThumbnail(ShapeThumbnailBounds.Appearance,1,1))
            using (Bitmap bitmap = shape.GetThumbnail())
            {
                // Save the image to disk in PNG format
                bitmap.Save(ThumbnailName+".png", ImageFormat.Png);
            }
        }
        static void Main(string[] args)
        {
            //Load Api license to use full features and avoid watermark
            //LoadLicense();

            //Loading the presentation
            using (Presentation presentation =new Presentation(@"MultipleChart.pptx"))
            {
                //Accessing the first slide
                ISlide slide = presentation.Slides[0];

                List<IShape> Chart1 = new List<IShape>();
                List<IShape> Chart2 = new List<IShape>();

                for (int i = 0; i < slide.Shapes.Count; i++)
                {
                    if (i > 0 && i < 11)
                    {
                        Chart1.Add(slide.Shapes[i]);
                    }
                    else if (i >= 11 && i < 15)
                    {
                        Chart2.Add(slide.Shapes[i]);
                    }
                    else
                        continue;


                }

                //Grouping Shapes of Chart element
                IShape GroupedChart1 = GroupShapes(Chart1,slide);
                String fileName = "Chart1" ;
                //Generating grouped Chart thumbnail
                GenerateShapeThumbnail(GroupedChart1, fileName);

                //Grouping Shapes of Chart element
                IShape GroupedChart2 = GroupShapes(Chart2, slide);
                fileName = "Chart2";
                //Generating grouped Chart thumbnail
                GenerateShapeThumbnail(GroupedChart2, fileName);

          


            }

        }
    }
}
