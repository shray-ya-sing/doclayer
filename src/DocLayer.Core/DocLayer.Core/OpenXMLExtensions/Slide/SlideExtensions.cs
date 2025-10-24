using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using Microsoft.SemanticKernel;

namespace OpenXMLExtensions
{
    public static class SlideExtensions
    {
#pragma warning disable
        public static Shape GetShape(this Slide slide, string shapeName) 
        {
            Dictionary<String, Shape> shapeCollection = new Dictionary<String, Shape>();
            Shape targetShape;

            foreach (Shape shape in slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                NonVisualShapeProperties nvShapeProps = shape.GetFirstChild<NonVisualShapeProperties>();
                NonVisualDrawingProperties nvShapeDrawingProps = nvShapeProps.GetFirstChild<NonVisualDrawingProperties>();
                shapeCollection.Add(nvShapeDrawingProps.Name, shape);
            }

            if (shapeCollection.TryGetValue(shapeName, out targetShape))
            {
                return targetShape;
            }

            else throw new Exception("The given shape name is not valid");
        }

        public static void AddRectangle(this Slide slide)
        {
            if(slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>()!= null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddRectangle(2, 0, 3, 5);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        public static void AddTextbox(this Slide slide, string text)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddTextbox(text, 1, 1);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }

        }

        public static void AddTable(this Slide slide, int numRows, int numCols)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddTable(numRows, numCols);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
            
        }

        public static GraphicFrame GetGraphicFrame(this Slide slide, string tableName)
        {
            Dictionary<string, GraphicFrame> tableCollection = new Dictionary<string, GraphicFrame>();

            foreach (GraphicFrame graphicFrame in slide.CommonSlideData.ShapeTree.Elements<GraphicFrame>())
            {
                tableCollection.Add(graphicFrame.GetName(), graphicFrame);
            }

            GraphicFrame targetGraphicFrame;
            if (tableCollection.TryGetValue(tableName, out targetGraphicFrame))
            {
                return targetGraphicFrame;
            }

            else throw new Exception("The given shape name is not valid");
        }

        public static void AddPicture(this Slide slide, string relId, int hpos, int vpos)
        {            
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddPicture(relId, hpos, vpos);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        public static void AddPicture(this Slide slide, string relId, decimal hpos, decimal vpos)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddPicture(relId, hpos, vpos);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        /// <summary>
        /// Adds a picture to the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="relId"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="hpos"></param>
        /// <param name="vpos"></param>
        /// <exception cref="Exception"></exception>
        public static void AddPicture(this Slide slide, string relId, decimal height, decimal width, decimal hpos, decimal vpos)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddPicture(relId, height, width, hpos, vpos);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        /// <summary>
        /// Gets the picture by its name from the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="pictureName"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static Picture GetPicture(this Slide slide, string pictureName)
        {
            Dictionary<String, Picture> pictureCollection = new Dictionary<String, Picture>();
           
            foreach (Picture picture in slide.CommonSlideData.ShapeTree.Elements<Picture>())
            {
                
                NonVisualPictureProperties? nvPicProps = picture.GetFirstChild<NonVisualPictureProperties>();
                Verify.NotNull(nvPicProps);
                NonVisualDrawingProperties? nvDrawingProps = nvPicProps.GetFirstChild<NonVisualDrawingProperties>();
                Verify.NotNull(nvDrawingProps);
                Verify.NotNull(nvDrawingProps.Name);
                pictureCollection.Add(nvDrawingProps.Name, picture);
            }

            Picture targetPicture;
            if (pictureCollection.TryGetValue(pictureName, out targetPicture))
            {
                return targetPicture;
            }

            else throw new Exception("The given picture name is not valid");
        }

        public static List<Picture> GetPictures(this Slide slide, List<string> pictureNames)
        {
            Dictionary<String, Picture> pictureCollection = new Dictionary<String, Picture>();
            List<Picture> pictures = new();
            foreach (Picture picture in slide.CommonSlideData.ShapeTree.Elements<Picture>())
            {
                NonVisualPictureProperties? nvPicProps = picture.GetFirstChild<NonVisualPictureProperties>();
                Verify.NotNull(nvPicProps);
                NonVisualDrawingProperties? nvDrawingProps = nvPicProps.GetFirstChild<NonVisualDrawingProperties>();
                Verify.NotNull(nvDrawingProps);
                Verify.NotNull(nvDrawingProps.Name);
                pictureCollection.Add(nvDrawingProps.Name, picture);
            }

            foreach (string pictureName in pictureNames) 
            {
                Picture targetPicture;
                if (pictureCollection.TryGetValue(pictureName, out targetPicture))
                {
                    pictures.Add(targetPicture);
                }

                else throw new Exception("The given picture name is not valid");
            }

            return pictures;
           
        }

        /// <summary>
        /// Adds a footnote to the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="footnoteText"></param>
        /// <exception cref="Exception"></exception>
        public static void AddFootnote(this Slide slide, string footnoteText = "Source:")
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddFootnote(footnoteText);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        /// <summary>
        /// Adds a page number to the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="pageNumber"></param>
        /// <exception cref="Exception"></exception>
        public static void AddPageNumber(this Slide slide, string pageNumber)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddPageNumber(pageNumber);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        /// <summary>
        /// Adds a title to the top of the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="titleText"></param>
        /// <exception cref="Exception"></exception>
        public static void AddTitle(this Slide slide, string titleText)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    shapeTree.AddTitle(titleText);
                }

                else
                {
                    throw new Exception("Shapetree not found");
                }
            }

            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }

        /// <summary>
        /// Adds a subtitle to the slide (for title layout slides)
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="subtitleText"></param>
        /// <exception cref="Exception"></exception>
        public static void AddSubtitle(this Slide slide, string subtitleText)
        {
            if (slide.GetFirstChild<CommonSlideData>() != null)
            {
                CommonSlideData cSldData = slide.GetFirstChild<CommonSlideData>();
                if (cSldData.GetFirstChild<ShapeTree>() != null)
                {
                    ShapeTree shapeTree = cSldData.GetFirstChild<ShapeTree>();
                    
                    // Find the subtitle shape placeholder
                    Shape subtitleShape = null;
                    foreach (Shape shape in shapeTree.Elements<Shape>())
                    {
                        PlaceholderShape placeholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<PlaceholderShape>();
                        if (placeholderShape != null && placeholderShape.Type != null && placeholderShape.Type == PlaceholderValues.SubTitle)
                        {
                            subtitleShape = shape;
                            break;
                        }
                    }

                    if (subtitleShape != null)
                    {
                        // Set the text in the subtitle shape
                        subtitleShape.SetText(subtitleText);
                    }
                    else
                    {
                        throw new Exception("Subtitle placeholder not found on slide");
                    }
                }
                else
                {
                    throw new Exception("Shapetree not found");
                }
            }
            else
            {
                throw new Exception("CommonSlideData not found");
            }
        }


        /// <summary>
        /// Sets the title of the slide to the passed text
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="titleText"></param>
        public static void SetTitleText(this Slide slide, string titleText)
        {
            Shape titleShape = slide.CommonSlideData.ShapeTree.Elements<Shape>().First();
            titleShape.SetText(titleText);
            slide.SetTitleFontSize();
        }

        /// <summary>
        /// Sets the font size of the title text
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="size"></param>
        public static void SetTitleFontSize(this Slide slide, int? size = 24)
        {
            Shape titleShape = slide.CommonSlideData.ShapeTree.Elements<Shape>().First();
            var d = titleShape.Descendants<D.Run>().Last();
            d.SetRunSize((int)size);
        }


        /// <summary>
        /// Gets the inner xml associated with each shape in the shapetree of the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <returns> List of xml strings for each shape</returns>
        public static List<string> GetShapesXml(this Slide slide) 
        {
            List<string> shapesInfo = new();
            var shapes = slide.CommonSlideData.ShapeTree.Elements<Shape>();
            foreach (Shape shape in shapes)
            {
                shapesInfo.Add(shape.InnerXml);
            }

            return shapesInfo;
        }

        /// <summary>
        /// Gets the inner xml associated with each shape in the shapetree of the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <returns> List of xml strings for each shape</returns>
        public static List<string> GetPicturesXml(this Slide slide)
        {
            List<string> shapesInfo = new();
            var shapes = slide.CommonSlideData.ShapeTree.Elements<Picture>();
            foreach (Picture shape in shapes)
            {
                shapesInfo.Add(shape.InnerXml);
            }

            return shapesInfo;
        }

        /// <summary>
        /// Creates shapes on the slide from the given inner xml strings for each shape. 
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="shapesInfo"> A list of strings representing the inner xml associated with each shape.</param>

        public static void AddShapesByXml(this Slide slide, List<string> shapesInfo) 
        {
            foreach (string shapeInfo in shapesInfo)
            {
                Shape shape = new();
                shape.InnerXml = shapeInfo;
                slide.CommonSlideData.ShapeTree.AppendChild<Shape>(shape);
            }
        }
        /// <summary>
        /// Creates pictures on the slide from the given inner xml strings for each shape.
        /// The pictures have no blip fill as the image is not provided.
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="shapesInfo"> A list of strings representing the inner xml associated with each shape.</param>
        public static void AddPicturesByXml(this Slide slide, List<string> shapesInfo)
        {
            foreach (string shapeInfo in shapesInfo)
            {
                Picture shape = new();
                shape.InnerXml = shapeInfo;
                slide.CommonSlideData.ShapeTree.AppendChild<Picture>(shape);
            }
        }

        /// <summary>
        /// Gets all the text on the slide
        /// </summary>
        /// <param name="slide"></param>
        /// <returns></returns>

        public static string GetSlideText(this Slide slide) 
        {
            string slideText = "";
            foreach (Shape shape in slide.CommonSlideData.ShapeTree.Elements<Shape>()) 
            {
                if (PresentationHelperMethods.IsTitleShape(shape)) {  slideText = "Title text:" + shape.GetText();}
                slideText += "\r\n" + shape.GetText().Trim();
            }

            return slideText;
        }
    }
}
