using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;


namespace OpenXMLExtensions
{
    public static class ShapePropertiesExtensions
    {
        #pragma warning disable
        
        /// <summary>
        /// Sets the width of the outline
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="width"></param>
        public static void SetOutlineWidth(this P.ShapeProperties shapeProperties, int width)
        {
            if (shapeProperties.GetFirstChild<D.Outline>() != null) 
            {
                D.Outline outline = shapeProperties.GetFirstChild<D.Outline>();
                if (outline.GetFirstChild<D.SolidFill>() is null)
                {
                    D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = D.SchemeColorValues.Accent1 });
                    outline.AddChild(solidFill);
                }
                outline.Width = width * 12700;
            }
            else
            {                
                D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = D.SchemeColorValues.Accent1 });
                D.Outline outline = new D.Outline() { Width = width * 12700 };
                outline.AddChild(solidFill);
                shapeProperties.AddChild(outline);
            }
        }

        public static void SetOutlineWidth(this P.ShapeProperties shapeProperties, double width)
        {
            if (shapeProperties.GetFirstChild<D.Outline>() != null)
            {
                D.Outline outline = shapeProperties.GetFirstChild<D.Outline>();
                if (outline.GetFirstChild<D.SolidFill>() is null)
                {
                    D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = D.SchemeColorValues.Accent1 });
                    outline.AddChild(solidFill);
                }
                outline.Width = (Int32)(width * 12700);
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = D.SchemeColorValues.Accent1 });
                D.Outline outline = new D.Outline() { Width = (Int32)(width * 12700) };
                outline.AddChild(solidFill);
                shapeProperties.AddChild(outline);
            }
        }

        /// <summary>
        /// Sets the horizontal position of the shape from top left corner
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="pts"></param>
        public static void SetHorizontalPosition(this P.ShapeProperties shapeProperties, int inches)
        {
            if (inches < 0) throw new Exception("Please enter a valid number for position");
            
            if(shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.X = inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.X = inches * 914400;
                    transform.AddChild(offset);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.X = inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetHorizontalPosition(this P.ShapeProperties shapeProperties, decimal inches)
        {
            if (inches < 0) throw new Exception("Please enter a valid number for position");

            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.X = (Int64)inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.X = (Int64)inches * 914400;
                    transform.AddChild(offset);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.X = (Int64)inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetHorizontalPosition(this P.ShapeProperties shapeProperties, double inches)
        {
            if (inches < 0) throw new Exception("Please enter a valid number for position");

            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.X = (Int64)inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.X = (Int64)inches * 914400;
                    transform.AddChild(offset);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.X = (Int64)inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }
        /// <summary>
        /// Sets the horizontal position in EMU
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="pts"></param>
        public static void SetHorizontalPosition(this P.ShapeProperties shapeProperties, Int64Value pts)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.X = pts;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.X = pts;
                    transform.AddChild(offset);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.X = pts;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }

        /// <summary>
        /// Gets the horizontal position in EMU
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>

        public static Int64Value? GetHorizontalPosition(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    if (offset.X != null)
                    {
                        Int64Value hPos = offset.X;
                        return hPos;
                    }
                    else
                    {
                        return null;
                    }

                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }





        /// <summary>
        /// Set vertical position of shape from top left corner
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="pts"></param>

        public static void SetVerticalPosition(this P.ShapeProperties shapeProperties, int inches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.Y = inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.Y = inches * 914400;
                    transform.AddChild(offset);
                }
                
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.Y = inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetVerticalPosition(this P.ShapeProperties shapeProperties, decimal inches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.Y = (Int64)inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.Y = (Int64)inches * 914400;
                    transform.AddChild(offset);
                }

            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.Y = (Int64)inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetVerticalPosition(this P.ShapeProperties shapeProperties, double inches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.Y = (Int64)inches * 914400;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.Y = (Int64)inches * 914400;
                    transform.AddChild(offset);
                }

            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.Y = (Int64)inches * 914400;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }
        public static void SetVerticalPosition(this P.ShapeProperties shapeProperties, Int64Value pts)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    offset.Y = pts;
                }
                else
                {
                    D.Offset offset = new D.Offset();
                    offset.Y = pts;
                    transform.AddChild(offset);
                }

            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Offset offset = new D.Offset();
                offset.Y = pts;
                transform.AddChild(offset);
                shapeProperties.AddChild(transform);
            }
        }


        public static Int64Value? GetVerticalPosition(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    if (offset.Y != null)
                    {
                        Int64Value vPos = offset.Y;
                        return vPos;
                    }
                    else
                    {
                        return null;
                    }

                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Sets width of shape
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="pts"></param>

        public static void SetWidth(this P.ShapeProperties shapeProperties, int inches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cx = inches * 914400;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cx = inches * 914400;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cx = inches * 914400;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetWidth(this P.ShapeProperties shapeProperties, decimal widthInInches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cx = (Int64)(widthInInches * 914400);
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cx = (Int64)(widthInInches * 914400);
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cx = (Int64)(widthInInches * 914400);
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetWidth(this P.ShapeProperties shapeProperties, double widthInInches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cx = (Int64)(widthInInches * 914400);
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cx = (Int64)(widthInInches * 914400);
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cx = (Int64)(widthInInches * 914400);
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetWidth(this P.ShapeProperties shapeProperties, Int64Value pts)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cx = pts;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cx = pts;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cx = pts;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }


        public static Int64Value? GetWidth(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    if (extents.Cx != null)
                    {
                        Int64Value width = extents.Cx;
                        return width;
                    }
                    else
                    {
                        return null;
                    }

                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Sets Height of the shape
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="pts"></param>

        public static void SetHeight(this P.ShapeProperties shapeProperties, int inches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cy = inches * 914400;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cy = inches * 914400;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cy = inches * 914400;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetHeight(this P.ShapeProperties shapeProperties, decimal heightInInches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cy = (Int64)heightInInches * 914400;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cy = (Int64)heightInInches * 914400;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cy = (Int64)heightInInches * 914400;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static void SetHeight(this P.ShapeProperties shapeProperties, double heightInInches)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cy = (Int64)heightInInches * 914400;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cy = (Int64)heightInInches * 914400;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cy = (Int64)heightInInches * 914400;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }
        public static void SetHeight(this P.ShapeProperties shapeProperties, Int64Value pts)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cy = pts;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cy = pts;
                    transform.AddChild(extents);
                }
            }
            else
            {
                D.Transform2D transform = new D.Transform2D();
                D.Extents extents = new D.Extents();
                extents.Cy = pts;
                transform.AddChild(extents);
                shapeProperties.AddChild(transform);
            }
        }

        public static Int64Value? GetHeight(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    Console.WriteLine("FOUND: " + extents.Cy.ToString());
                    if (extents.Cy != null)
                    {
                        Int64Value height = extents.Cy;
                        return height;
                    }
                    else
                    {
                        return null;
                    }

                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }



        /// <summary>
        /// Sets preset geometry of shape
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="shapeType"></param>

        public static void SetPresetGeometry(this P.ShapeProperties shapeProperties, D.ShapeTypeValues shapeType)
        {
            if (shapeProperties.GetFirstChild<D.PresetGeometry>() != null)
            {                
                D.PresetGeometry geom = shapeProperties.GetFirstChild<D.PresetGeometry>();
                geom.Preset = shapeType;
                if (geom.GetFirstChild<D.AdjustValueList>() is null)
                {
                    geom.AddChild(new D.AdjustValueList());
                }
            }
            else
            {
                D.PresetGeometry geom = new D.PresetGeometry();
                geom.Preset = shapeType;
                geom.AddChild(new D.AdjustValueList());
                shapeProperties.AddChild(geom);
            }
        }

        public static void SetHarveyBallGuides(this P.ShapeProperties shapeProperties, double percentAdjustment)
        {
            if (shapeProperties.GetFirstChild<D.PresetGeometry>() != null)
            {
                D.PresetGeometry geom = shapeProperties.GetFirstChild<D.PresetGeometry>();

                if (geom.Preset == D.ShapeTypeValues.Pie)
                {
                    int startGuide = 16200000;
                    int oneFourth = 10816111;
                    int oneHalf = 5432222;
                    int threeFourth = 21599805;

                    int endGuide;
                    int diff;

                    if (percentAdjustment > 0 && percentAdjustment <= 0.5)
                    {
                        diff = startGuide - oneHalf;
                        endGuide = (int)Math.Round(startGuide - (percentAdjustment * 2 * diff));
                        shapeProperties.SetShapeGuide(startGuide, endGuide);
                    }

                    else if (percentAdjustment > 0.5 && percentAdjustment <= 0.75)
                    {
                        diff = oneFourth - oneHalf;
                        double adj = percentAdjustment - 0.5;
                        endGuide = (int)Math.Round(oneHalf - (adj * 2 * diff));
                        shapeProperties.SetShapeGuide(startGuide, endGuide);
                    }
                    else if (percentAdjustment > 0.75 && percentAdjustment <= 0.1)
                    {
                        diff = threeFourth - startGuide;
                        double adj = (percentAdjustment - 0.75);
                        endGuide = (int)Math.Round(threeFourth - (adj * 4 * diff));
                        shapeProperties.SetShapeGuide(startGuide, endGuide);
                    }
                }                
            }      
        }

        /// <summary>
        /// Sets the shape guides required for a harvey ball
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="startGuide"></param>
        /// <param name="endGuide"></param>
        /// <exception cref="Exception"></exception>

        public static void SetShapeGuide(this P.ShapeProperties shapeProperties, int startGuide, int endGuide)
        {
            if (shapeProperties.GetFirstChild<D.PresetGeometry>() != null)
            {
                D.PresetGeometry geom = shapeProperties.GetFirstChild<D.PresetGeometry>();

                if (geom.GetFirstChild<D.AdjustValueList>() != null)
                {
                    D.AdjustValueList avList = geom.GetFirstChild<D.AdjustValueList>();
                    D.ShapeGuide shapeGuide1 = new D.ShapeGuide() { Name = "adj1", Formula = "val " + endGuide.ToString() };
                    D.ShapeGuide shapeGuide2 = new D.ShapeGuide() { Name = "adj2", Formula = "val " + startGuide.ToString() };
                    avList.AddChild(shapeGuide1);
                    avList.AppendChild<D.ShapeGuide>(shapeGuide2 );
                }
                else throw new Exception("AdjustValueList not found");
            }
            else throw new Exception("Preset geometry not found");
        }
            
    
        /// <summary>
        /// Rotates shape by degree
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="percentRotation"></param>
        public static void SetRotation(this P.ShapeProperties shapeProperties, int rotationAngleDegree)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if(transform.Rotation != null)
                {
                    transform.Rotation+= rotationAngleDegree * 60000;
                }
                else
                {
                    transform.Rotation = rotationAngleDegree * 60000;
                }              

            }
            else throw new Exception("Transform2D not found");
        }

        /// <summary>
        /// Gets rotation angle of shape in pts
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static int GetRotation(this P.ShapeProperties shapeProperties)
        {
            int rotation;
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                if (transform.Rotation != null)
                {
                    rotation = transform.Rotation;
                    return rotation;
                }
                else
                {
                    return 0;
                }

            }
            else throw new Exception("Transform2D not found");
        }


        /// <summary>
        /// Flips shape vertically
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="rotationAngleDegree"></param>
        /// <exception cref="Exception"></exception>

        public static void FlipVertically(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                transform.VerticalFlip = true;
            }
            else throw new Exception("Transform2D not found");
        }

        /// <summary>
        /// Flips shape horizontally
        /// </summary>
        /// <param name="shapeProperties"></param>
        /// <param name="rotationAngleDegree"></param>
        /// <exception cref="Exception"></exception>
        public static void FlipHorizontally(this P.ShapeProperties shapeProperties)
        {
            if (shapeProperties.GetFirstChild<D.Transform2D>() != null)
            {
                D.Transform2D transform = shapeProperties.GetFirstChild<D.Transform2D>();
                transform.HorizontalFlip = true;
            }
            else throw new Exception("Transform2D not found");
        }

        /// <summary>
        /// Sets fill by hex code
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rgbColorHex"></param>
        public static void SetHexFill(this P.ShapeProperties props, string rgbColorHex)
        {            
            if (props.GetFirstChild<D.SolidFill>() is not null)
            {
                D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new RgbColorModelHex() { Val = rgbColorHex }
                   );

                props.ReplaceChild<SolidFill>
                    (newSolidFill, solidFill);


            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(
                    new RgbColorModelHex() { Val = rgbColorHex }
                    );
                props.AddChild(solidFill);
            }

        }

        /// <summary>
        /// Sets fill by scheme color
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="accentNum"></param>

        public static void SetSchemeFill(this P.ShapeProperties props, int accentNum)
        {
            D.SchemeColorValues schemeColorVal;

            switch (accentNum)
            {
                case 1:
                    schemeColorVal = D.SchemeColorValues.Accent1;
                    break;

                case 2:
                    schemeColorVal = D.SchemeColorValues.Accent2;
                    break;

                case 3:
                    schemeColorVal = D.SchemeColorValues.Accent3;
                    break;

                case 4:
                    schemeColorVal = D.SchemeColorValues.Accent4;
                    break;
            }

            if (props.GetFirstChild<D.SolidFill>() is not null)
            {
                D.SolidFill solidFill = props.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new SchemeColor() { Val = schemeColorVal }
                   );

                props.ReplaceChild<SolidFill>(newSolidFill, solidFill);
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(
                   new SchemeColor() { Val = schemeColorVal }
                   );
                props.AddChild(solidFill);
            }

        }

        /// <summary>
        /// Sets Outline color by Hexcode
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="rgbColorHex"></param>
        public static void SetOutlineHexFill(this P.ShapeProperties props, string rgbColorHex)
        {
            if (props.GetFirstChild<D.Outline>() is not null)
            {
                D.Outline outline = props.GetFirstChild<D.Outline>();
                D.SolidFill solidFill = outline.GetFirstChild<D.SolidFill>();
                D.SolidFill newSolidFill = new D.SolidFill(
                   new RgbColorModelHex() { Val = rgbColorHex }
                   );

                outline.ReplaceChild<SolidFill>(newSolidFill, solidFill);
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(
                    new RgbColorModelHex() { Val = rgbColorHex }
                    );
                D.Outline outline = new D.Outline(solidFill);
                props.AddChild(outline);

            }

        }

        /// <summary>
        /// Sets outline fill by scheme color
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="accentNum"></param>
        public static void SetOutlineSchemeFill(this P.ShapeProperties props, int accentNum)
        {
            //IMP: Method assumes shape props not null

            D.SchemeColorValues schemeColorVal;

            switch (accentNum)
            {
                case 1:
                    schemeColorVal = D.SchemeColorValues.Accent1;
                    break;

                case 2:
                    schemeColorVal = D.SchemeColorValues.Accent2;
                    break;

                case 3:
                    schemeColorVal = D.SchemeColorValues.Accent3;
                    break;

                case 4:
                    schemeColorVal = D.SchemeColorValues.Accent4;
                    break;

            }

            if (props.GetFirstChild<D.Outline>() != null)
            {
                D.Outline outline = props.GetFirstChild<D.Outline>();
                if (outline.GetFirstChild<D.SolidFill>() != null)
                {
                    D.SolidFill solidFill = outline.GetFirstChild<D.SolidFill>();
                    D.SolidFill newSolidFill = new D.SolidFill(new SchemeColor() { Val = schemeColorVal });
                    outline.ReplaceChild<SolidFill>(newSolidFill, solidFill);
                }
                else
                {
                    D.SolidFill newSolidFill = new D.SolidFill(new SchemeColor() { Val = schemeColorVal });
                    outline.AddChild(newSolidFill);
                }
            }
            else
            {
                D.SolidFill solidFill = new D.SolidFill(new SchemeColor() { Val = schemeColorVal });
                D.Outline outline = new D.Outline();
                outline.AddChild(solidFill);
                props.AddChild(outline);
            }

        }
        //TODO: Fix int conversion, values are messed up
        public static void SetFillTransparency(this P.Shape shape, int transparency)
        {
            int opacity = (100 - transparency) * 1000;

            D.SolidFill solidFill = shape.ShapeProperties.GetFirstChild<D.SolidFill>();
            if (solidFill != null)
            {
                solidFill.Append(new D.Alpha() { Val = opacity });
            }
            else
            {
                shape.ShapeProperties.Append(new D.SolidFill(
                    new D.SchemeColor() { Val = SchemeColorValues.Accent1 },
                    new D.Alpha() { Val = opacity }
                    )
                   );
            }

        }
    }
}
