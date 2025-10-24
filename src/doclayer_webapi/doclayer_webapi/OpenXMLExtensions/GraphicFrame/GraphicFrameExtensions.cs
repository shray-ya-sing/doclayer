using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Runtime.CompilerServices;


namespace OpenXMLExtensions
{
    public static class GraphicFrameExtensions
    {
        public static string GetName(this P.GraphicFrame graphicFrame)
        {
            P.NonVisualGraphicFrameProperties gfProps = graphicFrame.GetFirstChild<P.NonVisualGraphicFrameProperties>();
            P.NonVisualDrawingProperties nvDrawingProps = gfProps.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvDrawingProps.Name.ToString();
        }

        public static UInt32Value GetId(this P.GraphicFrame graphicFrame)
        {
            P.NonVisualGraphicFrameProperties gfProps = graphicFrame.GetFirstChild<P.NonVisualGraphicFrameProperties>();
            P.NonVisualDrawingProperties nvDrawingProps = gfProps.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvDrawingProps.Id;
        }

        public static string GetElementClass(this P.GraphicFrame graphicFrame)
        {
            string name = graphicFrame.GetName();
            int end = name.LastIndexOf(' ');
            string elementClass = name.Substring(0, end);
            return elementClass;
        }
        public static D.Table GetTable(this P.GraphicFrame graphicFrame)
        {
            //    D.Graphic graphic = graphicFrame.GetFirstChild<D.Graphic>();
            return graphicFrame.Descendants<D.Table>().First();
        }

        /// <summary>
        /// Sets width
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetWidth(this P.GraphicFrame graphicFrame, int inches)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    extents.Cx = inches * 914_400;
                }
                else
                {
                    D.Extents extents = new D.Extents();
                    extents.Cx = inches * 914_400;
                    transform.AddChild(extents);
                }

            }
            else
            {
                P.Transform transform = new P.Transform();
                D.Extents extents = new D.Extents();
                extents.Cx = inches * 914_400;
                transform.AddChild(extents);
                graphicFrame.AddChild(transform);
            }
        }

        public static void SetWidth(this P.GraphicFrame graphicFrame, Int64Value pts)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
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
                P.Transform transform = new P.Transform();
                D.Extents extents = new D.Extents();
                extents.Cx = pts;
                transform.AddChild(extents);
                graphicFrame.AddChild(transform);
            }
        }

        public static Int64Value GetWidth(this P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
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
                        throw new NullReferenceException();
                    }
                }

                else
                {
                    throw new NullReferenceException();
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }


        /// <summary>
        /// Sets shape height
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHeight(this P.GraphicFrame graphicFrame, int inches)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
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
                P.Transform transform = new P.Transform();
                D.Extents extents = new D.Extents();
                extents.Cy = inches * 914400;
                transform.AddChild(extents);
                graphicFrame.AddChild(transform);
            }
        }

        public static void SetHeight(this P.GraphicFrame graphicFrame, Int64Value pts)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
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
                P.Transform transform = new P.Transform();
                D.Extents extents = new D.Extents();
                extents.Cy = pts;
                transform.AddChild(extents);
                graphicFrame.AddChild(transform);
            }
        }

        public static Int64Value GetHeight(this P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
                {
                    D.Extents extents = transform.GetFirstChild<D.Extents>();
                    if (extents.Cy != null)
                    {
                        Int64Value width = extents.Cx;
                        return width;
                    }
                    else
                    {
                        throw new NullReferenceException();
                    }
                }

                else
                {
                    throw new NullReferenceException();
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        /// <summary>
        /// Sets horizontal position
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHorizontalPosition(this P.GraphicFrame graphicFrame, int inches)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
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
                P.Transform transform = new P.Transform();
                D.Offset offset = new D.Offset();
                offset.X = inches * 914400;
                transform.AddChild(offset);
                graphicFrame.AddChild(transform);
            }
        }

        public static void SetHorizontalPosition(this P.GraphicFrame graphicFrame, Int64Value pts)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
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
                P.Transform transform = new P.Transform();
                D.Offset offset = new D.Offset();
                offset.X = pts;
                transform.AddChild(offset);
                graphicFrame.AddChild(transform);
            }
        }
        /// <summary>
        /// Gets the horizontal position in inches
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        public static Int64Value GetHorizontalPosition(this P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    if (offset.X != null)
                    {
                        Int64Value hpos = offset.X;
                        return hpos;
                    }
                    else
                    {
                        throw new NullReferenceException();
                    }
                }

                else
                {
                    throw new NullReferenceException();
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void SetHorizontalPositionRelativeTo(this P.GraphicFrame graphicFrame, P.Shape anchorShape, int relativeDistance)
        {
            graphicFrame.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + relativeDistance * 914400);
        }

        public static void SetVerticalPosition(this P.GraphicFrame graphicFrame, int inches)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
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
                P.Transform transform = new P.Transform();
                D.Offset offset = new D.Offset();
                offset.Y = inches * 914400;
                transform.AddChild(offset);
                graphicFrame.AddChild(transform);
            }
        }
        public static void SetVerticalPosition(this P.GraphicFrame graphicFrame, Int64Value pts)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() is not null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Extents>() != null)
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
                P.Transform transform = new P.Transform();
                D.Offset offset = new D.Offset();
                offset.Y = pts;
                transform.AddChild(offset);
                graphicFrame.AddChild(transform);
            }
        }
        public static Int64Value GetVerticalPosition(this P.GraphicFrame graphicFrame)
        {
            if (graphicFrame.GetFirstChild<P.Transform>() != null)
            {
                P.Transform transform = graphicFrame.GetFirstChild<P.Transform>();
                if (transform.GetFirstChild<D.Offset>() != null)
                {
                    D.Offset offset = transform.GetFirstChild<D.Offset>();
                    if (offset.Y != null)
                    {
                        Int64Value vpos = offset.Y;
                        return vpos;
                    }
                    else
                    {
                        throw new NullReferenceException();
                    }
                }

                else
                {
                    throw new NullReferenceException();
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }

        public static void SetVerticalPositionRelativeTo(this P.GraphicFrame graphicFrame, P.Shape anchorShape, int relativeDistance)
        {
            graphicFrame.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + relativeDistance * 914400);
        }

    }
}
