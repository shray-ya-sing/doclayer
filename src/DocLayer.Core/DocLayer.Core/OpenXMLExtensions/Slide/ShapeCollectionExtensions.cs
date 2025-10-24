using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OpenXMLExtensions
{
    public static class ShapeCollectionExtensions
    {
        public static void DistributeVertically(this List<P.Shape> shapes)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 topMostPos = shapes.First().GetVerticalPosition();
            Int64 bottomMostPos = shapes.Last().GetVerticalPosition();
            Int64 distance = Math.Abs(topMostPos - bottomMostPos);
            Int64 increment = distance / shapes.Count();
            Int64 posCounter = topMostPos;
            foreach (P.Shape shape in shapes)
            {
                if (topMostPos > bottomMostPos)
                {
                    shape.SetVerticalPosition(posCounter);
                    posCounter -= increment;
                }

                else if (bottomMostPos > topMostPos)
                {
                    shape.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }

                else
                {                    
                    shape.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }
            }
        }

        public static void DistributeHorizontally(this List<P.Shape> shapes)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 firstPos = shapes.First().GetHorizontalPosition();
            Int64 lastPos = shapes.Last().GetHorizontalPosition();
            Int64 distance = Math.Abs(firstPos - lastPos);
            Int64 increment = distance / shapes.Count();
            Int64 posCounter = firstPos;
            foreach (P.Shape shape in shapes)
            {
                if(firstPos > lastPos)
                {
                    shape.SetHorizontalPosition(posCounter);
                    posCounter -= increment;
                }
                else if (firstPos < lastPos)
                {
                    shape.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
                else
                {
                    shape.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
                
            }
        }

        public static void CopyFirstShapeFill(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {                
                if ( shape != firstShape)
                {
                    shape.CopyShapeFill(firstShape);
                }
            }
        }

        public static void CopyFirstShapeOutline(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {                
                if (shape != firstShape)
                {
                    shape.CopyOutline(firstShape);
                }
            }
        }

        public static void CopyFirstShapeOutlineFill(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {                
                if (shape != firstShape)
                {
                    shape.CopyOutlineFill (firstShape);
                }
            }
        }

        public static void CopyFirstShapeDimensions(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {                
                if (shape != firstShape)
                {
                    shape.CopyDimensions(firstShape);
                }
            }
        }

        public static void AlignTopToFirst(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {
                if (shape != firstShape)
                {
                    shape.AlignTop(firstShape);
                }
            }
        }
        public static void AlignBottomToFirst(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {
                if (shape != firstShape)
                {
                    shape.AlignBottom(firstShape);
                }
            }
        }

        public static void AlignLeftToFirst(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {
                if (shape != firstShape)
                {
                    shape.AlignLeft(firstShape);
                }
            }
        }

        public static void AlignRightToFirst(this List<P.Shape> shapes)
        {
            P.Shape firstShape = shapes[0];
            foreach (P.Shape shape in shapes)
            {
                if (shape != firstShape)
                {
                    shape.AlignRight(firstShape);
                }
            }
        }
    }
}
