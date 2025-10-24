using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using System.Collections.Generic;

namespace OpenXMLExtensions
{
    public static class PictureCollectionExtensions
    {
        #pragma warning disable
        public static void DistributeVertically(this List<Picture> pictures)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 topMostPos = pictures.First().GetVerticalPosition();
            Int64 bottomMostPos = pictures.Last().GetVerticalPosition();
            var distance = Math.Abs(topMostPos - bottomMostPos);
            var increment = distance / pictures.Count();
            Int64 posCounter = topMostPos;
            foreach (Picture picture in pictures)
            {
                if (topMostPos > bottomMostPos)
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter -= increment;
                }

                else if (bottomMostPos > topMostPos)
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }

                else
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }
            }
        }
        public static void DistributeVertically(this List<Picture> pictures, decimal? distanceBetweenPictures = 0.2m)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 topMostPos = pictures.First().GetVerticalPosition();
            Int64 bottomMostPos = pictures.Last().GetVerticalPosition();
            Int64 distBetween = (Int64)distanceBetweenPictures * 914400;
            var distance = Math.Abs(topMostPos - bottomMostPos);
            var increment = distance / pictures.Count() ;
            if ((distBetween * pictures.Count) < distance) { increment += distBetween; }
            Int64 posCounter = topMostPos;
            
            foreach (Picture picture in pictures)
            {
                if (topMostPos > bottomMostPos)
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter -= increment;
                }

                else if (bottomMostPos > topMostPos)
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }

                else
                {
                    picture.SetVerticalPosition(posCounter);
                    posCounter += increment;
                }
            }
        }

        public static void DistributeVerticallyInShape(this List<Picture> pictures, Shape anchorShape, double? padding = 0.2) 
        {
            pictures.First().SetHorizontalPosition(anchorShape.GetHorizontalPosition());    
            pictures.First().SetVerticalPosition(anchorShape.GetVerticalPosition() + (int)(padding * 914400));
            pictures.Last().SetVerticalPosition(anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() - (int)(padding * 914400));
            pictures.DistributeVertically();
            pictures.AlignLeftToFirst();
        }


        public static void DistributeVerticallyInTable(this List<Picture> pictures, GraphicFrame anchor, decimal? padding = 0.2m, decimal? distanceBetweenPictures= 0.2m)
        {
            pictures.First().SetHorizontalPosition(anchor.GetHorizontalPosition());
            pictures.First().SetVerticalPosition(anchor.GetVerticalPosition() + (int)(padding * 914400));
            pictures.Last().SetVerticalPosition(anchor.GetVerticalPosition() + (Int64)anchor.GetHeight() - (int)(padding * 914400));
            pictures.DistributeVertically(distanceBetweenPictures);
            pictures.AlignLeftToFirst();
        }
        public static void DistributeHorizontally(this List<Picture> pictures, decimal? distanceBetweenPictures = 0.2m)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 firstPos = pictures.First().GetHorizontalPosition();
            Int64 lastPos = pictures.Last().GetHorizontalPosition();
            Int64 distBetween = (Int64)distanceBetweenPictures * 914400;
            Int64 distance = Math.Abs(firstPos - lastPos);
            Int64 increment = distance / pictures.Count();
            if ((distBetween * pictures.Count) < distance) { increment += distBetween; }
            Int64 posCounter = firstPos;
            foreach (Picture picture in pictures)
            {
                if (firstPos > lastPos)
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter -= increment;
                }
                else if (firstPos < lastPos)
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
                else
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
            }
        }

        public static void DistributeHorizontally(this List<Picture> pictures)
        {
            // Starting from this shape to the last element of shapes, distribute the horizontal positions equally
            Int64 firstPos = pictures.First().GetHorizontalPosition();
            Int64 lastPos = pictures.Last().GetHorizontalPosition();
            Int64 distance = Math.Abs(firstPos - lastPos);
            Int64 increment = distance / pictures.Count();
            Int64 posCounter = firstPos;
            foreach (Picture picture in pictures)
            {
                if (firstPos > lastPos)
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter -= increment;
                }
                else if (firstPos < lastPos)
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
                else
                {
                    picture.SetHorizontalPosition(posCounter);
                    posCounter += increment;
                }
            }
        }

        public static void DistributeHorizontallyInShape(this List<Picture> pictures, Shape anchorShape, double? padding = 0.2)
        {
            pictures.First().SetHorizontalPosition(anchorShape.GetHorizontalPosition() + (int)(padding * 914400));
            pictures.Last().SetHorizontalPosition(anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetHeight() - (int)(padding * 914400));
            pictures.DistributeHorizontally();
        }

        public static void DistributeHorizontallyInTable(this List<Picture> pictures, GraphicFrame anchor, decimal? padding = 0.2m)
        {
            pictures.First().SetHorizontalPosition(anchor.GetHorizontalPosition() + (int)(padding * 914400));
            pictures.Last().SetHorizontalPosition(anchor.GetHorizontalPosition() + (Int64)anchor.GetHeight() - (int)(padding * 914400));
            pictures.DistributeHorizontally(padding);
        }

        public static void DistributeAsGridInShape(this List<Picture> pictures, Shape anchor, int ncols, int nrows, double? paddingFromTop = 0.1, double? paddingFromBottom = 0.1, double? paddingFromRight = 0.1, double? paddingFromLeft = 0.1, double? horizontalPaddingBetweenPictures = 0.1, double? verticalPaddingBetweenPictures = 0.1)
        {
            var pictureGrid = CreateGrid(pictures, ncols, nrows);
            int numPictures = pictures.Count;
            // Scale
            // Create the dimensions of the grid, after accounting for padding
            Int64Value topBound = anchor.GetVerticalPosition() + (Int64)(paddingFromTop * 914_400);
            Int64Value bottomBound = anchor.GetVerticalPosition() + (Int64)anchor.GetHeight() - (Int64)(paddingFromBottom * 914_400);
            Int64Value leftBound = anchor.GetHorizontalPosition() + (Int64)(paddingFromLeft * 914_400);
            Int64Value rightBound = anchor.GetHorizontalPosition() + (Int64)anchor.GetWidth() - (Int64)(paddingFromRight * 914_400);
            Int64Value gridWidth = rightBound - (Int64)leftBound;
            Int64Value gridHeight = bottomBound - (Int64)topBound;
            int numPicturesInColumn = nrows;
            int numPicturesInRow = ncols;
            // What should the maximmum size of each picture be? Divide the Height and Width of the anchoring shape by the number of pictures to get the dimensions of each picture
            Int64Value maxPictureHeight = (gridHeight - ((Int64)(verticalPaddingBetweenPictures * 914_400)*(numPicturesInColumn-1))) / numPicturesInColumn;
            Int64Value maxPictureWidth = (gridWidth - ((Int64)(horizontalPaddingBetweenPictures * 914_400))*(numPicturesInRow-1)) / numPicturesInRow;

            // Scale each picture to the optimal height and width
            foreach (Picture picture in pictures) 
            {
                if (picture.GetHeight() > maxPictureHeight) 
                {
                    picture.ScaleHeightTo(maxPictureHeight);
                }

                if (picture.GetWidth() > maxPictureWidth)
                {
                    picture.ScaleWidthTo(maxPictureWidth);
                }
                // picture.ScaleWidthTo(optimalPictureWidth);
                picture.SetAspectRatioLock(false);                
            }

            // Position
            Int64Value verticalSpan = GetGridVerticalSpan(pictures, numPictures, verticalPaddingBetweenPictures) + (Int64)(paddingFromTop * 914_400) + (Int64)(paddingFromBottom * 914_400);
            // use the padding top, bottom, left, and right padding to calculate the vertical and horizontal span of the grid
            Int64Value horizontalSpan = pictures.First().GetWidth() * numPictures + (Int64)(horizontalPaddingBetweenPictures * (numPictures - 1)) + (Int64)paddingFromLeft + (Int64)paddingFromLeft;
            // check if grid overflowing from shape. In that case, should adjust the padding to prevent overflow. Minimum padding of 0.1 inches should always be there to prevent overlap between pictures.
            if (verticalSpan > anchor.GetHeight())
            {
                // adjust padding to be 0.1
                verticalPaddingBetweenPictures = 0.1;
            }

            if (horizontalSpan < anchor.GetWidth())
            {
                horizontalPaddingBetweenPictures = 0.1;
            }

            Int64Value rowVerticalPositionCounter = topBound;
            Int64Value rowHorizontalPositionCounter = leftBound;

            foreach (List<Picture> row in pictureGrid)
            {
                foreach (Picture rowElement in row)
                {
                    rowElement.SetVerticalPosition(rowVerticalPositionCounter);
                    rowElement.SetHorizontalPosition(rowHorizontalPositionCounter);
                    rowHorizontalPositionCounter += (Int64)(horizontalPaddingBetweenPictures * 914_400) + (Int64)maxPictureWidth;
                }
                // Reset horizontal Position Counter
                rowHorizontalPositionCounter = leftBound;
                // Increment Vertical Position Counter
                rowVerticalPositionCounter += (Int64)(verticalPaddingBetweenPictures * 914_400) + (Int64)maxPictureHeight;
            }           

        }

        private static Int64Value GetGridVerticalSpan(List<Picture> pictures, int numPictures, double? verticalPaddingBetweenPictures) 
        {
            Int64Value verticalSpan = pictures.First().GetHeight() * numPictures + (Int64)verticalPaddingBetweenPictures * (numPictures - 1);
            return verticalSpan;
        }

        private static Int64Value GetGridHorizontalSpan(List<Picture> pictures, int numPictures, double horizontalPaddingBetweenPictures) 
        {
            Int64Value horizontalSpan = pictures.First().GetWidth() * numPictures + (Int64)horizontalPaddingBetweenPictures * (numPictures - 1);
            return horizontalSpan;
        }

        public static List<List<Picture>> CreateGrid(this List<Picture> pictures, int ncols, int nrows) 
        {
            if (nrows <= 0 || ncols <= 0)
                throw new ArgumentException("Number of rows and columns must be greater than zero.");

            if (nrows <= 0 || ncols <= 0)
                throw new ArgumentException("Number of rows and columns must be greater than zero.");

            int totalCells = nrows * ncols;
            if (pictures.Count > totalCells)
                throw new ArgumentException("The number of objects exceeds the grid capacity.");


            // Initialize the grid
            List<List<Picture>> grid = new();

            // Arrange objects in the grid
            int index = 0;
            for (int i = 0; i < nrows; i++)
            {
                List<Picture> row = new();
                for (int j = 0; j < ncols; j++)
                {
                    if (index < pictures.Count)
                    {
                        row.Add(pictures[index]);                         
                        index++;
                    }
                }
                grid.Add(row);
            }

            return grid;
        }


        public static void ScaleHeightInShape(this List<Picture> pictures, Shape anchorShape, decimal? padding = 0.2m) 
        {
            Int64Value shapeHeight = anchorShape.GetHeight();
            int numPics = pictures.Count;
            var totalPadding = padding * numPics;
            var height = shapeHeight - totalPadding / numPics;
            foreach (Picture picture in pictures) 
            {
                picture.ScaleHeightTo((double)height);
            }
        }

        public static void AlignLeftToFirst(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            foreach (var item in pictures)
            {
                if (item != firstShape)
                {
                    item.AlignLeft(firstShape);
                }
            }
        }

        public static void AlignRightToFirst(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            foreach (var item in pictures)
            {
                if (item != firstShape)
                {
                    item.AlignRight(firstShape);
                }
            }
        }

        public static void AlignTopToFirst(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            foreach (var item in pictures)
            {
                if (item != firstShape)
                {
                    item.AlignTop(firstShape);
                }
            }
        }

        public static void AlignBottomToFirst(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            foreach (var item in pictures)
            {
                if (item != firstShape)
                {
                    item.AlignBottom(firstShape);
                }
            }
        }

        public static void MiddleAlign(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            foreach (var item in pictures)
            {
                if (item != firstShape)
                {
                    item.AlignMiddle(firstShape);
                }
            }
        }

        public static void AlignVerticalCenters(this List<Picture> pictures) 
        {
            Picture firstShape = pictures[0];
            var centerPosn = firstShape.GetVerticalPosition() + (firstShape.GetHeight())/2;
            foreach (var item in pictures) { item.SetVerticalCenterPosition(centerPosn) ; }
        }

        public static void AlignHorizontalCenters(this List<Picture> pictures)
        {
            Picture firstShape = pictures[0];
            var centerPosn = firstShape.GetHorizontalPosition() + (firstShape.GetWidth()) / 2;
            foreach (var item in pictures) { item.SetHorizontalCenterPosition(centerPosn); }
        }
    }
}
