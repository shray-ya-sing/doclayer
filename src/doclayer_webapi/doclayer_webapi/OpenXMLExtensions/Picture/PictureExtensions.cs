using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Runtime.CompilerServices;
using Microsoft.SemanticKernel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;


namespace OpenXMLExtensions
{
    public static class PictureExtensions
    {
#pragma warning disable
        public static string GetName(this P.Picture picture)
        {
            P.NonVisualPictureProperties props = picture.GetFirstChild<P.NonVisualPictureProperties>();
            P.NonVisualDrawingProperties nvDrawingProps = props.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvDrawingProps.Name.ToString();
        }

        public static UInt32Value GetId(this P.Picture picture)
        {
            P.NonVisualPictureProperties props = picture.GetFirstChild<P.NonVisualPictureProperties>();
            P.NonVisualDrawingProperties nvDrawingProps = props.GetFirstChild<P.NonVisualDrawingProperties>();

            return nvDrawingProps.Id;
        }

        public static string GetElementClass(this P.Picture picture)
        {
            string name = picture.GetName();
            int end = name.LastIndexOf(' ');
            string elementClass = name.Substring(0, end);
            return elementClass;
        }

        /// <summary>
        /// Scales to given height while maintaining aspect ratio with proportionate width scaling
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="height"></param>
        public static void ScaleHeightTo(this P.Picture picture, Int64Value scaledHeight)
        {
            Int64Value pictureHeight = picture.GetHeight(); // the height before scaling
            // To preserve the picture integrity, we shoudl scale the width by the exact factor that we are scaling the height
            double scalingFactor = (double)scaledHeight / (double)pictureHeight; // implicit widening problem: ALWAYS convert numerator and denominator to the wider numeric. Otherwise results in rounding errors
            picture.SetHeight(scaledHeight);
            Int64Value scaledWidth = (Int64)(picture.GetWidth() * scalingFactor);
            picture.SetWidth(scaledWidth);
        }

        public static void ScaleHeightTo(this P.Picture picture, int height)
        {
            Int64Value pictureHeight = picture.GetHeight();
            double scalingFactor = (double)(height * 914400) / (double)pictureHeight;
            picture.SetHeight(height);
            Int64Value scaledWidth = (Int64)(picture.GetWidth() * scalingFactor);
            picture.SetWidth(scaledWidth);
        }

        public static void ScaleHeightTo(this P.Picture picture, double height)
        {
            Int64Value pictureHeight = picture.GetHeight();
            double scalingFactor = (double)(height * 914400) / (double)pictureHeight;
            Int64Value newWidth = (Int64Value)(picture.GetWidth() * scalingFactor);
            picture.SetHeight(height);
            Int64Value scaledWidth = (Int64)(picture.GetWidth() * scalingFactor);
            picture.SetWidth(scaledWidth);
        }

        /// <summary>
        /// Scale the picture's size by an integral factor
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="factor"></param>
        public static void ScaleSizeBy(this P.Picture picture, double factor)
        {
            Int64Value height = picture.GetHeight();
            Int64Value newHeight = (Int64Value)(factor * (double)height);
            picture.SetHeight(newHeight);
            Int64Value width = picture.GetWidth();
            Int64Value newWidth = (Int64Value)(factor * (double)width);
            picture.SetWidth(newWidth);
        }

        /// <summary>
        /// Scales to given width while maintaining aspect ratio with proportionate width scaling
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="width"></param>

        public static void ScaleWidthTo(this P.Picture picture, Int64Value width)
        {
            Int64Value pictureWidth = picture.GetWidth();
            double scalingFactor = (double)width / (double)pictureWidth;
            picture.SetWidth(width);
            Int64Value scaledHeight = (Int64)(picture.GetHeight() * scalingFactor);
            picture.SetHeight(scaledHeight);
        }
        public static void ScaleWidthTo(this P.Picture picture, int width)
        {
            Int64Value pictureWidth = picture.GetWidth();
            var scalingFactor = (double)(width * 914_400) / (double)pictureWidth;
            picture.SetWidth(width);
            Int64Value scaledHeight = (Int64)(picture.GetHeight() * scalingFactor);
            picture.SetHeight(scaledHeight);
        }

        public static void ScaleWidthTo(this P.Picture picture, double width)
        {
            Int64Value pictureWidth = picture.GetWidth();
            var scalingFactor = (width * 914400) / (double)pictureWidth;
            picture.SetWidth(width);
            Int64Value scaledHeight = (Int64)(picture.GetHeight() * scalingFactor);
            picture.SetHeight(scaledHeight);
            
        }

        /// <summary>
        /// Sets width
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetWidth(this P.Picture picture, int inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetWidth(inches);
        }

        public static void SetWidth(this P.Picture picture, Int64Value pts)
        {
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetWidth(pts);
        }

        public static void SetWidth(this P.Picture picture, decimal width)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetWidth(width);
        }

        public static void SetWidth(this P.Picture picture, double width)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetWidth(width);
        }

        public static Int64Value GetWidth(this P.Picture picture)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            return shapeProps.GetWidth();
        }


        /// <summary>
        /// Sets shape height
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHeight(this P.Picture picture, int inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHeight(inches);
        }

        public static void SetHeight(this P.Picture picture, Int64Value pts)
        {

            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHeight(pts);
        }

        public static void SetHeight(this P.Picture picture, decimal height)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHeight(height);
        }
        public static void SetHeight(this P.Picture picture, double height)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHeight(height);
        }
        public static Int64Value GetHeight(this P.Picture picture)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            return shapeProps.GetHeight();
        }

        /// <summary>
        /// Sets horizontal position
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="inches"></param>
        public static void SetHorizontalPosition(this P.Picture picture, int inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHorizontalPosition(inches);
        }

        public static void SetHorizontalPosition(this P.Picture picture, Int64Value pts)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHorizontalPosition(pts);
        }

        public static void SetHorizontalPosition(this P.Picture picture, decimal inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHorizontalPosition(inches);
        }

        public static void SetHorizontalPosition(this P.Picture picture, double inches)
        {
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetHorizontalPosition(inches);
        }

        /// <summary>
        /// Gets the horizontal position in inches
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        /// <exception cref="NullReferenceException"></exception>
        public static Int64Value GetHorizontalPosition(this P.Picture picture)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            return shapeProps.GetHorizontalPosition();
        }

        public static void SetHorizontalPositionRelativeTo(this P.Picture picture, P.Shape anchorShape, int relativeDistance)
        {
            picture.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + relativeDistance * 914400);
        }

        public static void SetHorizontalPositionRelativeTo(this P.Picture picture, P.Picture anchorShape, int relativeDistance)
        {
            picture.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + relativeDistance * 914400);
        }

        public static void SetHorizontalPositionRelativeTo(this P.Picture picture, P.Shape anchorShape, double relativeDistance)
        {
            picture.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + (Int64)(relativeDistance * 914400));
        }

        public static void SetHorizontalPositionRelativeTo(this P.Picture picture, P.Picture anchorShape, double relativeDistance)
        {
            picture.SetHorizontalPosition(pts: anchorShape.GetHorizontalPosition() + (Int64)anchorShape.GetWidth() + (Int64)(relativeDistance * 914400));
        }

        public static void SetVerticalPosition(this P.Picture picture, int inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetVerticalPosition(inches);
        }
        public static void SetVerticalPosition(this P.Picture picture, Int64Value pts)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetVerticalPosition(pts);
        }

        public static void SetVerticalPosition(this P.Picture picture, double inches)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            shapeProps.SetVerticalPosition(inches);
        }
        public static Int64Value GetVerticalPosition(this P.Picture picture)
        {
            Verify.NotNull(picture.GetFirstChild<P.ShapeProperties>());
            var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
            return shapeProps.GetVerticalPosition();
        }

        public static void SetVerticalPositionRelativeTo(this P.Picture picture, P.Shape anchorShape, int relativeDistance)
        {
            picture.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + relativeDistance * 914400);
        }

        public static void SetVerticalPositionRelativeTo(this P.Picture picture, P.Picture anchorShape, int relativeDistance)
        {
            picture.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + relativeDistance * 914400);
        }

        public static void SetVerticalPositionRelativeTo(this P.Picture picture, P.Shape anchorShape, double relativeDistance)
        {
            picture.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + (Int64)(relativeDistance * 914400));
        }

        public static void SetVerticalPositionRelativeTo(this P.Picture picture, P.Picture anchorShape, double relativeDistance)
        {
            picture.SetVerticalPosition(pts: anchorShape.GetVerticalPosition() + (Int64)anchorShape.GetHeight() + (Int64)(relativeDistance * 914400));
        }

        public static void AlignLeft(this P.Picture picture, P.Shape anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            picture.SetHorizontalPosition(anchorShape.GetHorizontalPosition());
        }

        public static void AlignLeft(this P.Picture picture, P.Picture anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            picture.SetHorizontalPosition(anchorShape.GetHorizontalPosition());
        }


        public static void AlignTop(this P.Picture picture, P.Shape anchorShape)
        {
            // Anchor shpae SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            picture.SetVerticalPosition(anchorShape.GetVerticalPosition());
        }

        public static void AlignTop(this P.Picture picture, P.Picture anchorShape)
        {
            // Anchor shpae SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            picture.SetVerticalPosition(anchorShape.GetVerticalPosition());
        }

        public static void AlignBottom(this P.Picture picture, P.Shape anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 vPos = (Int64)anchorShape.GetVerticalPosition();
            Int64 height = (Int64)anchorShape.GetHeight();
            Int64 bottomPos = vPos + height;
            picture.SetVerticalPosition(bottomPos - picture.GetHeight());
        }

        public static void AlignBottom(this P.Picture picture, P.Picture anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 vPos = (Int64)anchorShape.GetVerticalPosition();
            Int64 height = (Int64)anchorShape.GetHeight();
            Int64 bottomPos = vPos + height;
            picture.SetVerticalPosition(bottomPos - picture.GetHeight());
        }

        public static void AlignRight(this P.Picture picture, P.Shape anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 hPos = (Int64)anchorShape.GetHorizontalPosition();
            Int64 width = (Int64)anchorShape.GetWidth();
            Int64 rightPos = hPos + width;
            picture.SetHorizontalPosition(rightPos - picture.GetWidth());
        }

        public static void AlignRight(this P.Picture picture, P.Picture anchorShape)
        {
            // Anchor shape SHOULD NOT BE A TITLE OR SUBTITLE
            // Set horizontal position equal to that of another shape -- will align left
            Int64 hPos = (Int64)anchorShape.GetHorizontalPosition();
            Int64 width = (Int64)anchorShape.GetWidth();
            Int64 rightPos = hPos + width;
            picture.SetHorizontalPosition(rightPos - picture.GetWidth());
        }

        /// <summary>
        /// Sets the vertical center of the picture object to the passed arg
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="centerPosition"></param>
        public static void SetVerticalCenterPosition(this P.Picture picture, Int64Value centerPosition) 
        {
            Int64Value height = picture.GetHeight();
            var centerPoint = height / 2;
            picture.SetVerticalPosition(centerPosition - centerPoint);
        }
        /// <summary>
        /// Sets the horizontal center of the picture object to the passed arg
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="centerPosition"></param>
        public static void SetHorizontalCenterPosition(this P.Picture picture, Int64Value centerPosition)
        {
            Int64Value width = picture.GetWidth();
            var centerPoint = width / 2;
            picture.SetHorizontalPosition(centerPosition - centerPoint);
        }

        public static void AlignMiddle(this P.Picture picture, P.Picture anchorShape) 
        {
            picture.SetVerticalCenterPosition(anchorShape.GetVerticalPosition() + anchorShape.GetHeight()/2);
        }

        public static void AlignMiddle(this P.Picture picture, P.Shape anchorShape)
        {
            picture.SetVerticalCenterPosition(anchorShape.GetVerticalPosition() + anchorShape.GetHeight() / 2);
        }

        public static void AlignCenter(this P.Picture picture, P.Picture anchorShape)
        {
            var center = anchorShape.GetHorizontalPosition() + anchorShape.GetWidth() / 2;
            picture.SetHorizontalCenterPosition(center);
        }

        public static void AlignCenter(this P.Picture picture, P.Shape anchorShape)
        {
            var center = anchorShape.GetHorizontalPosition() + anchorShape.GetWidth() / 2;
            picture.SetHorizontalCenterPosition(center);
        }

        /// <summary>
        /// Moves the picture vertically by the given distance
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="factor"></param>
        public static void MoveVertical(this P.Picture picture, double distance) 
        {
            Int64Value v = picture.GetVerticalPosition();
            var y = (Int64)(distance * 914_400);
            Int64Value vnew = v + y;
            picture.SetVerticalPosition(vnew);
        }

        /// <summary>
        /// Moves the picture horizontally by the given distance
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="distance"></param>
        public static void MoveHorizontal(this P.Picture picture, double distance)
        {
            var h = picture.GetHorizontalPosition();
            var x = (Int64)(distance * 914_400);
            Int64Value hnew = h + x;
            picture.SetHorizontalPosition(hnew);
        }

        public static void SwapTopLeftPosition(this P.Picture picture, P.Picture anchor) 
        {
            Int64Value x1 = picture.GetHorizontalPosition();
            Int64Value y1 = picture.GetVerticalPosition();
            Int64Value x2 = anchor.GetHorizontalPosition();
            Int64Value y2 = anchor.GetVerticalPosition();

            picture.SetHorizontalPosition(x2);
            picture.SetVerticalPosition(y2);

            anchor.SetHorizontalPosition(x1);
            anchor.SetVerticalPosition(y1);
        }

        public static void SwapCenterPosition(this P.Picture picture, P.Picture anchor)
        {
            Int64Value x1 = picture.GetHorizontalPosition() + (picture.GetWidth()/2);
            Int64Value y1 = picture.GetVerticalPosition() + (picture.GetHeight() / 2);
            Int64Value x2 = anchor.GetHorizontalPosition() + (anchor.GetWidth() / 2);
            Int64Value y2 = anchor.GetVerticalPosition() + (anchor.GetHeight() / 2);

            picture.SetHorizontalPosition(x2);
            picture.SetVerticalPosition(y2);

            anchor.SetHorizontalPosition(x1);
            anchor.SetVerticalPosition(y1);
        }

        /// <summary>
        /// Sets an aspect ratio lock if bool argument is true
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="setLocked"></param>
        public static void SetAspectRatioLock(this P.Picture picture, bool setLocked) 
        {
            NonVisualPictureProperties props = picture.GetFirstChild<NonVisualPictureProperties>();
            if (props != null) 
            {
                NonVisualPictureDrawingProperties drawingProps = props.GetFirstChild<NonVisualPictureDrawingProperties>();
                if (drawingProps!= null) 
                {
                    PictureLocks pictureLocks = drawingProps.GetFirstChild<PictureLocks>();
                    if (pictureLocks != null)
                    {
                        if (setLocked)
                        {
                            pictureLocks.NoChangeAspect = true;
                        }
                        else pictureLocks.NoChangeAspect = false;
                    }

                    else 
                    {
                        PictureLocks newPictureLocks = new();
                        if (setLocked)
                        {
                            newPictureLocks.NoChangeAspect = true;
                        }
                        else newPictureLocks.NoChangeAspect = false;
                        drawingProps.Append(newPictureLocks);
                    }
                }
            }
            

        }

        /// <summary>
        /// Applies grayscale filter to picture
        /// </summary>
        /// <param name="picture"></param>
        public static void SetGrayScale(this P.Picture picture) 
        {
            D.Grayscale grayscale = new();

            if (picture.GetFirstChild<P.BlipFill>() != null) 
            {
                P.BlipFill blipFill = picture.GetFirstChild<P.BlipFill>();

                if (blipFill.GetFirstChild<D.Blip>() != null) 
                {
                    blipFill.AppendChild(grayscale);
                }
            }
        }

        /// <summary>
        /// Crops picture to circle
        /// </summary>
        /// <param name="picture"></param>
        public static void CropToCircle(this P.Picture picture) 
        {

            if (picture.GetFirstChild<P.ShapeProperties>() != null) 
            {
                var shapeProps = picture.GetFirstChild<P.ShapeProperties>();
                shapeProps.SetPresetGeometry(D.ShapeTypeValues.Ellipse);
            }
        
        }

        /// <summary>
        /// Applies offset to left and right edges to crop image to 1x1 aspect ratio
        /// </summary>
        /// <param name="picture"></param>
        public static void CropTo1x1AspectRatio(this P.Picture picture)
        {
            // Left and Right offset values represent offset from the left and right edges, respectively, in percentage terms
            // i.e a left offset of 7125 implies an offset from the left edge of 7.125%
            // 1_000 = 1%

            var width = picture.GetWidth();
            var height = picture.GetHeight();

            // Only crop if not already 1x1
            if (width != height) 
            {
                // Adjusting only width, not height. This is a choice to maintain cropping effect closest to powerpoint and prevent image distortion

                double aspectRatio = height / width; // Inverting formula for computation

                // Cut off the excess from both sides
                double requiredOffset = ((1 - aspectRatio) / 2) * 1_000;

                Int32Value leftOffset = (Int32Value)requiredOffset;

                Int32Value rightOffset = (Int32Value)requiredOffset;

                if (picture.GetFirstChild<P.ShapeProperties>() != null)
                {
                    P.BlipFill blipFill = picture.GetFirstChild<P.BlipFill>();

                    if (blipFill.GetFirstChild<D.Blip>() != null)
                    {
                        blipFill.AppendChild(new D.SourceRectangle() { Left = leftOffset, Right = rightOffset });
                    }
                }

            }

        }

        /// <summary>
        /// Crops an image according to an aspect ratio
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="width"> The desired aspect ratio width </param>
        /// <param name="height"> The desired aspect ratio height </param>
        public static void CropToAspectRatio(this P.Picture picture, int width, int height)
        {
            // Left and Right offset values represent offset from the left and right edges, respectively, in percentage terms
            // i.e a left offset of 7125 implies an offset from the left edge of 7.125%
            // 1_000 = 1%

            Int64Value currentWidth = picture.GetWidth();
            Int64Value currentHeight = picture.GetHeight();
            double currentAspectRatio = currentWidth / currentHeight;
            double requiredAspectRatio = width / height;

            // Crop along both x and y axes
            if (currentAspectRatio != requiredAspectRatio)
            {
                double requiredWidth = requiredAspectRatio * currentHeight;
                double widthScalingFactor = requiredWidth / currentWidth;
                if (widthScalingFactor < 1) 
                {
                    double horizontalOffset = ((1 - widthScalingFactor) / 2) * 1_000;

                    Int32Value leftOffset = (Int32Value)horizontalOffset;

                    Int32Value rightOffset = (Int32Value)horizontalOffset;

                    picture.CropImageHorizontally(leftOffset, rightOffset);
                }

                // Increasing image width not allowed. Cut off vertically if ideal width exceeds current width.

                if (widthScalingFactor > 1)
                {
                    double requiredHeight = currentWidth / requiredAspectRatio;
                    double heightScalingFactor = requiredHeight / currentHeight;
                    double verticalOffset = ((1 - heightScalingFactor) / 2) * 1_000;

                    Int32Value topOffset = (Int32Value)verticalOffset;

                    Int32Value bottomOffset = (Int32Value)verticalOffset;

                    picture.CropImageHorizontally(topOffset, bottomOffset);
                }
            }
        }

        /// <summary>
        /// Crops image along x-axis
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="leftOffset"> Offset from left edge. 1% = 1,000. Integer values should be adjusted for the desired percentage offset  </param>
        /// <param name="rightOffset"> Offset from right edge. 1% = 1,000. Integer values should be adjusted for the desired percentage offset </param>
        public static void CropImageHorizontally(this P.Picture picture, int leftOffset, int rightOffset) 
        {
            if (picture.GetFirstChild<P.ShapeProperties>() != null)
            {
                P.BlipFill blipFill = picture.GetFirstChild<P.BlipFill>();

                if (blipFill.GetFirstChild<D.Blip>() != null)
                {
                    blipFill.AppendChild(new D.SourceRectangle() { Left = leftOffset, Right = rightOffset });
                }
            }
        }

        /// <summary>
        /// Crops image along y-axis
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="topOffset"> Offset from top edge. 1% = 1,000. Integer values should be adjusted for the desired percentage offset </param>
        /// <param name="bottomOffset"> Offset from bottom edge. 1% = 1,000. Integer values should be adjusted for the desired percentage offset </param>
        public static void CropImageVertically(this P.Picture picture, int topOffset, int bottomOffset)
        {
            if (picture.GetFirstChild<P.ShapeProperties>() != null)
            {
                P.BlipFill blipFill = picture.GetFirstChild<P.BlipFill>();

                if (blipFill.GetFirstChild<D.Blip>() != null)
                {
                    blipFill.AppendChild(new D.SourceRectangle() { Top = topOffset, Bottom = bottomOffset });
                }
            }
        }

        /// <summary>
        /// Replaces the data content of an existing image element with a new image referenced by the relId.
        /// </summary>
        /// <param name="picture"></param>
        /// <param name="relid">Reference to the new image part.</param>
        public static void ReplaceImageBlipFill(this P.Picture picture, string relid) 
        {
            if (picture.GetFirstChild<P.BlipFill>() != null)
            {
                P.BlipFill blipFill = picture.GetFirstChild<P.BlipFill>();
                D.Blip blip = blipFill.GetFirstChild<D.Blip>();
                blip.Embed = relid;
            }
        }

    }
}