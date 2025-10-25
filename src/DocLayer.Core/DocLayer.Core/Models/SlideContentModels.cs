namespace DocLayer.Core.Models
{
    /// <summary>
    /// Represents the extracted content from a slide
    /// </summary>
    public class SlideContentInfo
    {
        public List<ShapeInfo> Shapes { get; set; } = new();
        public List<PictureInfo> Pictures { get; set; } = new();
    }

    /// <summary>
    /// Represents information about a shape element on a slide
    /// </summary>
    public class ShapeInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public Position? Position { get; set; }
        public Size? Size { get; set; }
        public string? FillColorHex { get; set; }
    }

    /// <summary>
    /// Represents information about a picture element on a slide
    /// </summary>
    public class PictureInfo
    {
        public string Name { get; set; } = string.Empty;
        public Position? Position { get; set; }
        public Size? Size { get; set; }
    }

    /// <summary>
    /// Represents a 2D position in EMUs (English Metric Units)
    /// </summary>
    public class Position
    {
        public long X { get; set; }
        public long Y { get; set; }

        public Position(long x, long y)
        {
            X = x;
            Y = y;
        }
    }

    /// <summary>
    /// Represents dimensions in EMUs (English Metric Units)
    /// </summary>
    public class Size
    {
        public long Width { get; set; }
        public long Height { get; set; }

        public Size(long width, long height)
        {
            Width = width;
            Height = height;
        }
    }

    /// <summary>
    /// Represents updates to apply to a slide element
    /// </summary>
    public class ElementUpdate
    {
        public string? Text { get; set; }
    }
}
