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

    /// <summary>
    /// Represents Excel source link metadata for an image in PowerPoint
    /// </summary>
    public class ExcelLinkInfo
    {
        /// <summary>
        /// Full file path to the Excel workbook
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Name of the worksheet
        /// </summary>
        public string SheetName { get; set; } = string.Empty;

        /// <summary>
        /// Cell range (e.g., "A1:D10") or chart object name
        /// </summary>
        public string RangeOrChartName { get; set; } = string.Empty;

        /// <summary>
        /// Timestamp when the link was created
        /// </summary>
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        /// <summary>
        /// Serializes to metadata string format: "FilePath|SheetName|Range|Timestamp"
        /// </summary>
        public string ToMetadataString()
        {
            return $"{FilePath}|{SheetName}|{RangeOrChartName}|{CreatedAt:O}";
        }

        /// <summary>
        /// Parses metadata string back to ExcelLinkInfo object
        /// </summary>
        public static ExcelLinkInfo? FromMetadataString(string metadata)
        {
            if (string.IsNullOrWhiteSpace(metadata))
                return null;

            var parts = metadata.Split('|');
            if (parts.Length < 3)
                return null;

            var linkInfo = new ExcelLinkInfo
            {
                FilePath = parts[0],
                SheetName = parts[1],
                RangeOrChartName = parts[2]
            };

            // Parse timestamp if available
            if (parts.Length >= 4 && DateTime.TryParse(parts[3], out DateTime timestamp))
            {
                linkInfo.CreatedAt = timestamp;
            }

            return linkInfo;
        }
    }
}
