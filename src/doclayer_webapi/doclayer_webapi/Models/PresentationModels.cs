using DocLayer.Core;
using System.Text.Json.Serialization;

namespace doclayer_webapi.Models
{
    /// <summary>
    /// Request model for creating a title slide presentation
    /// </summary>
    public class CreateTitleSlideRequest
    {
        public required string Title { get; set; }
        public string? Subtitle { get; set; }
        public string? Footnote { get; set; }
    }

    /// <summary>
    /// Request model for exporting an array of JSON to Pptx file
    /// </summary>
    public class ExportRequest
    {
        [JsonPropertyName("slideJsonArray")]
        public required List<SlideSchema> SlideJsonArray { get; set; }

        [JsonPropertyName("presentationName")]
        public string? PresentationName { get; set; }
    }

    /// <summary>
    /// Request model for creating a presentation with custom theme
    /// </summary>
    public class CreatePresentationWithThemeRequest
    {
        public required string Title { get; set; }
        public string? Subtitle { get; set; }
        public string? Footnote { get; set; }
        public string? FontName { get; set; }
        public List<string>? AccentColors { get; set; }
    }

    /// <summary>
    /// Request model for editing slide text
    /// </summary>
    public class EditSlideTextRequest
    {
        public required int SlideNumber { get; set; }
        public required string ElementName { get; set; }
        public required string NewText { get; set; }
    }

    /// <summary>
    /// Response model for slide content extraction
    /// </summary>
    public class SlideContentResponse
    {
        public List<ShapeInfoResponse> Shapes { get; set; } = new();
        public List<PictureInfoResponse> Pictures { get; set; } = new();
    }

    /// <summary>
    /// Response model for shape information
    /// </summary>
    public class ShapeInfoResponse
    {
        public required string Name { get; set; }
        public required string Text { get; set; }
        public PositionResponse? Position { get; set; }
        public SizeResponse? Size { get; set; }
    }

    /// <summary>
    /// Response model for picture information
    /// </summary>
    public class PictureInfoResponse
    {
        public required string Name { get; set; }
        public PositionResponse? Position { get; set; }
        public SizeResponse? Size { get; set; }
    }

    /// <summary>
    /// Response model for position
    /// </summary>
    public class PositionResponse
    {
        public long X { get; set; }
        public long Y { get; set; }
    }

    /// <summary>
    /// Response model for size
    /// </summary>
    public class SizeResponse
    {
        public long Width { get; set; }
        public long Height { get; set; }
    }

    /// <summary>
    /// Response model for render operations
    /// </summary>
    public class RenderResponse
    {
        public required string ImagePath { get; set; }
    }

    /// <summary>
    /// Response model for batch render operations
    /// </summary>
    public class RenderAllResponse
    {
        public required List<string> ImagePaths { get; set; }
    }

    /// <summary>
    /// Response model for slide count
    /// </summary>
    public class SlideCountResponse
    {
        public int Count { get; set; }
    }
}
