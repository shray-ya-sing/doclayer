using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocLayer.Core;
using DocLayer.Core.Models;
using InternalUtilities.Syncfusion;
using doclayer_webapi.Models;
using System.IO.Compression;

namespace doclayer_webapi.Controllers
{
    /// <summary>
    /// API Controller for PowerPoint presentation operations
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class PresentationController : ControllerBase
    {
        private readonly ILogger<PresentationController> _logger;

        public PresentationController(ILogger<PresentationController> logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// Create a new presentation with a title slide
        /// </summary>
        /// <param name="request">Title slide details</param>
        /// <returns>The created presentation file</returns>
        [HttpPost("create-title-slide")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public IActionResult CreateTitleSlide([FromBody] CreateTitleSlideRequest request)
        {
            try
            {
                _logger.LogInformation("Creating title slide presentation");

                var tempFile = Path.GetTempFileName();
                var pptxFile = Path.ChangeExtension(tempFile, ".pptx");

                using (var presentationDoc = PresentationHelper.CreatePresentation(pptxFile, true))
                {
                    var builder = new PresentationBuilder(presentationDoc);
                    builder.CreateTitleSlide(request.Title, request.Subtitle, request.Footnote);
                    presentationDoc.Save();
                }

                var fileBytes = System.IO.File.ReadAllBytes(pptxFile);
                System.IO.File.Delete(pptxFile);

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "presentation.pptx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating title slide");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Create a new presentation with custom theme and title slide
        /// </summary>
        /// <param name="request">Presentation details with theme</param>
        /// <returns>The created presentation file</returns>
        [HttpPost("create-with-theme")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public IActionResult CreatePresentationWithTheme([FromBody] CreatePresentationWithThemeRequest request)
        {
            try
            {
                _logger.LogInformation("Creating presentation with custom theme");

                if (request.AccentColors != null && request.AccentColors.Count != 4)
                {
                    return BadRequest(new { error = "Must provide exactly 4 accent colors" });
                }

                var tempFile = Path.GetTempFileName();
                var pptxFile = Path.ChangeExtension(tempFile, ".pptx");

                using (var presentationDoc = PresentationHelper.CreatePresentation(pptxFile, true))
                {
                    var builder = new PresentationBuilder(presentationDoc);

                    if (request.FontName != null || request.AccentColors != null)
                    {
                        builder.SetPresentationTheme(request.FontName, request.AccentColors);
                    }

                    builder.CreateTitleSlide(request.Title, request.Subtitle, request.Footnote);
                    presentationDoc.Save();
                }

                var fileBytes = System.IO.File.ReadAllBytes(pptxFile);
                System.IO.File.Delete(pptxFile);

                return File(fileBytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "presentation.pptx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating presentation with theme");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Upload a presentation file for processing
        /// </summary>
        /// <param name="file">The presentation file</param>
        /// <returns>File ID for subsequent operations</returns>
        [HttpPost("upload")]
        [ProducesResponseType(typeof(object), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<IActionResult> UploadPresentation(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                {
                    return BadRequest(new { error = "No file uploaded" });
                }

                if (!file.FileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))
                {
                    return BadRequest(new { error = "Only .pptx files are supported" });
                }

                var tempFile = Path.GetTempFileName();
                var pptxFile = Path.ChangeExtension(tempFile, ".pptx");

                using (var stream = new FileStream(pptxFile, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                var fileId = Path.GetFileName(pptxFile);
                _logger.LogInformation("Uploaded file with ID: {FileId}", fileId);

                return Ok(new { fileId, originalName = file.FileName });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error uploading presentation");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Get the number of slides in a presentation
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <returns>Slide count</returns>
        [HttpGet("{fileId}/slide-count")]
        [ProducesResponseType(typeof(SlideCountResponse), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult GetSlideCount(string fileId)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                using var builder = PresentationBuilder.FromFile(filePath, false);
                var count = builder.GetSlideCount();

                return Ok(new SlideCountResponse { Count = count });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting slide count");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Extract content from a specific slide
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <param name="slideNumber">Slide number (1-based)</param>
        /// <returns>Slide content with shapes and pictures</returns>
        [HttpGet("{fileId}/slides/{slideNumber}")]
        [ProducesResponseType(typeof(SlideContentResponse), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult ExtractSlideContent(string fileId, int slideNumber)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                using var builder = PresentationBuilder.FromFile(filePath, false);
                var content = builder.ExtractSlideContent(slideNumber);

                var response = new SlideContentResponse
                {
                    Shapes = content.Shapes.Select(s => new ShapeInfoResponse
                    {
                        Name = s.Name,
                        Text = s.Text,
                        Position = s.Position != null ? new PositionResponse { X = s.Position.X, Y = s.Position.Y } : null,
                        Size = s.Size != null ? new SizeResponse { Width = s.Size.Width, Height = s.Size.Height } : null
                    }).ToList(),
                    Pictures = content.Pictures.Select(p => new PictureInfoResponse
                    {
                        Name = p.Name,
                        Position = p.Position != null ? new PositionResponse { X = p.Position.X, Y = p.Position.Y } : null,
                        Size = p.Size != null ? new SizeResponse { Width = p.Size.Width, Height = p.Size.Height } : null
                    }).ToList()
                };

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting slide content");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Extract content from all slides
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <returns>Dictionary of slide numbers to content</returns>
        [HttpGet("{fileId}/slides")]
        [ProducesResponseType(typeof(Dictionary<int, SlideContentResponse>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult ExtractAllSlides(string fileId)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                using var builder = PresentationBuilder.FromFile(filePath, false);
                var allContent = builder.ExtractAllSlides();

                var response = allContent.ToDictionary(
                    kvp => kvp.Key,
                    kvp => new SlideContentResponse
                    {
                        Shapes = kvp.Value.Shapes.Select(s => new ShapeInfoResponse
                        {
                            Name = s.Name,
                            Text = s.Text,
                            Position = s.Position != null ? new PositionResponse { X = s.Position.X, Y = s.Position.Y } : null,
                            Size = s.Size != null ? new SizeResponse { Width = s.Size.Width, Height = s.Size.Height } : null
                        }).ToList(),
                        Pictures = kvp.Value.Pictures.Select(p => new PictureInfoResponse
                        {
                            Name = p.Name,
                            Position = p.Position != null ? new PositionResponse { X = p.Position.X, Y = p.Position.Y } : null,
                            Size = p.Size != null ? new SizeResponse { Width = p.Size.Width, Height = p.Size.Height } : null
                        }).ToList()
                    }
                );

                return Ok(response);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting all slides");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Edit text on a slide
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <param name="request">Edit details</param>
        /// <returns>Success response</returns>
        [HttpPut("{fileId}/edit-text")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult EditSlideText(string fileId, [FromBody] EditSlideTextRequest request)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                using (var builder = PresentationBuilder.FromFile(filePath, true))
                {
                    builder.EditSlideText(request.SlideNumber, request.ElementName, request.NewText);
                    builder.Save();
                }

                return Ok(new { success = true, message = "Text updated successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error editing slide text");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Download the modified presentation
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <returns>The presentation file</returns>
        [HttpGet("{fileId}/download")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult DownloadPresentation(string fileId)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", "presentation.pptx");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading presentation");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Render a specific slide to JPEG image
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <param name="slideNumber">Slide number (1-based)</param>
        /// <returns>JPEG image</returns>
        [HttpGet("{fileId}/slides/{slideNumber}/render")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult RenderSlideToImage(string fileId, int slideNumber)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                var imagePath = SyncfusionHelperMethods.ExportSlideToImage(filePath, slideNumber);
                var imageBytes = System.IO.File.ReadAllBytes(imagePath);
                System.IO.File.Delete(imagePath);

                return File(imageBytes, "image/jpeg", $"slide_{slideNumber}.jpg");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rendering slide to image");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Render all slides to JPEG images (returns ZIP)
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <returns>ZIP file containing all slide images</returns>
        [HttpGet("{fileId}/slides/render-all")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult RenderAllSlides(string fileId)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                var imagePaths = SyncfusionHelperMethods.ExportPptToImages(filePath);

                // Create ZIP file
                var zipPath = Path.GetTempFileName();
                using (var zip = System.IO.Compression.ZipFile.Open(zipPath, System.IO.Compression.ZipArchiveMode.Create))
                {
                    for (int i = 0; i < imagePaths.Count; i++)
                    {
                        zip.CreateEntryFromFile(imagePaths[i], $"slide_{i + 1}.jpg");
                        System.IO.File.Delete(imagePaths[i]);
                    }
                }

                var zipBytes = System.IO.File.ReadAllBytes(zipPath);
                System.IO.File.Delete(zipPath);

                return File(zipBytes, "application/zip", "slides.zip");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error rendering all slides");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Delete an uploaded presentation
        /// </summary>
        /// <param name="fileId">File ID from upload</param>
        /// <returns>Success response</returns>
        [HttpDelete("{fileId}")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public IActionResult DeletePresentation(string fileId)
        {
            try
            {
                var filePath = Path.Combine(Path.GetTempPath(), fileId);
                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound(new { error = "File not found" });
                }

                System.IO.File.Delete(filePath);
                return Ok(new { success = true, message = "File deleted successfully" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting presentation");
                return BadRequest(new { error = ex.Message });
            }
        }

        /// <summary>
        /// Export an array of json to a powerpoint .pptx file
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("export")]
        [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public IActionResult ExportJsonToPptx([FromBody] ExportRequest request)
        {
            // Basic validation
            if (request?.SlideJsonArray == null || request.SlideJsonArray.Count == 0)
            {
                return BadRequest("No slide data provided.");
            }

            try
            {
                _logger.LogInformation("Exporting {Count} slides to pptx", request.SlideJsonArray.Count);

                var exporter = new PresentationExporter();
                var fileBytes = exporter.ExportPptxBytesFromJsonArray(request.SlideJsonArray);

                // Sanitize the filename (remove invalid chars if necessary)
                var fileName = string.IsNullOrWhiteSpace(request.PresentationName)
                    ? "Presentation"
                    : request.PresentationName;

                return File(
                    fileBytes,  
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    $"{fileName}.pptx"
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting JSON to PPTX");
                return BadRequest(new { error = "An error occurred generating the presentation." });
            }
        }
    }
}
