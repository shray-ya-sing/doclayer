# C# / .NET API

The C# API is the core DocLayer implementation built on .NET 8.0 and DocumentFormat.OpenXml.

## Installation

Add the DocLayer.Core library to your .NET project:

```bash
dotnet add reference path/to/DocLayer.Core.csproj
```

Or build from source:

```bash
cd src/DocLayer.Core/DocLayer.Core
dotnet build
```

## Quick Start

```csharp
using DocLayer.Core;
using DocumentFormat.OpenXml.Packaging;

// Create a new presentation
using var presentationDoc = PresentationHelper.CreatePresentation(
    "presentation.pptx", 
    true
);

// Initialize the builder
var builder = new PresentationBuilder(presentationDoc);

// Create a title slide
builder.CreateTitleSlide(
    title: "Welcome to DocLayer",
    subtitle: "PowerPoint Generation Made Easy",
    footnote: "Source: DocLayer.Core"
);

presentationDoc.Save();
```

## PresentationBuilder Class

The main class for building PowerPoint presentations.

### Factory Methods

#### `FromFile(string filepath, bool isEditable)`

Opens an existing presentation for reading or editing.

```csharp
// Read-only
using var builder = PresentationBuilder.FromFile("presentation.pptx", false);
var count = builder.GetSlideCount();

// Editable
using var builder = PresentationBuilder.FromFile("presentation.pptx", true);
builder.EditSlideText(1, "Title 1", "New Text");
builder.Save();
```

### Creation Methods

#### `CreateTitleSlide(string title, string? subtitle = null, string? footnote = "Source:")`

Creates a title slide with optional subtitle and footnote.

```csharp
builder.CreateTitleSlide(
    title: "My Presentation",
    subtitle: "A subtitle",
    footnote: "Source: My Company"
);
```

#### `SetPresentationTheme(string? fontName = null, List<string>? accentColors = null)`

Sets custom theme with font and colors. Requires exactly 4 accent colors if provided.

```csharp
builder.SetPresentationTheme(
    fontName: "Arial",
    accentColors: new List<string> { "FF5733", "33FF57", "3357FF", "F3FF33" }
);
```

### Query Methods

#### `GetSlideCount()`

Returns the number of slides in the presentation.

```csharp
int count = builder.GetSlideCount();
Console.WriteLine($"Presentation has {count} slides");
```

#### `ExtractSlideContent(int slideNumber)`

Extracts shapes, text, and pictures from a specific slide.

```csharp
var content = builder.ExtractSlideContent(1);

foreach (var shape in content.Shapes)
{
    Console.WriteLine($"Shape: {shape.Name}");
    Console.WriteLine($"Text: {shape.Text}");
    Console.WriteLine($"Position: ({shape.Position?.X}, {shape.Position?.Y})");
    Console.WriteLine($"Size: {shape.Size?.Width}x{shape.Size?.Height}");
}

foreach (var picture in content.Pictures)
{
    Console.WriteLine($"Picture: {picture.Name}");
}
```

#### `ExtractAllSlides()`

Extracts content from all slides as a dictionary.

```csharp
var allContent = builder.ExtractAllSlides();

foreach (var kvp in allContent)
{
    Console.WriteLine($"Slide {kvp.Key}:");
    Console.WriteLine($"  Shapes: {kvp.Value.Shapes.Count}");
    Console.WriteLine($"  Pictures: {kvp.Value.Pictures.Count}");
}
```

### Edit Methods

#### `EditSlideText(int slideNumber, string elementName, string newText)`

Modifies text of a shape on a slide.

```csharp
using var builder = PresentationBuilder.FromFile("presentation.pptx", true);
builder.EditSlideText(1, "Title 1", "Updated Title");
builder.Save();
```

#### `Save()`

Saves changes to the presentation. Only needed when editing.

```csharp
builder.EditSlideText(1, "Title 1", "New Text");
builder.Save();
```

### Rendering Methods

::: warning
Rendering methods are available through `SyncfusionHelperMethods` after disposing the PresentationBuilder.
:::

```csharp
// Close the builder first
builder.Dispose();

// Then render
using InternalUtilities.Syncfusion;

var imagePath = SyncfusionHelperMethods.ExportSlideToImage("presentation.pptx", 1);
var allImages = SyncfusionHelperMethods.ExportPptToImages("presentation.pptx");
```

## Models

### SlideContentInfo

Contains extracted slide content.

```csharp
public class SlideContentInfo
{
    public List<ShapeInfo> Shapes { get; set; }
    public List<PictureInfo> Pictures { get; set; }
}
```

### ShapeInfo

Information about a shape element.

```csharp
public class ShapeInfo
{
    public string Name { get; set; }
    public string Text { get; set; }
    public Position? Position { get; set; }
    public Size? Size { get; set; }
}
```

### PictureInfo

Information about a picture element.

```csharp
public class PictureInfo
{
    public string Name { get; set; }
    public Position? Position { get; set; }
    public Size? Size { get; set; }
}
```

### Position

2D position coordinates (in EMUs - English Metric Units).

```csharp
public class Position
{
    public long X { get; set; }
    public long Y { get; set; }
}
```

### Size

Width and height dimensions (in EMUs).

```csharp
public class Size
{
    public long Width { get; set; }
    public long Height { get; set; }
}
```

## Complete Example

```csharp
using DocLayer.Core;
using DocLayer.Core.Models;
using DocumentFormat.OpenXml.Packaging;
using InternalUtilities.Syncfusion;

// Create presentation
using (var doc = PresentationHelper.CreatePresentation("demo.pptx", true))
{
    var builder = new PresentationBuilder(doc);
    
    // Set theme
    builder.SetPresentationTheme(
        fontName: "Calibri",
        accentColors: new List<string> { "4472C4", "ED7D31", "A5A5A5", "FFC000" }
    );
    
    // Create title slide
    builder.CreateTitleSlide(
        title: "Quarterly Report",
        subtitle: "Q4 2024 Results",
        footnote: "Source: Finance Department"
    );
    
    doc.Save();
}

// Extract and analyze
using (var builder = PresentationBuilder.FromFile("demo.pptx", false))
{
    int count = builder.GetSlideCount();
    Console.WriteLine($"Slides: {count}");
    
    var content = builder.ExtractSlideContent(1);
    Console.WriteLine($"Shapes: {content.Shapes.Count}");
    Console.WriteLine($"Pictures: {content.Pictures.Count}");
}

// Edit content
using (var builder = PresentationBuilder.FromFile("demo.pptx", true))
{
    builder.EditSlideText(1, "Title 1", "Updated Quarterly Report");
    builder.Save();
}

// Render to images
var images = SyncfusionHelperMethods.ExportPptToImages("demo.pptx");
Console.WriteLine($"Generated {images.Count} slide images");
```

## Requirements

- .NET 8.0 SDK
- DocumentFormat.OpenXml 3.3.0
- Syncfusion.Presentation.Net.Core 31.2.3 (for rendering)
- Syncfusion.PresentationRenderer.Net.Core 31.2.3 (for rendering)

## Next Steps

- [Python API](/api/python) - Python wrapper documentation
- [TypeScript API](/api/typescript) - TypeScript wrapper documentation
- [Web API](/api/webapi) - REST API documentation
