# DocLayer

A cross-platform library for generating PowerPoint presentations programmatically using OpenXML. DocLayer provides C#, Python (`doclayer-py`), and TypeScript/Node.js (`@doclayer/ts`) APIs for creating PPTX files with support for themes, slides, shapes, and text formatting.

## Overview

DocLayer Core is built on .NET 8.0 and leverages DocumentFormat.OpenXml to create PowerPoint files without requiring Microsoft Office. The library includes Python bindings via pythonnet and a TypeScript/Node.js wrapper for seamless integration with Python applications, Node.js services, and AI agent frameworks.

## Features

- **Create & Edit**: Generate title slides with custom text and formatting
- **Extract Content**: Extract text, shapes, and metadata from existing presentations
- **Edit Slides**: Modify text content in existing presentations
- **Render to Images**: Convert slides to JPEG images for thumbnails and previews
- **Custom Themes**: Set presentation themes with custom fonts and accent colors
- **Cross-platform**: Windows, macOS, Linux support via .NET 8.0
- **Multiple Language Bindings**: Python and TypeScript/Node.js wrappers
- **AI Agent Ready**: Perfect for LangChain, CrewAI, AutoGPT integrations
- **Industry Standard**: Built on OpenXML format

## Installation

### C# / .NET

Add the DocLayer.Core library to your .NET project:

```bash
dotnet add reference path/to/DocLayer.Core.csproj
```

Or build from source:

```bash
cd src/DocLayer.Core/DocLayer.Core
dotnet build
```

### Python

Install the Python package:

```bash
pip install doclayer-py
```

Or install from source:

```bash
cd python-wrapper
pip install pythonnet>=3.0.0
pip install -e .
```

Requirements:
- Python 3.8 or higher
- .NET 8.0 Runtime
- pythonnet 3.0.0 or higher

### TypeScript / Node.js

Install the TypeScript package:

```bash
npm install @doclayer/ts
```

Or install from source:

```bash
cd typescript-wrapper
npm install
npm run build
```

Requirements:
- Node.js 16.0 or higher
- Python 3.8+ with `doclayer-py` package installed
- .NET 8.0 Runtime

## Usage

### C# Example

```csharp
using DocLayer.Core;
using DocumentFormat.OpenXml.Packaging;

// Create a new presentation
using var presentationDoc = PresentationDocument.Create(
    "presentation.pptx", 
    PresentationDocumentType.Presentation
);

// Initialize the builder
var builder = new PresentationBuilder(presentationDoc);

// Set custom theme
builder.SetPresentationTheme(
    fontName: "Arial",
    accentColors: new List<string> { "4472C4", "ED7D31", "A5A5A5", "FFC000" }
);

// Create a title slide
builder.CreateTitleSlide(
    title: "Welcome to DocLayer",
    subtitle: "PowerPoint Generation Made Easy",
    footnote: "Source: DocLayer.Core"
);
```

### Python Example

```python
from doclayer_python import DocLayerClient

# Initialize the client
client = DocLayerClient()

# Create a title slide
pptx_bytes = client.create_title_slide(
    filepath="output.pptx",
    title="My Presentation",
    subtitle="Created with Python",
    footnote="Source: My Data"
)

print(f"Created presentation: {len(pptx_bytes)} bytes")
```

### Advanced Python Usage - Extract & Edit

```python
from doclayer_python import DocLayerClient

client = DocLayerClient()

# Get slide count
count = client.get_slide_count("presentation.pptx")
print(f"Slides: {count}")

# Extract content from a specific slide
content = client.extract_slide_content("presentation.pptx", slide_number=1)
for shape in content['shapes']:
    print(f"Shape: {shape['name']} - Text: {shape['text']}")

# Extract all slides
all_content = client.extract_all_slides("presentation.pptx")
for slide_num, slide_content in all_content.items():
    print(f"Slide {slide_num}: {len(slide_content['shapes'])} shapes")

# Edit text on a slide
client.edit_slide_text(
    "presentation.pptx",
    slide_number=1,
    element_name="Title 1",
    new_text="Updated Title"
)

# Render slide to image
image_path = client.render_slide_to_image("presentation.pptx", slide_number=1)
print(f"Rendered to: {image_path}")

# Render all slides to images
image_paths = client.render_all_slides("presentation.pptx")
print(f"Rendered {len(image_paths)} slides")
```

### TypeScript / Node.js Example

```typescript
import { DocLayerClient } from '@doclayer/ts';

const client = new DocLayerClient();

// Create a simple title slide
const buffer = await client.createTitleSlide('presentation.pptx', {
  title: 'Welcome to DocLayer',
  subtitle: 'PowerPoint Generation from Node.js',
  footnote: 'Source: DocLayer TypeScript Wrapper'
});

console.log(`Created presentation: ${buffer.length} bytes`);

// Get slide count
const count = await client.getSlideCount('presentation.pptx');
console.log(`Slides: ${count}`);

// Extract content from a slide
const content = await client.extractSlideContent('presentation.pptx', 1);
content.shapes.forEach(shape => {
  console.log(`Shape: ${shape.name} - Text: ${shape.text}`);
});

// Edit text on a slide
await client.editSlideText(
  'presentation.pptx',
  1,
  'Title 1',
  'Updated Title'
);

// Render slide to image
const imagePath = await client.renderSlideToImage('presentation.pptx', 1);
console.log(`Rendered to: ${imagePath}`);

// Render all slides
const imagePaths = await client.renderAllSlides('presentation.pptx');
console.log(`Rendered ${imagePaths.length} slides`);
```

## API Reference

### C# API

#### PresentationBuilder

Main class for building PowerPoint presentations.

**Methods:**

- `CreateTitleSlide(string title, string? subtitle = null, string? footnote = "Source:")` - Creates a title slide
- `SetPresentationTheme(string? fontName = null, List<string>? accentColors = null)` - Sets custom theme
- `GetSlideCount()` - Returns the number of slides in the presentation
- `ExtractSlideContent(int slideNumber)` - Extracts shapes, text, and pictures from a slide
- `ExtractAllSlides()` - Extracts content from all slides as a dictionary
- `EditSlideText(int slideNumber, string elementName, string newText)` - Edits text of a shape
- `RenderSlideToImage(int slideNumber)` - Renders a slide to JPEG image
- `RenderAllSlidesToImages()` - Renders all slides to JPEG images

### Python API

#### DocLayerClient

Python client for interacting with DocLayer.Core.

**Methods:**

- `create_title_slide(filepath, title, subtitle=None, footnote="Source:")` - Creates a presentation with a title slide
- `create_presentation_with_theme(filepath, title, subtitle=None, footnote="Source:", font_name=None, accent_colors=None)` - Creates presentation with custom theme
- `get_slide_count(filepath)` - Returns the number of slides
- `extract_slide_content(filepath, slide_number)` - Extracts content from a specific slide
- `extract_all_slides(filepath)` - Extracts content from all slides
- `edit_slide_text(filepath, slide_number, element_name, new_text)` - Edits text of a shape
- `render_slide_to_image(filepath, slide_number)` - Renders a slide to JPEG, returns image path
- `render_all_slides(filepath)` - Renders all slides to JPEG images, returns list of paths

#### TypeScript API

TypeScript/Node.js client that uses Python bridge to generate presentations.

**DocLayerClient Methods:**

- `createTitleSlide(filepath, options)` - Creates a presentation with a title slide
- `createPresentationWithTheme(filepath, options)` - Creates presentation with custom theme
- `getSlideCount(filepath)` - Returns the number of slides
- `extractSlideContent(filepath, slideNumber)` - Extracts content from a specific slide
- `extractAllSlides(filepath)` - Extracts content from all slides
- `editSlideText(filepath, slideNumber, elementName, newText)` - Edits text of a shape
- `renderSlideToImage(filepath, slideNumber)` - Renders a slide to JPEG, returns image path
- `renderAllSlides(filepath)` - Renders all slides to JPEG images, returns array of paths
- `checkEnvironment()` - Checks if Python and DocLayer dependencies are available

## Project Structure

```
doclayer/
├── src/
│   ├── DocLayer.Core/          # Core C# library
│   │   └── DocLayer.Core/
│   │       ├── PresentationBuilder.cs
│   │       └── DocLayer.Core.csproj
│   └── doclayer_webapi/        # Web API wrapper
├── python-wrapper/             # Python bindings
│   ├── doclayer_python/
│   ├── setup.py
│   └── README.md
├── typescript-wrapper/         # TypeScript/Node.js bindings
│   ├── src/
│   ├── test/
│   ├── package.json
│   └── README.md
├── examples/
│   ├── python_example.py       # Python usage examples
│   └── typescript_example.ts
├── test/
│   └── TestTitleSlide/         # C# unit tests
└── README.md
```

## Requirements

### C# Development
- .NET 8.0 SDK
- DocumentFormat.OpenXml 3.3.0
- Microsoft.SemanticKernel 1.66.0 (optional)
- Syncfusion.Presentation.Net.Core 31.2.3 (optional)

### Python Development
- Python 3.8+
- pythonnet 3.0.0+
- .NET 8.0 Runtime

### TypeScript/Node.js Development
- Node.js 16.0+
- Python 3.8+ with `doclayer-py` installed
- .NET 8.0 Runtime

## Use Cases

- AI agent document generation (LangChain, CrewAI, AutoGPT)
- Automated report generation
- Data visualization and dashboards
- Cloud-based presentation services
- Batch PowerPoint creation from data sources
- Integration with analytics pipelines

## Testing

### C# Tests
```bash
cd test/TestTitleSlide
dotnet test
```

### Python Tests
```bash
cd python-wrapper
python test_wrapper.py
```

### TypeScript Tests
```bash
cd typescript-wrapper
npm install
npm run build
npm test
```

## Contributing

Contributions are welcome. Please ensure all tests pass before submitting pull requests.

## License

MIT License

## Support

For issues and questions, please refer to the examples directory for comprehensive usage patterns.
