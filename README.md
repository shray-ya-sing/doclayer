# DocLayer

A cross-platform library for generating PowerPoint presentations programmatically using OpenXML. DocLayer provides both C# and Python APIs for creating PPTX files with support for themes, slides, shapes, and text formatting.

## Overview

DocLayer Core is built on .NET 8.0 and leverages DocumentFormat.OpenXml to create PowerPoint files without requiring Microsoft Office. The library includes Python bindings via pythonnet for seamless integration with Python applications and AI agent frameworks.

## Features

- Create title slides with custom text and formatting
- Set presentation themes with custom fonts and accent colors
- Cross-platform support (Windows, macOS, Linux via .NET)
- Python wrapper for easy integration with AI agents and data science workflows
- Support for footnotes and slide elements
- Built on industry-standard OpenXML format

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

Install the Python wrapper:

```bash
cd python-wrapper
pip install pythonnet>=3.0.0
pip install -e .
```

Requirements:
- Python 3.8 or higher
- .NET 8.0 Runtime
- pythonnet 3.0.0 or higher

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

### Advanced Python Usage

```python
from doclayer_python import DocLayerClient

client = DocLayerClient()

with client.create_presentation("advanced.pptx") as pres:
    # Set widescreen format
    pres.set_widescreen()
    
    # Add title slide
    slide1 = pres.add_slide()
    slide1.add_title("Quarterly Report")
    slide1.add_textbox("Q4 2024 Analysis")
    slide1.add_footnote("Generated automatically")
    
    # Add content slide with shapes
    slide2 = pres.add_slide()
    slide2.add_title("Key Metrics")
    
    shapes = slide2.get_shape_tree()
    shapes.add_rectangle(1, 2, 2, 3)
    shapes.add_textbox("Revenue: $2.5M", 1, 1)
```

## API Reference

### C# API

#### PresentationBuilder

Main class for building PowerPoint presentations.

**Methods:**

- `CreateTitleSlide(string title, string? subtitle = null, string? footnote = "Source:")` - Creates a title slide with optional subtitle and footnote
- `SetPresentationTheme(string? fontName = null, List<string>? accentColors = null)` - Sets custom theme with font and colors (requires exactly 4 accent colors if provided)

### Python API

#### DocLayerClient

Python client for interacting with DocLayer.Core.

**Methods:**

- `create_title_slide(filepath, title, subtitle=None, footnote="Source:")` - Creates a simple presentation with a title slide, returns bytes
- `create_presentation(filepath)` - Returns a context manager for building complex presentations

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

## Contributing

Contributions are welcome. Please ensure all tests pass before submitting pull requests.

## License

MIT License

## Support

For issues and questions, please refer to the examples directory for comprehensive usage patterns.
