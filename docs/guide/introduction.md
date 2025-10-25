# Introduction

DocLayer is a cross-platform library for generating PowerPoint presentations programmatically using OpenXML. It provides C#, Python, and TypeScript/Node.js APIs for creating PPTX files with support for themes, slides, shapes, and text formatting.

## What is DocLayer?

DocLayer Core is built on .NET 8.0 and leverages DocumentFormat.OpenXml to create PowerPoint files without requiring Microsoft Office. The library includes:

- **C# Core Library** - Native .NET implementation
- **Python Wrapper** - via pythonnet for easy Python integration
- **TypeScript Wrapper** - Node.js bindings via Python bridge
- **REST API** - ASP.NET Core Web API for HTTP access

## Key Features

### Create & Edit
Generate title slides with custom text and formatting. Modify text content in existing presentations with the edit API.

### Extract Content
Extract text, shapes, pictures, and metadata from existing presentations for analysis, processing, or migration.

### Render to Images
Convert slides to JPEG images for:
- Thumbnail generation
- Preview images
- Web display
- Batch processing

### Custom Themes
Set presentation themes with:
- Custom fonts (any TrueType font)
- Accent colors (4 customizable colors)
- Consistent branding across presentations

### Cross-Platform
Works on Windows, macOS, and Linux via .NET 8.0 runtime with native performance.

### Multiple Language Bindings
Choose your preferred language:
- **C# / .NET** - Direct library access
- **Python** - Native Python API via pythonnet
- **TypeScript/Node.js** - npm package
- **REST API** - HTTP endpoints for any language

### AI Agent Ready
Perfect for AI agent frameworks:
- LangChain document generation
- CrewAI presentation tasks
- AutoGPT integrations
- Custom agent workflows

## Architecture

```
┌─────────────────────────┐
│   Your Application      │
│  (C#/Python/TS/HTTP)    │
└───────────┬─────────────┘
            │
            v
┌─────────────────────────┐
│   DocLayer Wrapper      │
│  (Language-specific)    │
└───────────┬─────────────┘
            │
            v
┌─────────────────────────┐
│   DocLayer.Core (C#)    │
│  PresentationBuilder    │
└───────────┬─────────────┘
            │
            v
┌─────────────────────────┐
│   OpenXML SDK           │
│  DocumentFormat.OpenXml │
└───────────┬─────────────┘
            │
            v
┌─────────────────────────┐
│   PowerPoint Files      │
│   (.pptx format)        │
└─────────────────────────┘
```

## Use Cases

### Automated Report Generation
Generate presentations from database queries, analytics data, or business intelligence systems.

```python
# Weekly report automation
data = fetch_weekly_metrics()
client.create_presentation_with_theme(
    "weekly_report.pptx",
    title=f"Week {data.week_number} Report",
    subtitle=f"Revenue: ${data.revenue:,.0f}"
)
```

### AI Agent Integration
Use with AI frameworks to generate documents on demand.

```python
from langchain.tools import Tool

def generate_presentation(title: str, content: str) -> str:
    client.create_title_slide("output.pptx", title, content)
    return "Presentation created successfully"

presentation_tool = Tool(
    name="GeneratePresentation",
    func=generate_presentation,
    description="Creates a PowerPoint presentation"
)
```

### Batch Processing
Process hundreds of presentations for content extraction or modification.

```python
import glob

for pptx_file in glob.glob("*.pptx"):
    content = client.extract_all_slides(pptx_file)
    # Process content...
```

### Thumbnail Generation
Create preview images for presentation management systems.

```python
# Generate thumbnails for all slides
images = client.render_all_slides("presentation.pptx")
for i, img_path in enumerate(images, 1):
    upload_thumbnail(f"slide_{i}.jpg", img_path)
```

## Technology Stack

- **.NET 8.0** - Core runtime
- **DocumentFormat.OpenXml 3.3.0** - PowerPoint file generation
- **Syncfusion.Presentation** - Slide rendering to images
- **pythonnet 3.0.0+** - Python interop
- **Node.js 16.0+** - TypeScript wrapper runtime
- **ASP.NET Core 8.0** - REST API

## Requirements

### For C# Development
- .NET 8.0 SDK
- DocumentFormat.OpenXml 3.3.0
- Syncfusion.Presentation.Net.Core 31.2.3 (for rendering)

### For Python
- Python 3.8 or higher
- pythonnet 3.0.0 or higher
- .NET 8.0 Runtime

### For TypeScript/Node.js
- Node.js 16.0 or higher
- Python 3.8+ with doclayer-py installed
- .NET 8.0 Runtime

### For REST API
- .NET 8.0 SDK
- ASP.NET Core Runtime

## Next Steps

- [Installation Guide](/guide/installation) - Install DocLayer for your platform
- [Quick Start](/guide/getting-started) - Create your first presentation
- [API Reference](/api/csharp) - Detailed API documentation
