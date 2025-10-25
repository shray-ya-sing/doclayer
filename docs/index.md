# Introduction

DocLayer is a cross-platform library for generating PowerPoint presentations programmatically. It provides C#, Python, and TypeScript/Node.js APIs for creating, editing, and analyzing PPTX files.

## What is DocLayer?

DocLayer enables you to create PowerPoint files without requiring Microsoft Office. The library includes:

- **C# Library** - Native .NET API
- **Python Package** - pip-installable Python library
- **TypeScript Package** - npm-installable Node.js library
- **REST API** - HTTP endpoints for any language
- **MCP Server** - Model Context Protocol server for AI agents

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
Works on Windows, macOS, and Linux with native performance.

### Multiple Language Bindings
Choose your preferred language:
- **C# / .NET** - Direct library access
- **Python** - Native Python API via pythonnet
- **TypeScript/Node.js** - npm package
- **REST API** - HTTP endpoints for any language
- **MCP Server** - For AI agents (Claude, etc.)

### AI Agent Ready
Perfect for AI agent frameworks:
- **MCP Server** - Native integration with Claude Desktop, Claude Code, and other MCP-compatible agents
- LangChain document generation
- CrewAI presentation tasks
- AutoGPT integrations
- Custom agent workflows


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

**MCP Server (Recommended)**:
```json
// Configure Claude Desktop or Claude Code
{
  "mcpServers": {
    "doclayer": {
      "command": "python",
      "args": ["-m", "doclayer_mcp.server"]
    }
  }
}
```

**LangChain Integration**:
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


## Requirements

### For Python
- Python 3.8 or higher

### For TypeScript/Node.js
- Node.js 16.0 or higher

### For C# Development
- .NET 8.0 SDK

## Next Steps

- [Installation Guide](/guide/installation) - Install DocLayer for your platform
- [Quick Start](/guide/getting-started) - Create your first presentation
- [API Reference](/api/csharp) - Detailed API documentation
