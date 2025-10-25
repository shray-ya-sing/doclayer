---
layout: home

hero:
  name: "DocLayer"
  text: "PowerPoint Generation Made Easy"
  tagline: Cross-platform library for creating, editing, and rendering PowerPoint presentations
  actions:
    - theme: brand
      text: Get Started
      link: /guide/getting-started
    - theme: alt
      text: View on GitHub
      link: https://github.com/shray-ya-sing/doclayer

features:
  - title: Create & Edit
    details: Generate title slides with custom text and formatting. Modify existing presentations with ease.
  
  - title: Extract Content
    details: Extract text, shapes, and metadata from existing presentations for analysis and processing.
  
  - title: Render to Images
    details: Convert slides to JPEG images for thumbnails, previews, and web display.
  
  - title: Custom Themes
    details: Set presentation themes with custom fonts and accent colors to match your brand.
  
  - title: Cross-Platform
    details: Works on Windows, macOS, and Linux via .NET 8.0 with native performance.
  
  - title: Multiple Bindings
    details: Available for C#, Python, TypeScript/Node.js, and REST API. Choose your favorite language.
  
  - title: AI Agent Ready
    details: Perfect for LangChain, CrewAI, AutoGPT integrations with simple, intuitive APIs.
  
  - title: Industry Standard
    details: Built on OpenXML format with DocumentFormat.OpenXml and Syncfusion rendering.
---

## Quick Example

::: code-group

```python [Python]
from doclayer_python import DocLayerClient

client = DocLayerClient()

# Create presentation
client.create_title_slide(
    filepath="presentation.pptx",
    title="My Presentation",
    subtitle="Created with Python"
)

# Extract content
content = client.extract_slide_content("presentation.pptx", 1)
print(f"Found {len(content['shapes'])} shapes")

# Render to image
image = client.render_slide_to_image("presentation.pptx", 1)
```

```typescript [TypeScript]
import { DocLayerClient } from '@doclayer/ts';

const client = new DocLayerClient();

// Create presentation
await client.createTitleSlide('presentation.pptx', {
  title: 'My Presentation',
  subtitle: 'Created with TypeScript'
});

// Extract content
const content = await client.extractSlideContent('presentation.pptx', 1);
console.log(`Found ${content.shapes.length} shapes`);

// Render to image
const image = await client.renderSlideToImage('presentation.pptx', 1);
```

```csharp [C#]
using DocLayer.Core;
using DocumentFormat.OpenXml.Packaging;

// Create presentation
using var doc = PresentationHelper.CreatePresentation("presentation.pptx", true);
var builder = new PresentationBuilder(doc);

builder.CreateTitleSlide(
    title: "My Presentation",
    subtitle: "Created with C#"
);

doc.Save();

// Extract content
using var builder2 = PresentationBuilder.FromFile("presentation.pptx", false);
var content = builder2.ExtractSlideContent(1);
Console.WriteLine($"Found {content.Shapes.Count} shapes");
```

```bash [REST API]
# Create presentation
curl -X POST "https://api.yourdomain.com/api/presentation/create-title-slide" \
  -H "Content-Type: application/json" \
  -d '{"title":"My Presentation","subtitle":"Created via API"}' \
  -o presentation.pptx

# Upload and extract
FILE_ID=$(curl -X POST "https://api.yourdomain.com/api/presentation/upload" \
  -F "file=@presentation.pptx" | jq -r '.fileId')

curl "https://api.yourdomain.com/api/presentation/$FILE_ID/slides/1"
```

:::

## Installation

::: code-group

```bash [Python]
pip install doclayer-py
```

```bash [TypeScript]
npm install @doclayer/ts
```

```bash [.NET]
dotnet add reference path/to/DocLayer.Core.csproj
```

:::

## Use Cases

- **Automated Report Generation** - Generate presentations from data
- **AI Agent Integration** - Use with LangChain, CrewAI, AutoGPT
- **Data Visualization** - Create charts and dashboards programmatically
- **Batch Processing** - Process hundreds of presentations at scale
- **Thumbnail Generation** - Create preview images for presentations
- **Content Extraction** - Extract text and data for analysis

## Why DocLayer?

- **No Microsoft Office Required** - Pure OpenXML implementation  
- **Cross-Platform** - Works everywhere .NET runs  
- **Multi-Language** - C#, Python, TypeScript, REST API  
- **Production Ready** - Battle-tested with comprehensive tests  
- **Open Source** - MIT License, free for commercial use  
- **Well Documented** - Examples, API docs, and guides
