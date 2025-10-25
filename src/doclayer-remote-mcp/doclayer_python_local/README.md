# doclayer-py

Python bindings for the DocLayer.Core C# library, enabling PowerPoint generation, extraction, editing, and rendering from Python.

## Installation

```bash
pip install doclayer-py
```

Or install from source:

```bash
cd python-wrapper
pip install pythonnet>=3.0.0
pip install -e .
```

## Usage

### Simple Title Slide

```python
from doclayer_python import create_title_slide

# Create a presentation with a title slide
pptx_bytes = create_title_slide(
    filepath="presentation.pptx",
    title="Welcome to DocLayer",
    subtitle="PowerPoint Generation Made Easy",
    footnote="Source: DocLayer.Core"
)

print(f"Created presentation: {len(pptx_bytes)} bytes")
```

### Presentation with Custom Theme

```python
from doclayer_python import create_presentation_with_theme

# Create a presentation with custom font and colors
pptx_bytes = create_presentation_with_theme(
    filepath="custom_theme.pptx",
    title="Custom Theme Demo",
    subtitle="Arial font with brand colors",
    footnote="Source: My Company",
    font_name="Arial",
    accent_colors=[
        "FF5733",  # Red-Orange (Accent 1)
        "33FF57",  # Green (Accent 2) 
        "3357FF",  # Blue (Accent 3)
        "F3FF33"   # Yellow (Accent 4)
    ]
)

print(f"Created themed presentation: {len(pptx_bytes)} bytes")
```

### Using the Client

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

# Or create with custom theme
pptx_bytes = client.create_presentation_with_theme(
    filepath="themed.pptx",
    title="Themed Presentation",
    font_name="Calibri",
    accent_colors=["4472C4", "ED7D31", "A5A5A5", "FFC000"]
)
```

### Extract and Edit Presentations

```python
from doclayer_python import DocLayerClient

client = DocLayerClient()

# Get slide count
count = client.get_slide_count("presentation.pptx")
print(f"Presentation has {count} slides")

# Extract content from a specific slide
content = client.extract_slide_content("presentation.pptx", slide_number=1)
for shape in content['shapes']:
    print(f"Shape: {shape['name']}")
    print(f"Text: {shape['text']}")
    print(f"Position: {shape.get('position')}")
    print(f"Size: {shape.get('size')}")

# Extract all slides
all_content = client.extract_all_slides("presentation.pptx")
for slide_num, slide_content in all_content.items():
    print(f"Slide {slide_num}: {len(slide_content['shapes'])} shapes, {len(slide_content['pictures'])} pictures")

# Edit text on a slide
client.edit_slide_text(
    filepath="presentation.pptx",
    slide_number=1,
    element_name="Title 1",
    new_text="Updated Title Text"
)

# Render slide to JPEG image
image_path = client.render_slide_to_image("presentation.pptx", slide_number=1)
print(f"Slide rendered to: {image_path}")

# Render all slides to images
image_paths = client.render_all_slides("presentation.pptx")
for i, path in enumerate(image_paths, 1):
    print(f"Slide {i} rendered to: {path}")
```

## Requirements

- Python 3.8+
- pythonnet 3.0.0+
- .NET 8.0 Runtime
- Windows (for .NET Framework) or cross-platform (for .NET Core)

## Architecture

The Python wrapper uses `pythonnet` to call C# methods from the `DocLayer.Core` library:

```
Python → pythonnet → DocLayer.Core.dll → OpenXML SDK → PowerPoint Files
```

## API Reference

### DocLayerClient

**Creation Methods:**
- `create_title_slide(filepath, title, subtitle=None, footnote="Source:")` - Create presentation with title slide
- `create_presentation_with_theme(filepath, title, subtitle=None, footnote="Source:", font_name=None, accent_colors=None)` - Create with custom theme

**Query Methods:**
- `get_slide_count(filepath)` - Get number of slides in presentation
- `extract_slide_content(filepath, slide_number)` - Extract shapes and pictures from a slide
- `extract_all_slides(filepath)` - Extract content from all slides

**Edit Methods:**
- `edit_slide_text(filepath, slide_number, element_name, new_text)` - Modify text of a shape

**Rendering Methods:**
- `render_slide_to_image(filepath, slide_number)` - Render slide to JPEG, returns image path
- `render_all_slides(filepath)` - Render all slides to JPEG images, returns list of paths

## Testing

Run the test scripts:

```bash
# Test basic creation
python test_wrapper.py

# Test extraction, editing, and rendering
python test_new_methods.py
```

## License

MIT License
