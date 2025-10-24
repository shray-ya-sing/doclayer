# doclayer-py

Python bindings for the DocLayer.Core C# library, enabling PowerPoint generation from Python.

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

## Testing

Run the test script:

```bash
python test_wrapper.py
```

## License

MIT License
