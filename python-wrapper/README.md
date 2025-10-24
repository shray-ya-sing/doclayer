# DocLayer Python Wrapper

Python bindings for the DocLayer.Core C# library, enabling PowerPoint generation from Python.

## Installation

```bash
pip install pythonnet>=3.0.0
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
