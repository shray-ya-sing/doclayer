# DocLayer MCP Server

An MCP (Model Context Protocol) server that exposes DocLayer PowerPoint generation capabilities to AI agents like Claude Desktop.

## Features

- **Create Presentations**: Generate PowerPoint files with custom themes, fonts, and colors
- **Extract Content**: Extract text, shapes, and metadata from existing presentations
- **Visual Inspection**: Render slides as images with base64 encoding for AI agent viewing
- **Edit Slides**: Modify text content in presentations programmatically
- **Full Integration**: Works seamlessly with Claude Desktop and other MCP clients

## Installation

### Prerequisites

1. **Python 3.8+** - Required for running the MCP server
2. **doclayer-py** - DocLayer Python wrapper (installed automatically)
3. **.NET 8.0 Runtime** - Required by doclayer-py

### Install from source

```bash
cd mcp-server
pip install -e .
```

### Install from PyPI (when published)

```bash
pip install doclayer-mcp-server
```

## Configuration for AI Agents

The MCP server works with any MCP-compatible AI agent (Claude Desktop, Claude Code, Zed, etc.).

### For Claude Desktop / Claude Code

Add the following to your MCP configuration file:

**Claude Desktop**
- **macOS/Linux**: `~/.config/Claude/claude_desktop_config.json`  
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

**Claude Code (VS Code Extension)**
- Configure via the extension settings or MCP config panel

**Configuration:**

```json
{
  "mcpServers": {
    "doclayer": {
      "command": "python",
      "args": ["-m", "doclayer_mcp.server"],
      "env": {},
      "working_directory": null
    }
  }
}
```

**If using a virtual environment**, use the full path to Python:

```json
{
  "mcpServers": {
    "doclayer": {
      "command": "/path/to/your/venv/bin/python",
      "args": ["-m", "doclayer_mcp.server"],
      "env": {},
      "working_directory": null
    }
  }
}
```

**Windows example with virtual environment:**
```json
{
  "mcpServers": {
    "doclayer": {
      "command": "C:\\path\\to\\venv\\Scripts\\python.exe",
      "args": ["-m", "doclayer_mcp.server"],
      "env": {},
      "working_directory": null
    }
  }
}
```

## Restart Your AI Agent

After adding the configuration, restart your AI agent application (Claude Desktop, VS Code, etc.) for the changes to take effect.

## Available Tools

The MCP server exposes the following tools to AI agents:

### `create_title_slide`
Create a PowerPoint presentation with a title slide.

**Parameters:**
- `filepath` (string, required): Path where the presentation will be saved
- `title` (string, required): Main title text
- `subtitle` (string, optional): Subtitle text
- `footnote` (string, optional): Footnote text (defaults to "Source:")

### `create_presentation_with_theme`
Create a presentation with custom theme (fonts and colors).

**Parameters:**
- `filepath` (string, required): Path where the presentation will be saved
- `title` (string, required): Main title text
- `subtitle` (string, optional): Subtitle text
- `footnote` (string, optional): Footnote text
- `font_name` (string, optional): Font typeface name (e.g., "Arial", "Calibri")
- `accent_colors` (array, optional): List of 4 hex color codes (without #)

### `extract_slide_content`
Extract text content, shapes, and pictures from a specific slide.

**Parameters:**
- `filepath` (string, required): Path to the presentation file
- `slide_number` (integer, required): Slide number (1-based index)

### `extract_all_slides`
Extract content from all slides in a presentation.

**Parameters:**
- `filepath` (string, required): Path to the presentation file

### `render_slide_image`
Render a specific slide as an image (base64-encoded for AI viewing).

**Parameters:**
- `filepath` (string, required): Path to the presentation file
- `slide_number` (integer, required): Slide number (1-based index)

### `render_all_slides_images`
Render all slides as images (base64-encoded for AI viewing). **Limited to first 5 slides** to avoid overwhelming the AI agent's context window.

**Parameters:**
- `filepath` (string, required): Path to the presentation file

**Note**: For presentations with more than 5 slides, use `render_slide_image` to view specific slides.

### `edit_slide_text`
Edit the text content of a shape on a slide.

**Parameters:**
- `filepath` (string, required): Path to the presentation file
- `slide_number` (integer, required): Slide number (1-based index)
- `element_name` (string, required): Name of the shape element to edit
- `new_text` (string, required): New text content

### `get_slide_count`
Get the total number of slides in a presentation.

**Parameters:**
- `filepath` (string, required): Path to the presentation file

## Configuration: Image Return Format

The MCP server can return slide images in different formats. Edit `src/doclayer_mcp/server.py` line 37:

```python
IMAGE_FORMAT = "data_uri"  # Options: "data_uri", "base64_only", "file_path"
```

- `"data_uri"` (default): Returns `data:image/jpeg;base64,...` format
- `"base64_only"`: Returns only the base64 string
- `"file_path"`: Returns the file path to the rendered image (no encoding)

Different AI agents may work better with different formats. Experiment to find what works best.

## Example Usage with AI Agents

Once configured, you can interact with your AI agent like this:

> "Create a presentation called demo.pptx with the title 'Q4 Sales Report' and subtitle '2024 Performance Summary'"

> "Extract content from presentation.pptx and show me what's on each slide"

> "Render slide 1 from demo.pptx so I can see it"

> "Edit the title on slide 1 to 'Updated Q4 Report'"

The AI agent will use the MCP tools to create presentations, extract content, display slide images, and make edits.

## Demo Video Use Case

This MCP server is perfect for filming demo videos showing:
1. AI agent creating presentations from natural language
2. Agent analyzing existing presentations (both structure and visuals)
3. Agent matching extracted text to visual elements on slides
4. Agent editing presentations based on feedback

## Development

### Run in development mode

```bash
cd mcp-server
pip install -e ".[dev]"
```

### Run tests

```bash
pytest
```

### Format code

```bash
black src/
```

## Troubleshooting

### Server not showing up in AI agent

1. Check that the config file path is correct for your OS and application
2. Verify the Python path is correct (especially if using a virtual environment)
3. Restart your AI agent application completely
4. Check application logs for errors
5. Test that the server starts manually: `python -m doclayer_mcp.server`

### Import errors

Make sure doclayer-py is installed:
```bash
pip install doclayer-py
```

### .NET Runtime errors

Ensure .NET 8.0 Runtime is installed:
- **Windows**: Download from [Microsoft](https://dotnet.microsoft.com/download/dotnet/8.0)
- **macOS**: `brew install dotnet@8`
- **Linux**: Follow [Microsoft's guide](https://learn.microsoft.com/dotnet/core/install/linux)

## License

MIT License - See LICENSE file for details

## Links

- [DocLayer Documentation](https://docs.doclayer.dev)
- [MCP Protocol](https://modelcontextprotocol.io)
- [Claude Desktop](https://claude.ai/download)
