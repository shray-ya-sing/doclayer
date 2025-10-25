# Setup Guide for Demo Video

This guide will help you set up the DocLayer MCP server for filming a demo video with Claude Desktop.

## Prerequisites

1. ✅ **Python 3.8+** installed
2. ✅ **.NET 8.0 Runtime** installed
3. ✅ **Claude Desktop** installed ([download here](https://claude.ai/download))
4. ✅ **doclayer-py** Python package installed

## Step 1: Install the MCP Server

From the `mcp-server` directory:

```bash
cd mcp-server
pip install -e .
```

This will install:
- The MCP server package
- Required dependency: `mcp>=0.9.0`
- Required dependency: `doclayer-py>=0.1.0` (if not already installed)

## Step 2: Find Your Claude Desktop Config File

The config file location depends on your OS:

**Windows:**
```
%APPDATA%\Claude\claude_desktop_config.json
```
(Usually: `C:\Users\YourUsername\AppData\Roaming\Claude\claude_desktop_config.json`)

**macOS:**
```
~/.config/Claude/claude_desktop_config.json
```

**Linux:**
```
~/.config/Claude/claude_desktop_config.json
```

## Step 3: Configure Claude Desktop

1. Open the config file (create it if it doesn't exist)
2. Add the DocLayer MCP server configuration:

```json
{
  "mcpServers": {
    "doclayer": {
      "command": "python",
      "args": ["-m", "doclayer_mcp.server"]
    }
  }
}
```

**Important for Windows users:** If using a virtual environment, use the full path:
```json
{
  "mcpServers": {
    "doclayer": {
      "command": "C:\\path\\to\\venv\\Scripts\\python.exe",
      "args": ["-m", "doclayer_mcp.server"]
    }
  }
}
```

## Step 4: Restart Claude Desktop

Completely quit and restart Claude Desktop for the changes to take effect.

## Step 5: Verify the Server is Connected

In Claude Desktop, you should see a tool/plug icon indicating MCP servers are connected. You can ask Claude:

> "What MCP tools do you have access to?"

Claude should list the DocLayer tools:
- `create_title_slide`
- `create_presentation_with_theme`
- `extract_slide_content`
- `extract_all_slides`
- `render_slide_image`
- `render_all_slides_images`
- `edit_slide_text`
- `get_slide_count`

## Demo Video Script Ideas

### Scenario 1: Create and Visualize
```
You: "Create a presentation called demo.pptx with the title 'AI-Powered Presentations' and subtitle 'Made with DocLayer'"

You: "Now show me what it looks like"
```

Claude will:
1. Call `create_title_slide` to create the presentation
2. Call `render_all_slides_images` to render it
3. Display the slide images directly in the conversation

### Scenario 2: Analyze Existing Presentation
```
You: "I have a presentation at C:/presentations/report.pptx. Extract all the content and show me the slides."
```

Claude will:
1. Call `extract_all_slides` to get the text/structure
2. Call `render_all_slides_images` to get visuals
3. Describe the content and show you the images

### Scenario 3: Match Content to Visuals
```
You: "Look at the presentation and tell me which slide has the 'Q4 Revenue' chart"
```

Claude will:
1. Extract text content from all slides
2. Render all slides as images
3. Use vision to identify which slide contains the chart
4. Match the visual to the extracted metadata

### Scenario 4: Edit and Verify
```
You: "Change the title on slide 1 to 'Updated Report' and show me the result"
```

Claude will:
1. Call `edit_slide_text` to modify the presentation
2. Call `render_slide_image` for slide 1 to show the updated version

## Troubleshooting

### "Server not found" or tools not showing up

1. Check that the config file path is correct
2. Verify the Python path (try `which python` or `where python`)
3. Make sure you completely restarted Claude Desktop
4. Check Claude Desktop logs for errors

### Import errors when running the server

Install doclayer-py manually:
```bash
pip install doclayer-py
```

Verify it works:
```bash
python -c "from doclayer_python import DocLayerClient; print('Success!')"
```

### .NET Runtime errors

Make sure .NET 8.0 Runtime is installed:
```bash
dotnet --list-runtimes
```

You should see `Microsoft.NETCore.App 8.0.x` in the list.

## Tips for Filming

1. **Prepare sample presentations**: Have a few .pptx files ready to demonstrate extraction and editing
2. **Use full paths**: Easier for Claude to work with absolute file paths
3. **Show the rendered images**: The visual inspection capability is the key feature
4. **Demonstrate round-trip**: Create → View → Edit → View again
5. **Highlight AI understanding**: Show Claude matching text content to visual elements

## File Locations

During the demo, slides will be rendered to temporary image files. By default, these are saved in the same directory as your presentation with names like:
- `presentation_slide_1.jpeg`
- `presentation_slide_2.jpeg`

You may want to clean these up between takes or use a dedicated demo folder.

## Next Steps

Once you've verified everything works, you're ready to film! The MCP server will handle all the heavy lifting, and Claude will demonstrate intelligent use of your library's capabilities.

Good luck with your demo! 🎥
