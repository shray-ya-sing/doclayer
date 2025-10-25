#!/usr/bin/env python3
"""
DocLayer MCP Server

Exposes DocLayer PowerPoint generation capabilities as MCP tools for AI agents.
Includes base64-encoded image rendering for visual slide inspection.
"""

import json
import base64
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from mcp.server import Server
from mcp.types import (
    Resource,
    Tool,
    TextContent,
    ImageContent,
    EmbeddedResource,
)
import mcp.server.stdio

# Import doclayer_python - will be installed as dependency
from doclayer_python import DocLayerClient, DocLayerError


app = Server("doclayer-mcp-server")
doclayer_client = DocLayerClient()

# Configuration: Image return format
# Options:
#   "data_uri" - Return as data:image/jpeg;base64,... (default)
#   "base64_only" - Return only the base64 string without data URI prefix
#   "file_path" - Return the file path to the rendered image (no base64 encoding)
IMAGE_FORMAT = "data_uri"  # Change this to customize image return format


# Helper function to convert image file to base64 data URI
def image_to_base64_data_uri(image_path: str) -> str:
    """Convert an image file to a base64 data URI for MCP."""
    with open(image_path, 'rb') as f:
        image_data = base64.b64encode(f.read()).decode('utf-8')
    
    # Determine MIME type from extension
    ext = Path(image_path).suffix.lower()
    mime_map = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
    }
    mime_type = mime_map.get(ext, 'image/jpeg')
    
    if IMAGE_FORMAT == "base64_only":
        return image_data
    elif IMAGE_FORMAT == "file_path":
        return image_path
    else:  # "data_uri" (default)
        return f"data:{mime_type};base64,{image_data}"


@app.list_tools()
async def list_tools() -> List[Tool]:
    """List all available DocLayer MCP tools."""
    return [
        Tool(
            name="create_title_slide",
            description="Create a PowerPoint presentation with a title slide. Returns the file path.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path where the presentation will be saved (e.g., 'output.pptx')"
                    },
                    "title": {
                        "type": "string",
                        "description": "Main title text for the slide"
                    },
                    "subtitle": {
                        "type": "string",
                        "description": "Subtitle text (optional)"
                    },
                    "footnote": {
                        "type": "string",
                        "description": "Footnote text (optional, defaults to 'Source:')"
                    }
                },
                "required": ["filepath", "title"]
            }
        ),
        Tool(
            name="create_presentation_with_theme",
            description="Create a PowerPoint presentation with custom theme (fonts and colors). Returns the file path.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path where the presentation will be saved"
                    },
                    "title": {
                        "type": "string",
                        "description": "Main title text"
                    },
                    "subtitle": {
                        "type": "string",
                        "description": "Subtitle text (optional)"
                    },
                    "footnote": {
                        "type": "string",
                        "description": "Footnote text (optional)"
                    },
                    "font_name": {
                        "type": "string",
                        "description": "Font typeface name (e.g., 'Arial', 'Calibri')"
                    },
                    "accent_colors": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of 4 hex color codes (without #) for accent colors"
                    }
                },
                "required": ["filepath", "title"]
            }
        ),
        Tool(
            name="extract_slide_content",
            description="Extract text content, shapes, and pictures from a specific slide in a presentation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    },
                    "slide_number": {
                        "type": "integer",
                        "description": "Slide number to extract (1-based index)"
                    }
                },
                "required": ["filepath", "slide_number"]
            }
        ),
        Tool(
            name="extract_all_slides",
            description="Extract content from all slides in a presentation. Returns structured data about all slides.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    }
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="render_slide_image",
            description="Render a specific slide as an image and return it as base64-encoded data URI for visual inspection.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    },
                    "slide_number": {
                        "type": "integer",
                        "description": "Slide number to render (1-based index)"
                    }
                },
                "required": ["filepath", "slide_number"]
            }
        ),
        Tool(
            name="render_all_slides_images",
            description="Render all slides as images and return them as base64-encoded data URIs for visual inspection. Limited to first 5 slides to avoid context overflow.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    }
                },
                "required": ["filepath"]
            }
        ),
        Tool(
            name="edit_slide_text",
            description="Edit the text content of a shape on a slide by its element name.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    },
                    "slide_number": {
                        "type": "integer",
                        "description": "Slide number (1-based index)"
                    },
                    "element_name": {
                        "type": "string",
                        "description": "Name of the shape element to edit"
                    },
                    "new_text": {
                        "type": "string",
                        "description": "New text content for the element"
                    }
                },
                "required": ["filepath", "slide_number", "element_name", "new_text"]
            }
        ),
        Tool(
            name="get_slide_count",
            description="Get the total number of slides in a presentation.",
            inputSchema={
                "type": "object",
                "properties": {
                    "filepath": {
                        "type": "string",
                        "description": "Path to the presentation file"
                    }
                },
                "required": ["filepath"]
            }
        ),
    ]


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> List[TextContent | ImageContent]:
    """Handle tool calls from MCP clients."""
    
    try:
        if name == "create_title_slide":
            filepath = str(arguments["filepath"])
            title = str(arguments["title"])
            subtitle = str(arguments["subtitle"]) if arguments.get("subtitle") else None
            footnote = str(arguments.get("footnote", "Source:"))
            
            doclayer_client.create_title_slide(filepath, title, subtitle, footnote)
            
            return [TextContent(
                type="text",
                text=f"Successfully created presentation at: {os.path.abspath(filepath)}"
            )]
        
        elif name == "create_presentation_with_theme":
            filepath = str(arguments["filepath"])
            title = str(arguments["title"])
            subtitle = str(arguments["subtitle"]) if arguments.get("subtitle") else None
            footnote = str(arguments.get("footnote", "Source:"))
            font_name = str(arguments["font_name"]) if arguments.get("font_name") else None
            accent_colors = arguments.get("accent_colors")
            
            doclayer_client.create_presentation_with_theme(
                filepath, title, subtitle, footnote, font_name, accent_colors
            )
            
            return [TextContent(
                type="text",
                text=f"Successfully created themed presentation at: {os.path.abspath(filepath)}"
            )]
        
        elif name == "extract_slide_content":
            filepath = str(arguments["filepath"])
            slide_number = int(arguments["slide_number"])
            
            content = doclayer_client.extract_slide_content(filepath, slide_number)
            
            return [TextContent(
                type="text",
                text=f"Slide {slide_number} content:\n{json.dumps(content, indent=2)}"
            )]
        
        elif name == "extract_all_slides":
            filepath = str(arguments["filepath"])
            
            all_content = doclayer_client.extract_all_slides(filepath)
            
            return [TextContent(
                type="text",
                text=f"All slides content:\n{json.dumps(all_content, indent=2)}"
            )]
        
        elif name == "render_slide_image":
            filepath = str(arguments["filepath"])
            slide_number = int(arguments["slide_number"])
            
            # Render slide to image file
            image_path = doclayer_client.render_slide_to_image(filepath, slide_number)
            
            # Convert to base64 data URI
            data_uri = image_to_base64_data_uri(image_path)
            
            return [
                TextContent(
                    type="text",
                    text=f"Rendered slide {slide_number} from {filepath}"
                ),
                ImageContent(
                    type="image",
                    data=data_uri,
                    mimeType="image/jpeg"
                )
            ]
        
        elif name == "render_all_slides_images":
            filepath = str(arguments["filepath"])
            
            # Render all slides to image files
            image_paths = doclayer_client.render_all_slides(filepath)
            
            # Limit to first 5 slides to avoid overwhelming the LLM context
            MAX_SLIDES = 5
            if len(image_paths) > MAX_SLIDES:
                limited_paths = image_paths[:MAX_SLIDES]
                results = [TextContent(
                    type="text",
                    text=f"Presentation has {len(image_paths)} slides. Showing first {MAX_SLIDES} slides. Use render_slide_image for specific slides."
                )]
            else:
                limited_paths = image_paths
                results = [TextContent(
                    type="text",
                    text=f"Rendered {len(image_paths)} slides from {filepath}"
                )]
            
            # Convert to base64 and return as multiple image contents
            for idx, image_path in enumerate(limited_paths, start=1):
                data_uri = image_to_base64_data_uri(image_path)
                results.append(ImageContent(
                    type="image",
                    data=data_uri,
                    mimeType="image/jpeg"
                ))
            
            return results
        
        elif name == "edit_slide_text":
            filepath = str(arguments["filepath"])
            slide_number = int(arguments["slide_number"])
            element_name = str(arguments["element_name"])
            new_text = str(arguments["new_text"])
            
            doclayer_client.edit_slide_text(filepath, slide_number, element_name, new_text)
            
            return [TextContent(
                type="text",
                text=f"Successfully edited element '{element_name}' on slide {slide_number}"
            )]
        
        elif name == "get_slide_count":
            filepath = str(arguments["filepath"])
            
            # Use extract_all_slides to count (alternative approach)
            all_slides = doclayer_client.extract_all_slides(filepath)
            count = len(all_slides)
            
            return [TextContent(
                type="text",
                text=f"Presentation has {count} slide(s)"
            )]
        
        else:
            return [TextContent(
                type="text",
                text=f"Unknown tool: {name}"
            )]
    
    except DocLayerError as e:
        return [TextContent(
            type="text",
            text=f"DocLayer error: {str(e)}"
        )]
    except Exception as e:
        return [TextContent(
            type="text",
            text=f"Error executing tool '{name}': {str(e)}"
        )]


async def main():
    """Run the MCP server."""
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await app.run(
            read_stream,
            write_stream,
            app.create_initialization_options()
        )


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
