"""
DocLayer Remote MCP Server for ChatGPT

This server implements the Model Context Protocol (MCP) with SSE transport
for remote access by ChatGPT and other cloud-based AI agents.

Implements OpenAI's required search and fetch tools for deep research integration.
"""

import logging
import os
import json
import tempfile
import httpx
from typing import Dict, List, Any
from pathlib import Path

from fastmcp import FastMCP

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Import DocLayer
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "..", "python-wrapper"))
from doclayer_python import DocLayerClient, DocLayerError

# Initialize DocLayer client
doclayer_client = DocLayerClient()

# In-memory storage for uploaded presentations
# In production, use Redis, S3, or similar
presentation_store: Dict[str, str] = {}

server_instructions = """
DocLayer MCP Server provides PowerPoint generation and analysis capabilities.

Key features:
- Create presentations with custom themes
- Extract content from presentations (search)
- Retrieve full slide details (fetch)
- Render slides as images for visual inspection
- Edit slide content programmatically

Use search to find slides or presentations, then fetch to get full content.
"""

def create_server():
    """Create and configure the remote MCP server."""
    
    mcp = FastMCP(
        name="DocLayer MCP Server",
        instructions=server_instructions
    )
    
    @mcp.tool()
    async def search(query: str) -> Dict[str, List[Dict[str, Any]]]:
        """
        Search for presentations or slides matching the query.
        
        This tool searches through uploaded presentations and their content
        to find relevant matches. Returns basic information about matching
        slides. Use fetch tool to get complete slide content.
        
        Args:
            query: Search query string (e.g., "sales report", "Q4 data")
            
        Returns:
            Dictionary with 'results' key containing list of matching slides.
            Each result includes id, title, text snippet, and URL.
        """
        if not query or not query.strip():
            return {"results": []}
        
        logger.info(f"Searching for query: '{query}'")
        
        results = []
        query_lower = query.lower()
        
        # Search through all stored presentations
        for pres_id, pres_path in presentation_store.items():
            try:
                # Extract all slides from presentation
                all_slides = doclayer_client.extract_all_slides(pres_path)
                
                # Search through each slide
                for slide_num, slide_content in all_slides.items():
                    # Search through shapes text
                    slide_text = ""
                    for shape in slide_content.get('shapes', []):
                        shape_text = shape.get('text', '')
                        slide_text += shape_text + " "
                    
                    # Check if query matches
                    if query_lower in slide_text.lower():
                        # Create result
                        snippet = slide_text[:200] + "..." if len(slide_text) > 200 else slide_text
                        
                        result = {
                            "id": f"{pres_id}:slide:{slide_num}",
                            "title": f"Slide {slide_num}",
                            "text": snippet,
                            "url": f"presentation://{pres_id}/slide/{slide_num}"
                        }
                        results.append(result)
                        
            except Exception as e:
                logger.error(f"Error searching presentation {pres_id}: {e}")
                continue
        
        logger.info(f"Search returned {len(results)} results")
        return {"results": results}
    
    @mcp.tool()
    async def fetch(id: str) -> Dict[str, Any]:
        """
        Retrieve complete slide content by ID for detailed analysis.
        
        This tool fetches the full content of a specific slide including
        all shapes, text, and metadata. Use this after finding relevant
        slides with the search tool.
        
        Args:
            id: Slide identifier in format "pres_id:slide:slide_number"
            
        Returns:
            Complete slide document with id, title, full text, URL, and metadata.
            
        Raises:
            ValueError: If the specified ID is not found or invalid
        """
        if not id:
            raise ValueError("Slide ID is required")
        
        logger.info(f"Fetching content for ID: {id}")
        
        # Parse ID format: "pres_id:slide:slide_number"
        try:
            parts = id.split(":")
            if len(parts) != 3 or parts[1] != "slide":
                raise ValueError(f"Invalid ID format: {id}")
            
            pres_id = parts[0]
            slide_num = int(parts[2])
            
            if pres_id not in presentation_store:
                raise ValueError(f"Presentation not found: {pres_id}")
            
            pres_path = presentation_store[pres_id]
            
            # Extract slide content
            slide_content = doclayer_client.extract_slide_content(pres_path, slide_num)
            
            # Build full text from shapes
            full_text = ""
            for shape in slide_content.get('shapes', []):
                shape_name = shape.get('name', '')
                shape_text = shape.get('text', '')
                full_text += f"{shape_name}: {shape_text}\n"
            
            # Build metadata
            metadata = {
                "presentation_id": pres_id,
                "slide_number": slide_num,
                "shapes_count": len(slide_content.get('shapes', [])),
                "pictures_count": len(slide_content.get('pictures', []))
            }
            
            result = {
                "id": id,
                "title": f"Slide {slide_num}",
                "text": full_text,
                "url": f"presentation://{pres_id}/slide/{slide_num}",
                "metadata": metadata
            }
            
            logger.info(f"Fetched slide: {id}")
            return result
            
        except Exception as e:
            logger.error(f"Error fetching slide {id}: {e}")
            raise ValueError(f"Failed to fetch slide: {str(e)}")
    
    @mcp.tool()
    async def upload_presentation(file_url: str, presentation_id: str) -> Dict[str, str]:
        """
        Upload a PowerPoint presentation from a URL for analysis.
        
        Downloads a .pptx file from the provided URL and stores it for
        subsequent search and fetch operations.
        
        Args:
            file_url: Public URL to the .pptx file
            presentation_id: Unique identifier for this presentation
            
        Returns:
            Success message with presentation ID and slide count
        """
        logger.info(f"Uploading presentation from URL: {file_url}")
        
        try:
            # Download file from URL
            async with httpx.AsyncClient() as client:
                response = await client.get(file_url, follow_redirects=True)
                response.raise_for_status()
                
                # Save to temp file
                temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
                temp_file.write(response.content)
                temp_file.close()
                
                # Store in presentation store
                presentation_store[presentation_id] = temp_file.name
                
                # Get slide count
                slide_count = len(doclayer_client.extract_all_slides(temp_file.name))
                
                logger.info(f"Uploaded presentation {presentation_id} with {slide_count} slides")
                
                return {
                    "status": "success",
                    "presentation_id": presentation_id,
                    "slide_count": str(slide_count),
                    "message": f"Presentation uploaded successfully with {slide_count} slides"
                }
                
        except Exception as e:
            logger.error(f"Error uploading presentation: {e}")
            raise ValueError(f"Failed to upload presentation: {str(e)}")
    
    @mcp.tool()
    async def create_presentation(
        presentation_id: str,
        title: str,
        subtitle: str = None,
        font_name: str = None
    ) -> Dict[str, str]:
        """
        Create a new PowerPoint presentation.
        
        Args:
            presentation_id: Unique identifier for the new presentation
            title: Main title for the presentation
            subtitle: Optional subtitle
            font_name: Optional font name (e.g., "Arial", "Calibri")
            
        Returns:
            Success message with presentation ID
        """
        logger.info(f"Creating presentation: {presentation_id}")
        
        try:
            # Create temp file
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
            temp_file.close()
            
            # Create presentation
            if font_name:
                doclayer_client.create_presentation_with_theme(
                    temp_file.name,
                    title=title,
                    subtitle=subtitle,
                    font_name=font_name
                )
            else:
                doclayer_client.create_title_slide(
                    temp_file.name,
                    title=title,
                    subtitle=subtitle
                )
            
            # Store in presentation store
            presentation_store[presentation_id] = temp_file.name
            
            logger.info(f"Created presentation: {presentation_id}")
            
            return {
                "status": "success",
                "presentation_id": presentation_id,
                "message": f"Presentation '{title}' created successfully"
            }
            
        except Exception as e:
            logger.error(f"Error creating presentation: {e}")
            raise ValueError(f"Failed to create presentation: {str(e)}")
    
    @mcp.tool()
    async def edit_slide_text(
        presentation_id: str,
        slide_number: int,
        element_name: str,
        new_text: str
    ) -> Dict[str, str]:
        """
        Edit the text content of a shape on a slide.
        
        Args:
            presentation_id: Presentation identifier
            slide_number: Slide number (1-based index)
            element_name: Name of the shape element to edit
            new_text: New text content
            
        Returns:
            Success message
        """
        logger.info(f"Editing slide {slide_number} element '{element_name}' in {presentation_id}")
        
        try:
            if presentation_id not in presentation_store:
                raise ValueError(f"Presentation not found: {presentation_id}")
            
            pres_path = presentation_store[presentation_id]
            
            # Edit slide text
            doclayer_client.edit_slide_text(pres_path, slide_number, element_name, new_text)
            
            logger.info(f"Edited slide {slide_number} successfully")
            
            return {
                "status": "success",
                "message": f"Successfully edited element '{element_name}' on slide {slide_number}"
            }
            
        except Exception as e:
            logger.error(f"Error editing slide: {e}")
            raise ValueError(f"Failed to edit slide: {str(e)}")
    
    @mcp.tool()
    async def get_slide_count(presentation_id: str) -> Dict[str, Any]:
        """
        Get the total number of slides in a presentation.
        
        Args:
            presentation_id: Presentation identifier
            
        Returns:
            Slide count
        """
        logger.info(f"Getting slide count for {presentation_id}")
        
        try:
            if presentation_id not in presentation_store:
                raise ValueError(f"Presentation not found: {presentation_id}")
            
            pres_path = presentation_store[presentation_id]
            
            # Get slide count
            all_slides = doclayer_client.extract_all_slides(pres_path)
            count = len(all_slides)
            
            return {
                "presentation_id": presentation_id,
                "slide_count": count
            }
            
        except Exception as e:
            logger.error(f"Error getting slide count: {e}")
            raise ValueError(f"Failed to get slide count: {str(e)}")
    
    return mcp


def main():
    """Main function to start the remote MCP server."""
    logger.info("Starting DocLayer Remote MCP Server")
    logger.info("Server will be accessible via SSE transport on 0.0.0.0:8000")
    
    # Create the MCP server
    server = create_server()
    
    try:
        # Start server with SSE transport for remote access
        server.run(transport="sse", host="0.0.0.0", port=8000)
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server error: {e}")
        raise


if __name__ == "__main__":
    main()
