"""
DocLayer Python Client Library
Provides Python bindings for the C# DocLayer.Core library
"""

import os
import sys
from typing import Dict, List, Optional, Union
from pathlib import Path

try:
    from pathlib import Path as _Path
    from clr_loader import get_coreclr
    from pythonnet import set_runtime
    
    # Set up .NET runtime
    _bin_path = _Path(__file__).parent / "bin"
    _runtime = get_coreclr()
    set_runtime(_runtime)
    
    import clr
    # Add the bin directory to assembly search path
    import sys
    sys.path.append(str(_bin_path.absolute()))
except ImportError:
    raise ImportError(
        "pythonnet is required. Install with: pip install pythonnet"
    )


class DocLayerError(Exception):
    """Base exception for DocLayer operations"""
    pass


class DocLayerClient:
    """Python wrapper for C# DocLayer.Core library"""
    
    def __init__(self):
        # Load the C# assembly
        self._load_assembly()
        
    def _load_assembly(self):
        """Load the C# DocLayer.Core assembly"""
        try:
            # Add reference to DocLayer.Core DLL
            bin_path = Path(__file__).parent / "bin"
            dll_path = bin_path / "DocLayer.Core.dll"
            
            if not dll_path.exists():
                raise FileNotFoundError(f"DocLayer.Core.dll not found at {dll_path}")
            
            # Store bin path for assembly resolver
            self._bin_path = bin_path
            
            # Add bin directory to .NET assembly search path
            import System
            System.AppDomain.CurrentDomain.AssemblyResolve += self._assembly_resolver
            
            # Add all DLLs from bin directory
            import sys
            sys.path.append(str(bin_path.absolute()))
            
            # Add references to key assemblies with full paths
            clr.AddReference(str((bin_path / "DocumentFormat.OpenXml.dll").absolute()))
            clr.AddReference(str((bin_path / "DocLayer.Core.dll").absolute()))
            
            # Now import the .NET namespaces
            from DocumentFormat.OpenXml.Packaging import PresentationDocument
            from DocumentFormat.OpenXml.Presentation import Slide
            from OpenXMLExtensions import SlideExtensions, ShapeTreeExtensions, PresentationExtensions, PresentationHelperMethods
            from DocLayer.Core import PresentationBuilder, PresentationHelper
            from DocLayer.Core.Models import SlideContentInfo
            
            self.PresentationDocument = PresentationDocument
            self.Slide = Slide
            self.SlideExtensions = SlideExtensions
            self.ShapeTreeExtensions = ShapeTreeExtensions
            self.PresentationExtensions = PresentationExtensions
            self.PresentationHelperMethods = PresentationHelperMethods
            self.PresentationBuilder = PresentationBuilder
            self.PresentationHelper = PresentationHelper
            
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            raise DocLayerError(f"Failed to load C# assembly: {e}\n\nDetails:\n{error_details}")
    
    def _assembly_resolver(self, sender, args):
        """Resolve assembly dependencies from bin directory"""
        try:
            import System
            assembly_name = System.Reflection.AssemblyName(args.Name)
            dll_path = self._bin_path / f"{assembly_name.Name}.dll"
            if dll_path.exists():
                return System.Reflection.Assembly.LoadFrom(str(dll_path.absolute()))
        except:
            pass
        return None
    
    def create_title_slide(
        self, 
        filepath: str, 
        title: str, 
        subtitle: Optional[str] = None,
        footnote: Optional[str] = "Source:"
    ) -> bytes:
        """
        Create a PowerPoint presentation with a title slide
        
        Args:
            filepath: Path where the presentation will be saved
            title: Main title text
            subtitle: Subtitle text (optional)
            footnote: Footnote text (optional, defaults to "Source:")
            
        Returns:
            Bytes content of the created presentation file
        """
        try:
            # Create presentation using PresentationHelper
            presentation_doc = self.PresentationHelper.CreatePresentation(filepath, True)
            
            try:
                # Create PresentationBuilder
                builder = self.PresentationBuilder(presentation_doc)
                
                # Create title slide
                builder.CreateTitleSlide(title, subtitle, footnote)
                
                # Save and dispose
                presentation_doc.Save()
                presentation_doc.Dispose()
                
            except Exception as e:
                presentation_doc.Dispose()
                raise
                
            # Read and return file content
            with open(filepath, 'rb') as f:
                return f.read()
                
        except Exception as e:
            raise DocLayerError(f"Failed to create title slide: {e}")
    
    def create_presentation_with_theme(
        self,
        filepath: str,
        title: str,
        subtitle: Optional[str] = None,
        footnote: Optional[str] = "Source:",
        font_name: Optional[str] = None,
        accent_colors: Optional[List[str]] = None
    ) -> bytes:
        """
        Create a PowerPoint presentation with custom theme and title slide
        
        Args:
            filepath: Path where the presentation will be saved
            title: Main title text
            subtitle: Subtitle text (optional)
            footnote: Footnote text (optional, defaults to "Source:")
            font_name: Font typeface name (e.g., "Arial", "Calibri") - optional
            accent_colors: List of 4 hex color codes for accent colors - optional
            
        Returns:
            Bytes content of the created presentation file
            
        Example:
            >>> pptx_bytes = client.create_presentation_with_theme(
            ...     "presentation.pptx",
            ...     title="Custom Theme",
            ...     subtitle="With custom colors",
            ...     font_name="Arial",
            ...     accent_colors=["FF5733", "33FF57", "3357FF", "F3FF33"]
            ... )
        """
        try:
            # Create presentation using PresentationHelper
            presentation_doc = self.PresentationHelper.CreatePresentation(filepath, True)
            
            try:
                # Create PresentationBuilder
                builder = self.PresentationBuilder(presentation_doc)
                
                # Set theme if any theme parameters provided
                if font_name or accent_colors:
                    # Convert Python list to .NET List for accent colors
                    net_colors = None
                    if accent_colors:
                        if len(accent_colors) != 4:
                            raise ValueError("Must provide exactly 4 accent colors")
                        import System.Collections.Generic as Generic
                        net_colors = Generic.List[str]()
                        for color in accent_colors:
                            net_colors.Add(color)
                    
                    builder.SetPresentationTheme(font_name, net_colors)
                
                # Create title slide
                builder.CreateTitleSlide(title, subtitle, footnote)
                
                # Save and dispose
                presentation_doc.Save()
                presentation_doc.Dispose()
                
            except Exception as e:
                presentation_doc.Dispose()
                raise
                
            # Read and return file content
            with open(filepath, 'rb') as f:
                return f.read()
                
        except Exception as e:
            raise DocLayerError(f"Failed to create presentation with theme: {e}")
    
    def get_slide_count(self, filepath: str) -> int:
        """
        Get the number of slides in a presentation
        
        Args:
            filepath: Path to the presentation file
            
        Returns:
            Number of slides
        """
        try:
            builder = self.PresentationBuilder.FromFile(filepath, False)
            try:
                return builder.GetSlideCount()
            finally:
                builder.Dispose()
        except Exception as e:
            raise DocLayerError(f"Failed to get slide count: {e}")
    
    def extract_slide_content(self, filepath: str, slide_number: int) -> Dict:
        """
        Extract content from a specific slide
        
        Args:
            filepath: Path to the presentation file
            slide_number: Slide number (1-based index)
            
        Returns:
            Dictionary with 'shapes' and 'pictures' lists containing element info
        """
        try:
            builder = self.PresentationBuilder.FromFile(filepath, False)
            try:
                content = builder.ExtractSlideContent(slide_number)
                
                # Convert to Python dict
                result = {
                    'shapes': [],
                    'pictures': []
                }
                
                for shape in content.Shapes:
                    shape_dict = {
                        'name': shape.Name,
                        'text': shape.Text
                    }
                    if shape.Position:
                        shape_dict['position'] = {'x': shape.Position.X, 'y': shape.Position.Y}
                    if shape.Size:
                        shape_dict['size'] = {'width': shape.Size.Width, 'height': shape.Size.Height}
                    result['shapes'].append(shape_dict)
                
                for picture in content.Pictures:
                    pic_dict = {'name': picture.Name}
                    if picture.Position:
                        pic_dict['position'] = {'x': picture.Position.X, 'y': picture.Position.Y}
                    if picture.Size:
                        pic_dict['size'] = {'width': picture.Size.Width, 'height': picture.Size.Height}
                    result['pictures'].append(pic_dict)
                
                return result
            finally:
                builder.Dispose()
        except Exception as e:
            raise DocLayerError(f"Failed to extract slide content: {e}")
    
    def extract_all_slides(self, filepath: str) -> Dict[int, Dict]:
        """
        Extract content from all slides in a presentation
        
        Args:
            filepath: Path to the presentation file
            
        Returns:
            Dictionary mapping slide numbers to their content
        """
        try:
            builder = self.PresentationBuilder.FromFile(filepath, False)
            try:
                all_content = builder.ExtractAllSlides()
                
                result = {}
                for slide_num in all_content.Keys:
                    content = all_content[slide_num]
                    
                    slide_dict = {
                        'shapes': [],
                        'pictures': []
                    }
                    
                    for shape in content.Shapes:
                        shape_dict = {
                            'name': shape.Name,
                            'text': shape.Text
                        }
                        if shape.Position:
                            shape_dict['position'] = {'x': shape.Position.X, 'y': shape.Position.Y}
                        if shape.Size:
                            shape_dict['size'] = {'width': shape.Size.Width, 'height': shape.Size.Height}
                        slide_dict['shapes'].append(shape_dict)
                    
                    for picture in content.Pictures:
                        pic_dict = {'name': picture.Name}
                        if picture.Position:
                            pic_dict['position'] = {'x': picture.Position.X, 'y': picture.Position.Y}
                        if picture.Size:
                            pic_dict['size'] = {'width': picture.Size.Width, 'height': picture.Size.Height}
                        slide_dict['pictures'].append(pic_dict)
                    
                    result[slide_num] = slide_dict
                
                return result
            finally:
                builder.Dispose()
        except Exception as e:
            raise DocLayerError(f"Failed to extract all slides: {e}")
    
    def edit_slide_text(self, filepath: str, slide_number: int, element_name: str, new_text: str) -> None:
        """
        Edit the text of a shape on a slide
        
        Args:
            filepath: Path to the presentation file
            slide_number: Slide number (1-based index)
            element_name: Name of the shape to edit
            new_text: New text content
        """
        try:
            builder = self.PresentationBuilder.FromFile(filepath, True)
            try:
                builder.EditSlideText(slide_number, element_name, new_text)
                builder.Save()
            finally:
                builder.Dispose()
        except Exception as e:
            raise DocLayerError(f"Failed to edit slide text: {e}")
    
    def render_slide_to_image(self, filepath: str, slide_number: int) -> str:
        """
        Render a slide to a JPEG image
        
        Args:
            filepath: Path to the presentation file
            slide_number: Slide number (1-based index)
            
        Returns:
            Path to the generated image file
        """
        try:
            # Note: Document must be closed before rendering
            from InternalUtilities.Syncfusion import SyncfusionHelperMethods
            return SyncfusionHelperMethods.ExportSlideToImage(filepath, slide_number)
        except Exception as e:
            raise DocLayerError(f"Failed to render slide to image: {e}")
    
    def render_all_slides(self, filepath: str) -> List[str]:
        """
        Render all slides to JPEG images
        
        Args:
            filepath: Path to the presentation file
            
        Returns:
            List of paths to generated image files
        """
        try:
            from InternalUtilities.Syncfusion import SyncfusionHelperMethods
            import System.Collections.Generic as Generic
            
            images = SyncfusionHelperMethods.ExportPptToImages(filepath)
            return [str(img) for img in images]
        except Exception as e:
            raise DocLayerError(f"Failed to render all slides: {e}")


# Convenience functions
def create_title_slide(
    filepath: str,
    title: str,
    subtitle: Optional[str] = None,
    footnote: Optional[str] = "Source:"
) -> bytes:
    """
    Convenience function to create a title slide presentation
    
    Args:
        filepath: Path where the presentation will be saved
        title: Main title text
        subtitle: Subtitle text (optional)
        footnote: Footnote text (optional, defaults to "Source:")
        
    Returns:
        Bytes content of the created presentation file
        
    Example:
        >>> from doclayer_python import create_title_slide
        >>> pptx_bytes = create_title_slide(
        ...     "presentation.pptx",
        ...     title="Welcome to DocLayer",
        ...     subtitle="PowerPoint Generation Made Easy",
        ...     footnote="Source: DocLayer.Core"
        ... )
    """
    client = DocLayerClient()
    return client.create_title_slide(filepath, title, subtitle, footnote)


def create_presentation_with_theme(
    filepath: str,
    title: str,
    subtitle: Optional[str] = None,
    footnote: Optional[str] = "Source:",
    font_name: Optional[str] = None,
    accent_colors: Optional[List[str]] = None
) -> bytes:
    """
    Convenience function to create a presentation with custom theme
    
    Args:
        filepath: Path where the presentation will be saved
        title: Main title text
        subtitle: Subtitle text (optional)
        footnote: Footnote text (optional, defaults to "Source:")
        font_name: Font typeface name (e.g., "Arial", "Calibri") - optional
        accent_colors: List of 4 hex color codes for accent colors - optional
        
    Returns:
        Bytes content of the created presentation file
        
    Example:
        >>> from doclayer_python import create_presentation_with_theme
        >>> pptx_bytes = create_presentation_with_theme(
        ...     "custom.pptx",
        ...     title="Custom Theme Demo",
        ...     subtitle="With Arial and custom colors",
        ...     font_name="Arial",
        ...     accent_colors=["FF5733", "33FF57", "3357FF", "F3FF33"]
        ... )
    """
    client = DocLayerClient()
    return client.create_presentation_with_theme(
        filepath, title, subtitle, footnote, font_name, accent_colors
    )


# Export public API
__all__ = [
    'DocLayerClient',
    'create_title_slide',
    'create_presentation_with_theme',
    'DocLayerError'
]
