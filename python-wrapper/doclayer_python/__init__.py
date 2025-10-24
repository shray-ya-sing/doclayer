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
