"""
DocLayer Python Client Library
Provides Python bindings for the C# OpenXMLExtensions library
"""

import os
import sys
from typing import Dict, List, Optional, Union
from pathlib import Path

try:
    import clr
    import pythonnet
    pythonnet.load("coreclr")
except ImportError:
    raise ImportError(
        "pythonnet is required. Install with: pip install pythonnet"
    )

class DocLayerError(Exception):
    """Base exception for DocLayer operations"""
    pass

class SlideBuilder:
    """Python wrapper for C# SlideExtensions"""
    
    def __init__(self):
        # Load the C# assembly
        self._load_assembly()
        
    def _load_assembly(self):
        """Load the C# OOXML extensions assembly"""
        try:
            # Add reference to your compiled DLL
            dll_path = Path(__file__).parent / "bin" / "doclayer_webapi.dll"
            clr.AddReference(str(dll_path))
            
            # Import .NET types
            from DocumentFormat.OpenXml.Packaging import PresentationDocument
            from DocumentFormat.OpenXml.Presentation import Slide
            from OpenXMLExtensions import SlideExtensions, ShapeTreeExtensions, PresentationExtensions
            
            self.PresentationDocument = PresentationDocument
            self.Slide = Slide
            self.SlideExtensions = SlideExtensions
            self.ShapeTreeExtensions = ShapeTreeExtensions
            self.PresentationExtensions = PresentationExtensions
            
        except Exception as e:
            raise DocLayerError(f"Failed to load C# assembly: {e}")
    
    def create_presentation(self, filepath: str) -> 'Presentation':
        """Create a new presentation"""
        return Presentation(filepath, self)

class Presentation:
    """Python wrapper for PowerPoint presentation"""
    
    def __init__(self, filepath: str, builder: SlideBuilder):
        self.filepath = filepath
        self.builder = builder
        self._slides = []
        self._presentation_doc = None
        
    def __enter__(self):
        """Context manager entry"""
        from DocumentFormat.OpenXml import PresentationDocumentType
        
        self._presentation_doc = self.builder.PresentationDocument.Create(
            self.filepath, 
            PresentationDocumentType.Presentation
        )
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - save and close"""
        if self._presentation_doc:
            self._presentation_doc.Save()
            self._presentation_doc.Close()
    
    def set_widescreen(self) -> 'Presentation':
        """Set presentation to 16:9 widescreen format"""
        if self._presentation_doc:
            presentation_part = self._presentation_doc.PresentationPart
            if presentation_part and presentation_part.Presentation:
                self.builder.PresentationExtensions.SetSlideSizeWidescreen(
                    presentation_part.Presentation
                )
        return self
        
    def add_slide(self) -> 'SlideWrapper':
        """Add a new slide to the presentation"""
        if not self._presentation_doc:
            raise DocLayerError("Presentation not initialized. Use 'with' statement.")
            
        # Create slide part
        slide_part = self._presentation_doc.PresentationPart.AddNewPart[
            self.builder.PresentationDocument.SlidePart
        ]()
        
        # Create slide with basic structure
        from DocumentFormat.OpenXml.Presentation import (
            Slide, CommonSlideData, ShapeTree, NonVisualGroupShapeProperties,
            NonVisualDrawingProperties, NonVisualGroupShapeDrawingProperties,
            ApplicationNonVisualDrawingProperties, GroupShapeProperties, 
            ColorMapOverride
        )
        from DocumentFormat.OpenXml.Drawing import TransformGroup, MasterColorMapping
        
        slide = Slide(
            CommonSlideData(
                ShapeTree(
                    NonVisualGroupShapeProperties(
                        NonVisualDrawingProperties(Id=1, Name=""),
                        NonVisualGroupShapeDrawingProperties(),
                        ApplicationNonVisualDrawingProperties()
                    ),
                    GroupShapeProperties(TransformGroup())
                )
            ),
            ColorMapOverride(MasterColorMapping())
        )
        
        slide_part.Slide = slide
        
        # Add to presentation slide list
        from DocumentFormat.OpenXml.Presentation import SlideIdList, SlideId
        
        slide_id_list = self._presentation_doc.PresentationPart.Presentation.SlideIdList
        if slide_id_list is None:
            slide_id_list = SlideIdList()
            self._presentation_doc.PresentationPart.Presentation.SlideIdList = slide_id_list
            
        slide_id = len(slide_id_list) + 256
        slide_id_list.Append(
            SlideId(
                Id=slide_id,
                RelationshipId=self._presentation_doc.PresentationPart.GetIdOfPart(slide_part)
            )
        )
        
        slide_wrapper = SlideWrapper(slide, self.builder)
        self._slides.append(slide_wrapper)
        return slide_wrapper

class SlideWrapper:
    """Python wrapper for individual slides"""
    
    def __init__(self, slide, builder: SlideBuilder):
        self.slide = slide
        self.builder = builder
        
    def add_title(self, text: str) -> 'SlideWrapper':
        """Add title to slide"""
        self.builder.SlideExtensions.AddTitle(self.slide, text)
        return self
        
    def add_textbox(self, text: str) -> 'SlideWrapper':
        """Add text box to slide"""
        self.builder.SlideExtensions.AddTextbox(self.slide, text)
        return self
        
    def add_table(self, rows: int, cols: int) -> 'SlideWrapper':
        """Add table to slide"""
        self.builder.SlideExtensions.AddTable(self.slide, rows, cols)
        return self
        
    def add_rectangle(self) -> 'SlideWrapper':
        """Add rectangle shape to slide"""
        self.builder.SlideExtensions.AddRectangle(self.slide)
        return self
        
    def add_footnote(self, text: str = "Source:") -> 'SlideWrapper':
        """Add footnote to slide"""
        self.builder.SlideExtensions.AddFootnote(self.slide, text)
        return self
        
    def add_page_number(self, page_num: str) -> 'SlideWrapper':
        """Add page number to slide"""
        self.builder.SlideExtensions.AddPageNumber(self.slide, page_num)
        return self
        
    def get_shape_tree(self) -> 'ShapeTreeWrapper':
        """Get shape tree for advanced operations"""
        shape_tree = self.slide.CommonSlideData.ShapeTree
        return ShapeTreeWrapper(shape_tree, self.builder)

class ShapeTreeWrapper:
    """Python wrapper for shape tree operations"""
    
    def __init__(self, shape_tree, builder: SlideBuilder):
        self.shape_tree = shape_tree
        self.builder = builder
        
    def add_rectangle(self, x: int, y: int, height: int, width: int) -> 'ShapeTreeWrapper':
        """Add positioned rectangle"""
        self.builder.ShapeTreeExtensions.AddRectangle(self.shape_tree, x, y, height, width)
        return self
        
    def add_circle(self, x: int, y: int, height: int, width: int) -> 'ShapeTreeWrapper':
        """Add circle shape"""
        self.builder.ShapeTreeExtensions.AddCircle(self.shape_tree, x, y, height, width)
        return self
        
    def add_triangle(self, x: int, y: int, height: int, width: int) -> 'ShapeTreeWrapper':
        """Add triangle shape"""
        self.builder.ShapeTreeExtensions.AddTriangle(self.shape_tree, x, y, height, width)
        return self
        
    def add_textbox(self, text: str, x: int, y: int) -> 'ShapeTreeWrapper':
        """Add positioned text box"""
        self.builder.ShapeTreeExtensions.AddTextbox(self.shape_tree, text, x, y)
        return self
        
    def add_right_arrow(self, x: int, y: int, height: int, width: int) -> 'ShapeTreeWrapper':
        """Add right arrow shape"""
        self.builder.ShapeTreeExtensions.AddRightArrow(self.shape_tree, x, y, height, width)
        return self

# Convenience functions for common operations
def create_basic_presentation(filepath: str, title: str, content: str = "") -> bytes:
    """Create a basic presentation with title and content"""
    builder = SlideBuilder()
    
    with builder.create_presentation(filepath) as pres:
        pres.set_widescreen()
        
        slide = pres.add_slide()
        slide.add_title(title)
        
        if content:
            slide.add_textbox(content)
            
        slide.add_footnote()
        slide.add_page_number("1")
    
    # Return file content as bytes for cloud usage
    with open(filepath, 'rb') as f:
        return f.read()

def create_data_presentation(filepath: str, title: str, data: List[List[str]]) -> bytes:
    """Create presentation with data table"""
    builder = SlideBuilder()
    
    with builder.create_presentation(filepath) as pres:
        pres.set_widescreen()
        
        slide = pres.add_slide()
        slide.add_title(title)
        slide.add_table(len(data), len(data[0]) if data else 2)
        slide.add_footnote("Generated by DocLayer")
        slide.add_page_number("1")
    
    with open(filepath, 'rb') as f:
        return f.read()

# Export public API
__all__ = [
    'SlideBuilder', 'Presentation', 'SlideWrapper', 'ShapeTreeWrapper',
    'create_basic_presentation', 'create_data_presentation', 'DocLayerError'
]