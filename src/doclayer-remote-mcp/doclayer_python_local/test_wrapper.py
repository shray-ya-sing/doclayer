"""
Test script for DocLayer Python wrapper
"""

import sys
from pathlib import Path

# Add parent directory to path to import the package
sys.path.insert(0, str(Path(__file__).parent))

from doclayer_python import create_title_slide, create_presentation_with_theme, DocLayerError

def test_create_title_slide():
    """Test creating a title slide presentation"""
    print("[Test 1] Basic Title Slide")
    print("-" * 50)
    
    output_path = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    output_path.parent.mkdir(exist_ok=True)
    
    try:
        print(f"Creating presentation at: {output_path}")
        
        pptx_bytes = create_title_slide(
            filepath=str(output_path),
            title="Welcome to DocLayer Python",
            subtitle="PowerPoint Generation from Python via C#",
            footnote="Source: DocLayer Python Wrapper v0.1"
        )
        
        print(f"✓ Success! Created presentation")
        print(f"✓ File size: {len(pptx_bytes)} bytes")
        print(f"✓ Saved to: {output_path}")
        
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_create_presentation_with_theme():
    """Test creating a presentation with custom theme"""
    print("\n[Test 2] Presentation with Custom Theme")
    print("-" * 50)
    
    output_path = Path(__file__).parent / "test_outputs" / "python_test_theme.pptx"
    output_path.parent.mkdir(exist_ok=True)
    
    try:
        print(f"Creating presentation with custom theme at: {output_path}")
        
        pptx_bytes = create_presentation_with_theme(
            filepath=str(output_path),
            title="Custom Theme Test",
            subtitle="Arial font with custom colors from Python",
            footnote="Source: DocLayer Python Theme Test",
            font_name="Arial",
            accent_colors=["FF5733", "33FF57", "3357FF", "F3FF33"]
        )
        
        print(f"✓ Success! Created presentation with custom theme")
        print(f"✓ File size: {len(pptx_bytes)} bytes")
        print(f"✓ Saved to: {output_path}")
        print("  Theme: Arial font, custom accent colors")
        
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Testing DocLayer Python Wrapper")
    print("=" * 50)
    print()
    
    success1 = test_create_title_slide()
    success2 = test_create_presentation_with_theme()
    
    print("\n" + "=" * 50)
    if success1 and success2:
        print("✓ All tests passed!")
        print("\nOpen the files in PowerPoint to view the results!")
        sys.exit(0)
    else:
        print("✗ Some tests failed")
        sys.exit(1)
