"""
Test script for DocLayer Python wrapper
"""

import sys
from pathlib import Path

# Add parent directory to path to import the package
sys.path.insert(0, str(Path(__file__).parent))

from doclayer_python import create_title_slide, DocLayerError

def test_create_title_slide():
    """Test creating a title slide presentation"""
    print("Testing DocLayer Python Wrapper")
    print("=" * 50)
    
    output_path = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    output_path.parent.mkdir(exist_ok=True)
    
    try:
        print(f"\nCreating presentation at: {output_path}")
        
        pptx_bytes = create_title_slide(
            filepath=str(output_path),
            title="Welcome to DocLayer Python",
            subtitle="PowerPoint Generation from Python via C#",
            footnote="Source: DocLayer Python Wrapper v0.1"
        )
        
        print(f"✓ Success! Created presentation")
        print(f"✓ File size: {len(pptx_bytes)} bytes")
        print(f"✓ Saved to: {output_path}")
        print("\nOpen the file in PowerPoint to view the slide!")
        
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
    success = test_create_title_slide()
    sys.exit(0 if success else 1)
