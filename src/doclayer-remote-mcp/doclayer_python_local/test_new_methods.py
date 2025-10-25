"""
Test script for new DocLayer Python wrapper methods
Tests slide extraction, editing, and rendering functionality
"""

import sys
from pathlib import Path

# Add parent directory to path to import the package
sys.path.insert(0, str(Path(__file__).parent))

from doclayer_python import DocLayerClient, DocLayerError

def test_get_slide_count():
    """Test getting slide count from a presentation"""
    print("[Test 1] Get Slide Count")
    print("-" * 50)
    
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        client = DocLayerClient()
        count = client.get_slide_count(str(test_file))
        
        print(f"✓ Success! Slide count: {count}")
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_extract_slide_content():
    """Test extracting content from a specific slide"""
    print("\n[Test 2] Extract Slide Content")
    print("-" * 50)
    
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        client = DocLayerClient()
        content = client.extract_slide_content(str(test_file), slide_number=1)
        
        print(f"✓ Success! Extracted content from slide 1:")
        print(f"  - Number of shapes: {len(content.get('shapes', []))}")
        print(f"  - Number of pictures: {len(content.get('pictures', []))}")
        
        # Print shape details
        for i, shape in enumerate(content.get('shapes', [])):
            print(f"  Shape {i+1}: {shape.get('name', 'Unnamed')} - '{shape.get('text', '')[:50]}'")
        
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_extract_all_slides():
    """Test extracting content from all slides"""
    print("\n[Test 3] Extract All Slides")
    print("-" * 50)
    
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        client = DocLayerClient()
        all_content = client.extract_all_slides(str(test_file))
        
        print(f"✓ Success! Extracted content from {len(all_content)} slide(s):")
        for slide_num, content in all_content.items():
            print(f"  Slide {slide_num}: {len(content.get('shapes', []))} shapes, {len(content.get('pictures', []))} pictures")
        
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_edit_slide_text():
    """Test editing text on a slide"""
    print("\n[Test 4] Edit Slide Text")
    print("-" * 50)
    
    # Create a copy to edit
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    output_file = Path(__file__).parent / "test_outputs" / "python_test_edited.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        import shutil
        shutil.copy(test_file, output_file)
        
        client = DocLayerClient()
        
        # First, get the slide content to find shape names
        content = client.extract_slide_content(str(output_file), slide_number=1)
        shapes = content.get('shapes', [])
        
        if not shapes:
            print(f"✗ No shapes found in the slide")
            return False
        
        # Edit the first shape's text
        shape_name = shapes[0].get('name', '')
        print(f"  Editing shape: {shape_name}")
        
        client.edit_slide_text(
            str(output_file),
            slide_number=1,
            element_name=shape_name,
            new_text="EDITED: This text was changed by Python wrapper!"
        )
        
        print(f"✓ Success! Text edited in {output_file}")
        return True
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_render_slide_to_image():
    """Test rendering a single slide to JPEG"""
    print("\n[Test 5] Render Slide to Image")
    print("-" * 50)
    
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        client = DocLayerClient()
        image_path = client.render_slide_to_image(str(test_file), slide_number=1)
        
        print(f"✓ Success! Rendered slide to: {image_path}")
        
        # Check if file exists
        if Path(image_path).exists():
            size = Path(image_path).stat().st_size
            print(f"  Image size: {size} bytes")
            return True
        else:
            print(f"✗ Image file not found at: {image_path}")
            return False
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_render_all_slides():
    """Test rendering all slides to JPEG images"""
    print("\n[Test 6] Render All Slides to Images")
    print("-" * 50)
    
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    
    if not test_file.exists():
        print(f"✗ Test file not found: {test_file}")
        return False
    
    try:
        client = DocLayerClient()
        image_paths = client.render_all_slides(str(test_file))
        
        print(f"✓ Success! Rendered {len(image_paths)} slide(s):")
        for i, image_path in enumerate(image_paths, 1):
            if Path(image_path).exists():
                size = Path(image_path).stat().st_size
                print(f"  Slide {i}: {image_path} ({size} bytes)")
            else:
                print(f"  ✗ Slide {i}: File not found at {image_path}")
        
        return len(image_paths) > 0 and all(Path(p).exists() for p in image_paths)
        
    except DocLayerError as e:
        print(f"✗ DocLayer Error: {e}")
        return False
    except Exception as e:
        print(f"✗ Unexpected Error: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Testing DocLayer Python Wrapper - New Methods")
    print("=" * 50)
    print()
    
    # Check if test file exists, if not, create it
    test_file = Path(__file__).parent / "test_outputs" / "python_test_title_slide.pptx"
    if not test_file.exists():
        print("Test file not found. Please run test_wrapper.py first to create test files.")
        sys.exit(1)
    
    results = []
    results.append(("Get Slide Count", test_get_slide_count()))
    results.append(("Extract Slide Content", test_extract_slide_content()))
    results.append(("Extract All Slides", test_extract_all_slides()))
    results.append(("Edit Slide Text", test_edit_slide_text()))
    results.append(("Render Slide to Image", test_render_slide_to_image()))
    results.append(("Render All Slides", test_render_all_slides()))
    
    print("\n" + "=" * 50)
    print("Test Results:")
    print("-" * 50)
    
    for test_name, success in results:
        status = "✓ PASSED" if success else "✗ FAILED"
        print(f"{test_name}: {status}")
    
    passed = sum(1 for _, success in results if success)
    total = len(results)
    
    print("-" * 50)
    print(f"Total: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n✓ All tests passed!")
        sys.exit(0)
    else:
        print(f"\n✗ {total - passed} test(s) failed")
        sys.exit(1)
