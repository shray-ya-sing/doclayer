"""Test script to render slides from a presentation."""

import sys
import os

# Add python wrapper to path
sys.path.insert(0, r"C:\Users\shrey\projects\doclayer\python-wrapper")

from doclayer_python import DocLayerClient

def test_render():
    filepath = r"C:\Users\shrey\OneDrive\Desktop\docs\Netstar\Presentations\template_Discussion Materials.pptx"
    output_dir = r"C:\Users\shrey\projects\doclayer\mcp-server\test_output"
    
    # Create output directory
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"Testing render for: {filepath}")
    print(f"Output directory: {output_dir}")
    
    client = DocLayerClient()
    
    try:
        # Render all slides
        print("\nRendering all slides...")
        image_paths = client.render_all_slides(filepath)
        
        print(f"\nRendered {len(image_paths)} slides:")
        for i, img_path in enumerate(image_paths, 1):
            print(f"  Slide {i}: {img_path}")
            
            # Check if file exists
            if os.path.exists(img_path):
                size = os.path.getsize(img_path)
                print(f"    File size: {size} bytes")
            else:
                print(f"    ERROR: File not found!")
        
        # Also test single slide render
        print("\n\nTesting single slide render (slide 1)...")
        single_image = client.render_slide_to_image(filepath, 1)
        print(f"Single slide image: {single_image}")
        if os.path.exists(single_image):
            size = os.path.getsize(single_image)
            print(f"File size: {size} bytes")
        
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_render()
