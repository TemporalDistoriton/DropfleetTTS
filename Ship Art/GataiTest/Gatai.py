#!/usr/bin/env python3
"""
Image Combination Script
Combines multiple images from a folder with a base image using specific transformations
"""

from PIL import Image
import os
from pathlib import Path


def combine_images(image1, image2):
    """
    Combines two images with specific transformations.
    
    Args:
        image1: PIL Image object (base image)
        image2: PIL Image object (image to overlay)
        
    Returns:
        Combined PIL Image object
    """
    # Make a copy of image1 so we don't modify the original
    result = image1.copy()
    
    # Step 1: Create a blank 500x365 canvas and fit Image 2 onto it (preserving aspect ratio)
    # Note: We use 500x365 so that after 270° rotation, it becomes 365x500
    canvas_width, canvas_height = 500, 365
    
    # Calculate the scaling factor to fit Image 2 within the canvas while preserving aspect ratio
    width_ratio = canvas_width / image2.width
    height_ratio = canvas_height / image2.height
    scale_factor = min(width_ratio, height_ratio)
    
    # Calculate new dimensions
    new_width = int(image2.width * scale_factor)
    new_height = int(image2.height * scale_factor)
    
    # Resize Image 2 with aspect ratio preserved
    image2_resized = image2.resize((new_width, new_height), Image.Resampling.LANCZOS)
    
    # Create blank canvas (transparent background)
    canvas = Image.new('RGBA', (canvas_width, canvas_height), (0, 0, 0, 0))
    
    # Calculate position to center Image 2 on the canvas
    paste_x = (canvas_width - new_width) // 2
    paste_y = (canvas_height - new_height) // 2
    
    # Paste the resized Image 2 onto the canvas
    if image2_resized.mode != 'RGBA':
        image2_resized = image2_resized.convert('RGBA')
    canvas.paste(image2_resized, (paste_x, paste_y), image2_resized)
    
    image2_scaled = canvas
    
    # Step 2: Rotate Image 2 by 270 degrees
    image2_rotated = image2_scaled.rotate(270, expand=True)
    
    # Step 3: Paste rotated Image 2 into Image 1 at position (55, 450)
    # Convert to RGBA to handle transparency if needed
    if result.mode != 'RGBA':
        result = result.convert('RGBA')
    if image2_rotated.mode != 'RGBA':
        image2_rotated = image2_rotated.convert('RGBA')
    
    result.paste(image2_rotated, (55, 450), image2_rotated)
    
    # Step 4: Flip the rotated Image 2 vertically
    image2_flipped = image2_rotated.transpose(Image.Transpose.FLIP_TOP_BOTTOM)
    
    # Step 5: Paste the flipped version at position (530, 450)
    result.paste(image2_flipped, (530, 450), image2_flipped)
    
    return result


def process_folder(image1_path, image2_folder):
    """
    Process all images in a folder, combining each with the base image.
    
    Args:
        image1_path: Path to the base image (Image 1)
        image2_folder: Path to folder containing images to process
    """
    # Supported image extensions
    image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'}
    
    # Create output folder if it doesn't exist
    output_folder = Path("Output")
    output_folder.mkdir(exist_ok=True)
    
    # Load the base image once
    print(f"Loading base image: {image1_path}")
    image1 = Image.open(image1_path)
    
    # Get the folder path
    folder_path = Path(image2_folder)
    
    if not folder_path.exists():
        print(f"Error: Folder '{image2_folder}' does not exist")
        return
    
    # Get all image files in the folder
    image_files = [f for f in folder_path.iterdir() 
                   if f.is_file() and f.suffix.lower() in image_extensions]
    
    if not image_files:
        print(f"No image files found in '{image2_folder}'")
        return
    
    print(f"Found {len(image_files)} image(s) to process")
    print()
    
    # Process each image
    for image_file in image_files:
        try:
            print(f"Processing: {image_file.name}")
            
            # Load Image 2
            image2 = Image.open(image_file)
            
            # Combine the images
            result = combine_images(image1, image2)
            
            # Create output filename: original_name_ShipModel.ext
            output_filename = f"{image_file.stem}_ShipModel{image_file.suffix}"
            output_path = output_folder / output_filename
            
            # Save the result
            result.save(output_path)
            print(f"  ✓ Saved: {output_path}")
            print()
            
        except Exception as e:
            print(f"  ✗ Error processing {image_file.name}: {e}")
            print()
    
    print(f"Processing complete! All images saved to '{output_folder}' folder")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 3:
        print("Usage: python combine_images.py <image1_path> <image2_folder>")
        print("Example: python combine_images.py base.png ships/")
        print()
        print("This will process all images in the specified folder,")
        print("combine each with the base image, and save results to 'Output' folder")
        print("with '_ShipModel' appended to the filename.")
        sys.exit(1)
    
    image1_path = sys.argv[1]
    image2_folder = sys.argv[2]
    
    try:
        process_folder(image1_path, image2_folder)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)