#!/usr/bin/env python3
"""
PDF Image Extractor
Extracts all images from PDFs in the current directory and saves them
with filenames based on the first line of text from each page.
All images are saved as PNG with transparency preserved.
"""

import os
import re
from pathlib import Path
from io import BytesIO

try:
    import fitz  # PyMuPDF
except ImportError:
    print("PyMuPDF not found. Installing...")
    import subprocess
    subprocess.check_call(["pip", "install", "pymupdf", "--break-system-packages"])
    import fitz

try:
    from PIL import Image
except ImportError:
    print("Pillow not found. Installing...")
    import subprocess
    subprocess.check_call(["pip", "install", "pillow", "--break-system-packages"])
    from PIL import Image


def sanitize_filename(filename):
    """Remove or replace characters that aren't safe for filenames."""
    # Remove or replace invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    # Replace multiple spaces with single space
    filename = re.sub(r'\s+', ' ', filename)
    # Trim whitespace
    filename = filename.strip()
    # Limit length to avoid filesystem issues
    if len(filename) > 100:
        filename = filename[:100]
    # If empty after sanitization, use a default
    if not filename:
        filename = "untitled"
    return filename


def extract_images_from_pdf(pdf_path):
    """Extract all images from a PDF and save them with proper naming."""
    pdf_name = pdf_path.stem
    output_folder = Path(pdf_path.parent) / pdf_name
    output_folder.mkdir(exist_ok=True)
    
    print(f"\nProcessing: {pdf_path.name}")
    print(f"Output folder: {output_folder}")
    
    try:
        doc = fitz.open(pdf_path)
        total_images = 0
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            
            # Extract the first line of text from the page
            text = page.get_text()
            first_line = text.split('\n')[0].strip() if text else f"page_{page_num + 1}"
            
            # Sanitize the first line for use as filename
            base_filename = sanitize_filename(first_line)
            if not base_filename:
                base_filename = f"page_{page_num + 1}"
            
            # Get all images from the page
            image_list = page.get_images(full=True)
            
            if image_list:
                print(f"  Page {page_num + 1}: Found {len(image_list)} image(s) - First line: '{first_line[:50]}...'")
            
            # Extract each image
            for img_index, img in enumerate(image_list):
                xref = img[0]
                
                try:
                    # Extract image
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # Open image with PIL to convert to PNG
                    pil_image = Image.open(BytesIO(image_bytes))
                    
                    # Convert to RGBA if the image has transparency, otherwise RGB
                    if pil_image.mode in ('RGBA', 'LA', 'P'):
                        # Image has or might have transparency
                        if pil_image.mode == 'P':
                            # Palette mode - convert to RGBA to preserve any transparency
                            pil_image = pil_image.convert('RGBA')
                        elif pil_image.mode == 'LA':
                            # Grayscale with alpha
                            pil_image = pil_image.convert('RGBA')
                        # RGBA is already good
                    else:
                        # No transparency - convert to RGB for smaller file size
                        if pil_image.mode != 'RGB':
                            pil_image = pil_image.convert('RGB')
                    
                    # Create filename (always .png now)
                    if len(image_list) == 1:
                        # Single image on page
                        filename = f"{base_filename}.png"
                    else:
                        # Multiple images on page - append number
                        filename = f"{base_filename}_{img_index + 1}.png"
                    
                    # Save image
                    image_path = output_folder / filename
                    
                    # Handle duplicate filenames
                    counter = 1
                    original_path = image_path
                    while image_path.exists():
                        stem = original_path.stem
                        image_path = output_folder / f"{stem}_dup{counter}.png"
                        counter += 1
                    
                    # Save as PNG
                    pil_image.save(image_path, 'PNG', optimize=True)
                    
                    total_images += 1
                    print(f"    Saved: {image_path.name}")
                    
                except Exception as e:
                    print(f"    Error extracting image {img_index + 1} from page {page_num + 1}: {e}")
        
        doc.close()
        print(f"Total images extracted: {total_images}")
        
    except Exception as e:
        print(f"Error processing {pdf_path.name}: {e}")


def main():
    """Main function to process all PDFs in the current directory."""
    # Get current directory
    current_dir = Path.cwd()
    
    # Find all PDF files
    pdf_files = list(current_dir.glob("*.pdf"))
    
    if not pdf_files:
        print("No PDF files found in the current directory.")
        return
    
    print(f"Found {len(pdf_files)} PDF file(s) to process")
    
    for pdf_path in pdf_files:
        extract_images_from_pdf(pdf_path)
    
    print("\n✓ All PDFs processed!")


if __name__ == "__main__":
    main()