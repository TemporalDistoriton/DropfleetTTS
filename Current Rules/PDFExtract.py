import fitz  # PyMuPDF
from PIL import Image
import io
import re
import os

def extract_ship_name(page):
    """
    Extract the ship name and type from the top of the page.
    Returns the ship name or a default name if not found.
    """
    # Get text from the page
    text = page.get_text("text")
    
    # Split into lines and get the first few lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    if len(lines) >= 3:
        # Line 0 contains ship name (e.g., "Santiago")
        # Line 2 contains ship type (e.g., "Corvette")
        
        ship_name = lines[0].strip()
        ship_type = lines[2].strip()
        
        # Combine ship name and type
        full_name = ship_name + " " + ship_type
        
        # Clean up the full name (remove special characters)
        full_name = re.sub(r'[<>:"/\\|?*]', '', full_name)
        
        return full_name
    
    return "Unknown_Ship"

def get_crop_rect(page, search_text="Famous Ships of the class"):
    """
    Find the text on the page and return a clip rectangle that excludes
    everything from that text onwards.
    
    Returns:
        fitz.Rect: Clipping rectangle, or None if text not found (use full page)
    """
    # Search for the text
    text_instances = page.search_for(search_text)
    
    if text_instances:
        # Get the first instance of the text
        first_match = text_instances[0]
        
        # Get the y-coordinate where the text starts
        crop_y = first_match.y0  # Top of the text bounding box
        
        # Create a rectangle from top of page to just before the text
        # Format: fitz.Rect(x0, y0, x1, y1)
        page_rect = page.rect
        clip_rect = fitz.Rect(0, 0, page_rect.width, crop_y)
        
        return clip_rect
    
    # Return None if text not found (will use full page)
    return None

def pdf_to_png(pdf_path, output_folder="output", dpi=300, crop_text="Famous Ships of the class"):
    """
    Convert each page of a PDF to PNG with ship name in filename.
    If crop_text is found on a page, only render content before that text.
    
    Args:
        pdf_path: Path to the input PDF file
        output_folder: Folder to save PNG files (default: 'output')
        dpi: Resolution for the output images (default: 300)
        crop_text: Text to search for cropping (default: 'Famous Ships of the class')
    """
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Open the PDF
    pdf_document = fitz.open(pdf_path)
    
    print(f"Processing PDF: {pdf_path}")
    print(f"Total pages: {len(pdf_document)}")
    
    # Process each page
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        
        # Extract ship name from the page
        ship_name = extract_ship_name(page)
        
        # Create the output filename
        output_filename = f"{ship_name}_CardFrontImage.png"
        output_path = os.path.join(output_folder, output_filename)
        
        # Calculate zoom factor based on DPI (72 is the default PDF DPI)
        zoom = dpi / 72
        mat = fitz.Matrix(zoom, zoom)
        
        # Check if we need to crop the page
        clip_rect = get_crop_rect(page, crop_text)
        
        if clip_rect:
            # Render only the cropped portion
            pix = page.get_pixmap(matrix=mat, clip=clip_rect)
            print(f"Page {page_num + 1}: Cropped at '{crop_text}' - Saved as '{output_filename}'")
        else:
            # Render the full page
            pix = page.get_pixmap(matrix=mat)
            print(f"Page {page_num + 1}: Full page - Saved as '{output_filename}'")
        
        # Save as PNG
        pix.save(output_path)
    
    pdf_document.close()
    print(f"\nConversion complete! Files saved in '{output_folder}' folder.")

# Example usage
if __name__ == "__main__":
    import sys
    
    # Check if PDF filename was provided as command line argument
    if len(sys.argv) < 2:
        print("Usage: python script.py <pdf_filename>")
        print("Example: python script.py my_ships.pdf")
        sys.exit(1)
    
    pdf_file = sys.argv[1]
    
    # Create output folder name from PDF filename (without extension)
    pdf_basename = os.path.splitext(os.path.basename(pdf_file))[0]
    output_folder = pdf_basename
    
    # DPI setting (can be adjusted here if needed)
    dpi = 300  # Higher DPI = better quality but larger files
    
    # Check if file exists
    if os.path.exists(pdf_file):
        pdf_to_png(pdf_file, output_folder, dpi)
    else:
        print(f"Error: File '{pdf_file}' not found!")
        print("Please check the filename and try again.")