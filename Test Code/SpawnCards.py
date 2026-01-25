#!/usr/bin/env python3
"""
TTS Ship Card Generator
Creates Tabletop Simulator ship cards from Excel data and a template JSON save file.
"""

import json
import pandas as pd
import copy
import re
import sys
import random
import uuid
from pathlib import Path


def load_excel_data(excel_path):
    """Load ship data from Excel file."""
    print(f"Loading Excel data from: {excel_path}")
    df = pd.read_excel(excel_path)
    print(f"Loaded {len(df)} ships from Excel")
    return df

def random_bag_colour(min_val=0.25, max_val=0.85):
    """Generate a pleasant random colour (avoids too-dark or too-bright)."""
    return {
        "r": random.uniform(min_val, max_val),
        "g": random.uniform(min_val, max_val),
        "b": random.uniform(min_val, max_val),
    }


def get_ship_card_image_url(row):
    """
    Returns the card face (top) image URL from the Excel row.
    Supports multiple possible column names.
    """
    candidates = [
        "SHIPCardURL",
        "ShipCardURL",
        "Ship Card URL",
        "Ship Card Image",
        "Card Image",   # your existing column
    ]

    for col in candidates:
        if col in row and not pd.isna(row[col]) and str(row[col]).strip():
            return str(row[col]).strip()

    return ""


def set_tts_custom_tile_face_image(card_obj, face_url):
    """
    Sets the Custom_Tile 'top image' (face) to face_url in the TTS JSON object.
    For Custom_Tile objects, this is typically CustomImage.ImageURL.
    """
    if not face_url:
        return

    if "CustomImage" not in card_obj or not isinstance(card_obj["CustomImage"], dict):
        card_obj["CustomImage"] = {}

    card_obj["CustomImage"]["ImageURL"] = face_url


def load_tts_template(json_path):
    """Load the TTS save file template."""
    print(f"Loading TTS template from: {json_path}")
    with open(json_path, 'r', encoding='utf-8') as f:
        tts_data = json.load(f)
    
    # Find the ship card object (should be Custom_Tile with Lua script)
    ship_card_template = None
    for i, obj in enumerate(tts_data['ObjectStates']):
        if obj.get('Name') == 'Custom_Tile' and 'SHIP_ID' in obj.get('LuaScript', ''):
            ship_card_template = obj
            print(f"Found ship card template at index {i}")
            break
    
    if not ship_card_template:
        raise ValueError("Could not find ship card template in TTS save file!")
    
    return tts_data, ship_card_template


def replace_lua_variables(lua_script, row):
    """Replace variables in the Lua script with values from Excel row."""
    
    # Handle NaN values - convert to empty string or default
    def safe_value(val, default=''):
        if pd.isna(val):
            return default
        return val
    
    # Define replacement patterns (old_value, new_value)
    replacements = [
        # SHIP_ID
        (r'local SHIP_ID = "BASETEMPLATECARD"', 
         f'local SHIP_ID = "{safe_value(row["Name"], "UNKNOWN")}"'),
        
        # Ship name
        (r"local name = 'BASE TEMPLATE CARD'", 
         f"local name = '{safe_value(row['Name'], 'UNKNOWN')}'"),
        
        # Points
        (r'local points = 66', 
         f'local points = {int(safe_value(row["Points"], 0))}'),
        
        # Base size
        (r'local baseSize = 30', 
         f'local baseSize = {int(safe_value(row["Size"], 30))}'),
        
        # Thrust
        (r'local thrust = 6', 
         f'local thrust = {safe_value(row["Thrust"], 0)}'),
        
        # Scan
        (r'local scan = 6', 
         f'local scan = {safe_value(row["Scan"], 0)}'),
        
        # Signature
        (r'local sig = 66', 
         f'local sig = {safe_value(row["Sig"], 0)}'),
        
        # Hull/Health
        (r'local health = 2', 
         f'local health = {safe_value(row["Hull"], 0)}'),
        
        # Model Image
        (r"local modelImage = 'https://raw\.githubusercontent\.com/TemporalDistoriton/DropfleetTTS/main/Ship Art/30mmTestImage\.png'",
         f"local modelImage = '{safe_value(row['Model Image'], 'https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Ship Art/30mmTestImage.png')}'"),
        
        # Card Image
        (r"local cardFrontImage = 'https://raw\.githubusercontent\.com/TemporalDistoriton/DropfleetTTS/main/Assets/2dModel/MissingNo\.png'",
         f"local cardFrontImage = '{safe_value(row['Card Image'], 'https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/2dModel/MissingNo.png')}'"),
    ]
    
    modified_script = lua_script
    for pattern, replacement in replacements:
        modified_script = re.sub(pattern, replacement, modified_script)
    
    return modified_script


def create_ship_card(template, row):
    """Create a new ship card object from template with modified values."""
    new_card = copy.deepcopy(template)

    # Update Name (TTS Nickname)
    new_card['Nickname'] = str(row['Name']) if not pd.isna(row['Name']) else 'UNKNOWN'

    # Update Description (Type)
    new_card['Description'] = str(row['Type']) if not pd.isna(row['Type']) else ''

    # Update Lua Script with new variables
    new_card['LuaScript'] = replace_lua_variables(template['LuaScript'], row)

    # C) Set the card TOP image to the passed URL from Excel
    face_url = get_ship_card_image_url(row)
    set_tts_custom_tile_face_image(new_card, face_url)

    return new_card


def create_bag_object(bag_name, contained_objects, posX=0.0, posY=3.0, posZ=0.0, colour=None):
    """Create a TTS bag object to hold ship cards."""

    if colour is None:
        colour = {"r": 0.7, "g": 0.7, "b": 0.7}

    bag = {
        "GUID": uuid.uuid4().hex[:6],
        "Name": "Bag",
        "Transform": {
            "posX": float(posX),
            "posY": float(posY),
            "posZ": float(posZ),
            "rotX": 0.0,
            "rotY": 0.0,
            "rotZ": 0.0,
            "scaleX": 1.0,
            "scaleY": 1.0,
            "scaleZ": 1.0
        },
        "Nickname": bag_name,
        "Description": f"Ships from {bag_name}",
        "GMNotes": "",
        "ColorDiffuse": colour,
        "LayoutGroupSortIndex": 0,
        "Value": 0,
        "Locked": False,
        "Grid": True,
        "Snap": True,
        "IgnoreFoW": False,
        "MeasureMovement": False,
        "DragSelectable": True,
        "Autoraise": True,
        "Sticky": True,
        "Tooltip": True,
        "GridProjection": False,
        "HideWhenFaceDown": False,
        "Hands": False,
        "MaterialIndex": -1,
        "MeshIndex": -1,
        "Bag": {"Order": 0},
        "ContainedObjects": contained_objects
    }

    return bag


def generate_tts_save(excel_path, template_path, output_path):
    """Main function to generate TTS save file from Excel data."""
        # Grid layout settings (tweak to taste)
    grid_spacing_x = 6.0
    grid_spacing_z = 6.0
    bags_per_row = 6
    start_x = 0.0
    start_z = 0.0
    bag_y = 3.0

    # Load data
    df = load_excel_data(excel_path)
    tts_data, ship_template = load_tts_template(template_path)
    
    # Group ships by PDF ID
    print("\nGrouping ships by PDF ID...")
    grouped = df.groupby('PDF ID')
    
    # Create bags for each PDF ID
    bags = []
    total_cards = 0
    
    for pdf_id, group in grouped:
        if pd.isna(pdf_id):
            pdf_id = "Uncategorized"
        
        print(f"\nProcessing PDF ID: {pdf_id} ({len(group)} ships)")
        
        # Create ship cards for this group
        ship_cards = []
        for idx, row in group.iterrows():
            card = create_ship_card(ship_template, row)
            ship_cards.append(card)
            print(f"  - Created card: {row['Name']}")
            total_cards += 1
        
        bag_index = len(bags)

        col = bag_index % bags_per_row
        row_i = bag_index // bags_per_row

        posX = start_x + col * grid_spacing_x
        posZ = start_z + row_i * grid_spacing_z

        colour = random_bag_colour()

        bag = create_bag_object(
            bag_name=str(pdf_id),
            contained_objects=ship_cards,
            posX=posX,
            posY=bag_y,
            posZ=posZ,
            colour=colour
        )

        bags.append(bag)

    
    # Remove the original ship card template and add bags
    tts_data['ObjectStates'] = [obj for obj in tts_data['ObjectStates'] 
                                 if not (obj.get('Name') == 'Custom_Tile' and 'SHIP_ID' in obj.get('LuaScript', ''))]
    tts_data['ObjectStates'].extend(bags)
    
    # Save the output
    print(f"\nSaving TTS file to: {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(tts_data, f, indent=2)
    
    print(f"\n✓ Success!")
    print(f"  Created {len(bags)} bags")
    print(f"  Generated {total_cards} ship cards")
    print(f"  Output saved to: {output_path}")


def main():
    """Main entry point."""
    if len(sys.argv) != 4:
        print("Usage: python generate_tts_cards.py <excel_file> <template_json> <output_json>")
        print("\nExample:")
        print("  python generate_tts_cards.py ship_data.xlsx template.json output.json")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    template_path = sys.argv[2]
    output_path = sys.argv[3]
    
    # Validate input files exist
    if not Path(excel_path).exists():
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)
    
    if not Path(template_path).exists():
        print(f"Error: Template JSON file not found: {template_path}")
        sys.exit(1)
    
    # Generate the TTS save file
    generate_tts_save(excel_path, template_path, output_path)


if __name__ == "__main__":
    main()