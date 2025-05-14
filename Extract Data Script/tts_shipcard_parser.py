import os
import re
import json
import pandas as pd
import requests
from urllib.parse import urlparse

# Define the parameters we want to extract
PARAMETERS = [
    'baseScale',
    'health',
    'sig',
    'points',
    'modelImage',
    'name',
    'cardFrontImage'
    # Removed 'faction' from here as we're determining it from containers
]

# Valid factions list
VALID_FACTIONS = [
    'UCM',
    'PHR',
    'Shaltari',
    'Scourge',
    'Resistance',
    'Bioficers'
]

# Containers to ignore
IGNORED_CONTAINERS = [
    "Old 2.0 Content"
]

def extract_parameter(content, param_name):
    """
    Extract parameter value from the lua script
    """
    # Different patterns based on the parameter type
    if param_name in ['baseScale', 'health', 'sig', 'points']:
        # Number parameters (may have local or direct assignment)
        pattern = rf'(?:local\s+{param_name}\s*=\s*|{param_name}\s*=\s*)([0-9.]+)'
    elif param_name in ['modelImage', 'cardFrontImage']:
        # URL parameters (often within quotes)
        pattern = rf"(?:local\s+{param_name}\s*=\s*|{param_name}\s*=\s*)['\"](https?://[^'\"]+)['\"]"
    elif param_name == 'name':
        # Name parameter (string)
        pattern = rf"(?:local\s+{param_name}\s*=\s*|{param_name}\s*=\s*)['\"](.*?)['\"]"
    
    # Search for the pattern for parameters
    match = re.search(pattern, content)
    if match:
        return match.group(1).strip()
    
    # Return default values if not found
    if param_name in ['baseScale', 'health', 'sig', 'points']:
        return 0
    return "Unknown"

def determine_faction_from_container(container_name):
    """
    Determine faction based on container name
    """
    container_name = container_name.lower()
    
    # Check for each faction in the container name
    for faction in VALID_FACTIONS:
        if faction.lower() in container_name:
            return faction
    
    # Special case for common variations
    if "ucm" in container_name:
        return "UCM"
    if "phr" in container_name:
        return "PHR"
    if "shaltari" in container_name:
        return "Shaltari"
    if "scourge" in container_name:
        return "Scourge"
    if "resistance" in container_name:
        return "Resistance"
    if "bio" in container_name:
        return "Bioficers"
    
    # Default to Neutral if no faction is identified
    return "Neutral"

def is_ship_card_script(content):
    """
    Check if the content contains a ship card script based on key indicators
    """
    # Look for specific indicators in the ship card sample you provided
    indicators = [
        "rebuildUI()",
        "createModel",
        "cardFrontImage",
        "modelImage",
        "baseScale",
        "onSave()"
    ]
    
    # Count how many indicators are present
    matches = sum(1 for indicator in indicators if indicator in content)
    
    # If at least 3 indicators are found, consider it a ship card script
    return matches >= 3

def download_image(url, destination):
    """
    Download an image from URL and save it to the destination
    """
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        
        # Create the directory if it doesn't exist
        os.makedirs(os.path.dirname(destination), exist_ok=True)
        
        # Save the image
        with open(destination, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        
        print(f"Downloaded: {url} -> {destination}")
        return True
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def sanitize_filename(name):
    """
    Convert a string to a valid filename
    """
    # Replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    
    # Remove leading/trailing spaces
    return name.strip()

def should_skip_container(container_path):
    """
    Check if we should skip this container and its contents
    """
    for container_name in container_path:
        if container_name in IGNORED_CONTAINERS:
            return True
    return False

def process_tts_save_file(save_file_path):
    """
    Process a TTS save file (JSON), extract ship card data,
    download images, and create an Excel file
    """
    # Directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(save_file_path))
    if not script_dir:  # If empty string (current directory)
        script_dir = os.getcwd()
    
    # List to store extracted data
    extracted_data = []
    
    try:
        # Read the TTS save file
        with open(save_file_path, 'r', encoding='utf-8', errors='ignore') as file:
            save_data = json.load(file)
        
        # Process the save data
        # The structure of a TTS save file should have an "ObjectStates" array
        if "ObjectStates" in save_data:
            # Create a container hierarchy to track parent containers
            container_hierarchy = {}
            
            # First pass: build the container hierarchy
            build_container_hierarchy(save_data["ObjectStates"], container_hierarchy)
            
            # Second pass: process objects using the container hierarchy
            process_object_states(save_data["ObjectStates"], extracted_data, script_dir, container_hierarchy)
        else:
            print("Invalid TTS save file format. Missing 'ObjectStates' array.")
    except json.JSONDecodeError:
        print(f"Error: {save_file_path} is not a valid JSON file.")
        return 0
    except Exception as e:
        print(f"Error processing {save_file_path}: {e}")
        import traceback
        traceback.print_exc()
        return 0
    
    # Create Excel file if we found any data
    if extracted_data:
        df = pd.DataFrame(extracted_data)
        excel_file = os.path.join(script_dir, "ShipCardData.xlsx")
        try:
            df.to_excel(excel_file, index=False, engine='openpyxl')
            print(f"Saved data to Excel file: {excel_file}")
        except ImportError:
            # If openpyxl is not installed, save as CSV instead
            csv_file = os.path.join(script_dir, "ShipCardData.csv")
            df.to_csv(csv_file, index=False)
            print(f"Openpyxl not installed. Saved data to CSV file instead: {csv_file}")
            print("To save as Excel, install openpyxl: pip install openpyxl")
        return len(extracted_data)
    else:
        print("No ship card scripts found.")
        return 0

def build_container_hierarchy(object_states, hierarchy, parent_guid=None, depth=0):
    """
    Build a hierarchy of container GUIDs to track parent-child relationships
    """
    if depth > 10:  # Prevent infinite recursion
        return
    
    for obj in object_states:
        if "GUID" in obj:
            obj_guid = obj["GUID"]
            
            # Record the parent-child relationship
            if parent_guid:
                hierarchy[obj_guid] = {
                    "parent_guid": parent_guid,
                    "nickname": obj.get("Nickname", "")
                }
            else:
                hierarchy[obj_guid] = {
                    "parent_guid": None,
                    "nickname": obj.get("Nickname", "")
                }
            
            # Process contained objects recursively
            if "ContainedObjects" in obj and obj["ContainedObjects"]:
                build_container_hierarchy(obj["ContainedObjects"], hierarchy, obj_guid, depth + 1)

def find_container_path(hierarchy, guid):
    """
    Find the path of container names leading to this object
    """
    path = []
    current_guid = guid
    
    # Maximum number of iterations to prevent infinite loops
    for _ in range(10):  # Maximum depth of 10
        if current_guid not in hierarchy:
            break
        
        container_info = hierarchy[current_guid]
        parent_guid = container_info["parent_guid"]
        
        if parent_guid is None:
            break
        
        # Get the parent's nickname
        parent_nickname = hierarchy.get(parent_guid, {}).get("nickname", "")
        if parent_nickname:
            path.append(parent_nickname)
        
        current_guid = parent_guid
    
    return path[::-1]  # Reverse to get top-to-bottom order

def process_object_states(object_states, extracted_data, script_dir, container_hierarchy, parent_path=None, depth=0):
    """
    Recursively process object states to find and extract ship card scripts
    """
    if depth > 10:  # Prevent infinite recursion
        return
    
    if parent_path is None:
        parent_path = []
    
    # Skip this branch if it's in an ignored container
    if should_skip_container(parent_path):
        print(f"Skipping content in ignored container: {' > '.join(parent_path)}")
        return
    
    for obj in object_states:
        obj_guid = obj.get("GUID", "")
        obj_nickname = obj.get("Nickname", "Unnamed Object")
        
        # The current path includes the parent path plus this object
        current_path = parent_path + [obj_nickname] if obj_nickname else parent_path
        
        # Skip this object if it's in an ignored container
        if should_skip_container(current_path):
            print(f"Skipping content in ignored container: {' > '.join(current_path)}")
            continue
        
        # Check if the object has a LuaScript
        if "LuaScript" in obj and obj["LuaScript"]:
            lua_script = obj["LuaScript"]
            
            # Check if this is a ship card script
            if is_ship_card_script(lua_script):
                print(f"Found ship card script in object: {obj_nickname}")
                
                # Find container path for this object
                container_path = find_container_path(container_hierarchy, obj_guid)
                
                # Skip this object if it's in an ignored container
                if should_skip_container(container_path):
                    print(f"Skipping ship card in ignored container: {' > '.join(container_path)}")
                    continue
                
                # Extract parameters
                ship_data = {}
                for param in PARAMETERS:
                    ship_data[param] = extract_parameter(lua_script, param)
                
                # Add object info
                ship_data['object_name'] = obj_nickname
                ship_data['object_guid'] = obj_guid
                ship_data['container_path'] = " > ".join(container_path)
                
                # Determine faction from container path
                determined_faction = "Neutral"
                
                for container_name in container_path:
                    faction = determine_faction_from_container(container_name)
                    if faction != "Neutral":
                        determined_faction = faction
                        break
                
                # If we still don't have a faction, try the current object's name
                if determined_faction == "Neutral":
                    determined_faction = determine_faction_from_container(obj_nickname)
                
                # Use the container-determined faction as the main faction field
                ship_data['faction'] = determined_faction
                
                # Add to the list
                extracted_data.append(ship_data)
                
                # Download images if they have URLs
                if ship_data['modelImage'] != "Unknown" and ship_data['modelImage'].startswith("http"):
                    # Use the determined faction for the folder
                    faction = determined_faction
                    
                    # Sanitize name
                    name = sanitize_filename(ship_data['name'])
                    if name == "Unknown":
                        # Use the object name as fallback
                        name = sanitize_filename(obj_nickname)
                    
                    # Create folder for faction
                    faction_folder = os.path.join(script_dir, faction)
                    os.makedirs(faction_folder, exist_ok=True)
                    
                    print(f"Saving images for {name} in faction folder: {faction}")
                    
                    # Download model image
                    model_ext = os.path.splitext(urlparse(ship_data['modelImage']).path)[1]
                    if not model_ext:
                        model_ext = '.png'  # Default extension
                    model_filename = f"{name}_ModelImage{model_ext}"
                    download_image(ship_data['modelImage'], os.path.join(faction_folder, model_filename))
                    
                    # Download card front image
                    if ship_data['cardFrontImage'] != "Unknown" and ship_data['cardFrontImage'].startswith("http"):
                        card_ext = os.path.splitext(urlparse(ship_data['cardFrontImage']).path)[1]
                        if not card_ext:
                            card_ext = '.png'  # Default extension
                        card_filename = f"{name}_CardFrontImage{card_ext}"
                        download_image(ship_data['cardFrontImage'], os.path.join(faction_folder, card_filename))
        
        # Check if this object contains other objects (recursively)
        if "ContainedObjects" in obj and obj["ContainedObjects"]:
            process_object_states(obj["ContainedObjects"], extracted_data, script_dir, container_hierarchy, current_path, depth + 1)

if __name__ == "__main__":
    # Check if openpyxl is installed
    try:
        import openpyxl
        print("openpyxl is installed. Excel output is available.")
    except ImportError:
        print("WARNING: openpyxl is not installed. Will save as CSV instead of Excel.")
        print("To save as Excel, install openpyxl: pip install openpyxl")
    
    # Look for a TTS save file (.json) in the same directory as the script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if not script_dir:  # If empty string (current directory)
        script_dir = os.getcwd()
        
    json_files = [f for f in os.listdir(script_dir) if f.endswith('.json')]
    
    if not json_files:
        print("No JSON files found in the script directory.")
        save_file = input("Enter the path to your TTS save file (.json): ")
    else:
        print("Found JSON files:")
        for i, f in enumerate(json_files):
            print(f"{i+1}. {f}")
        
        selection = input("Enter the number of the TTS save file to process (or enter a different file path): ")
        
        try:
            index = int(selection) - 1
            if 0 <= index < len(json_files):
                save_file = os.path.join(script_dir, json_files[index])
            else:
                raise ValueError("Invalid selection")
        except ValueError:
            save_file = selection if os.path.exists(selection) else None
    
    if not save_file or not os.path.exists(save_file):
        print("Invalid file path.")
    else:
        # Process the TTS save file
        count = process_tts_save_file(save_file)
        print(f"Process completed. Found {count} ship card scripts.")