import os
import re
import json
import requests
import pandas as pd
from urllib.parse import quote

# GitHub repository information
GITHUB_REPO = "TemporalDistoriton/DropfleetTTS"

# GitHub URL formats - MODIFIED for your specific repo format
GITHUB_RAW_BASE_URL = f"https://github.com/{GITHUB_REPO}/blob/main"  # Changed to blob/main
GITHUB_API_BASE_URL = f"https://api.github.com/repos/{GITHUB_REPO}/contents"

# Use raw=true for URLs
USE_RAW_TRUE = True  # Set to True to append ?raw=true to image URLs

# Valid factions list
VALID_FACTIONS = [
    'UCM',
    'PHR',
    'Shaltari',
    'Scourge',
    'Resistance',
    'Bioficers', 
    'Neutral'  # Added Neutral as a valid faction
]

# Containers to ignore
IGNORED_CONTAINERS = [
    "Old 2.0 Content"
]

# Parameters to extract for the Excel file
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

# Flag to actually save changes (set to False for testing)
SAVE_CHANGES = True

# Initialize error log
error_log = []

# List to store extracted data for Excel
extracted_data = []

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

def get_github_path(faction, ship_name, image_type):
    """
    Get the appropriate GitHub path based on the configured format
    """
    sanitized_name = sanitize_filename(ship_name)
    
    # Use faction folder format
    return f"{faction}/{sanitized_name}_{image_type}.png"

def check_image_exists(faction, ship_name, image_type):
    """
    Check if an image exists in the GitHub repository
    First checks using the API, then tries a direct HTTP request if needed
    """
    # Get the path for the image
    path = get_github_path(faction, ship_name, image_type)
    
    # URL encode the path for API request
    encoded_path = quote(f"{faction}/{sanitized_filename(ship_name)}_{image_type}.png")
    api_url = f"{GITHUB_API_BASE_URL}/{encoded_path}"
    
    try:
        # Try the API first
        response = requests.get(api_url)
        if response.status_code == 200:
            return True, path
        
        # If API fails, try a direct HTTP request to the raw URL
        raw_url = get_github_image_url(faction, ship_name, image_type)
        direct_response = requests.head(raw_url)
        if direct_response.status_code == 200:
            return True, path
        
        return False, path
    except Exception as e:
        print(f"Error checking image existence: {e}")
        return False, path

def sanitized_filename(name):
    """
    Create a URL-safe filename
    """
    # Replace spaces with %20 for URL encoding
    sanitized = sanitize_filename(name)
    return sanitized

def get_github_image_url(faction, ship_name, image_type):
    """
    Get the raw GitHub URL for an image
    """
    # Get the path for the image
    sanitized_name = sanitized_filename(ship_name)
    filename = f"{sanitized_name}_{image_type}.png"
    
    # URL encode the path components for the raw URL
    encoded_path = f"{faction}/{quote(filename)}"
    url = f"{GITHUB_RAW_BASE_URL}/{encoded_path}"
    
    # Add ?raw=true if required
    if USE_RAW_TRUE:
        url += "?raw=true"
    
    return url

def test_repository_structure(faction="Neutral", ship_name="M-Type Barge"):
    """
    Test different repository URL structures to find the correct one
    """
    print(f"Testing repository structure with {faction}/{ship_name}...")
    
    # Test different URL formats with a sample ship
    test_formats = [
        # Format 1: Standard raw githubusercontent format
        f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{faction}/{quote(sanitized_filename(ship_name))}_CardFrontImage.png",
        
        # Format 2: GitHub blob format
        f"https://github.com/{GITHUB_REPO}/blob/main/{faction}/{quote(sanitized_filename(ship_name))}_CardFrontImage.png",
        
        # Format 3: GitHub blob format with raw=true
        f"https://github.com/{GITHUB_REPO}/blob/main/{faction}/{quote(sanitized_filename(ship_name))}_CardFrontImage.png?raw=true",
        
        # Format 4: Raw GitHub format
        f"https://github.com/{GITHUB_REPO}/raw/main/{faction}/{quote(sanitized_filename(ship_name))}_CardFrontImage.png"
    ]
    
    # Try all test formats
    for i, url in enumerate(test_formats, 1):
        try:
            print(f"Testing format {i}: {url}")
            response = requests.head(url)
            if response.status_code == 200:
                print(f"SUCCESS! URL format {i} works: {url}")
                
                # Update the global constants based on which format worked
                global GITHUB_RAW_BASE_URL, USE_RAW_TRUE
                
                if i == 1:
                    GITHUB_RAW_BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main"
                    USE_RAW_TRUE = False
                elif i == 2:
                    GITHUB_RAW_BASE_URL = f"https://github.com/{GITHUB_REPO}/blob/main"
                    USE_RAW_TRUE = False
                elif i == 3:
                    GITHUB_RAW_BASE_URL = f"https://github.com/{GITHUB_REPO}/blob/main"
                    USE_RAW_TRUE = True
                elif i == 4:
                    GITHUB_RAW_BASE_URL = f"https://github.com/{GITHUB_REPO}/raw/main"
                    USE_RAW_TRUE = False
                
                return True
            else:
                print(f"Format {i} failed with status code: {response.status_code}")
        except Exception as e:
            print(f"Error testing format {i}: {e}")
    
    print("Could not detect working URL format automatically.")
    return False

def update_ship_card_script(lua_script, determined_faction, ship_name, object_name):
    """
    Update the ship card script with GitHub image URLs and correct faction
    Returns:
        - modified script
        - new model image URL (or None if not updated)
        - new card front image URL (or None if not updated)
    """
    # Check if images exist in GitHub
    card_front_exists, card_front_path = check_image_exists(determined_faction, ship_name, "CardFrontImage")
    model_exists, model_path = check_image_exists(determined_faction, ship_name, "ModelImage")
    
    # Get current image URLs
    current_card_front = extract_parameter(lua_script, 'cardFrontImage')
    current_model = extract_parameter(lua_script, 'modelImage')
    
    # Log errors if images don't exist
    if not card_front_exists:
        error_log.append(f"ERROR: CardFrontImage not found for {ship_name} in {determined_faction} faction (tried path: {card_front_path})")
    
    if not model_exists:
        error_log.append(f"ERROR: ModelImage not found for {ship_name} in {determined_faction} faction (tried path: {model_path})")
    
    # If neither image exists, don't update the script
    if not card_front_exists and not model_exists:
        error_log.append(f"SKIPPING: No images found for {ship_name} in {determined_faction} faction")
        return lua_script, None, None
    
    # Get GitHub URLs for images
    card_front_url = get_github_image_url(determined_faction, ship_name, "CardFrontImage") if card_front_exists else None
    model_url = get_github_image_url(determined_faction, ship_name, "ModelImage") if model_exists else None
    
    # Create a modified script
    modified_script = lua_script
    
    # Update the cardFrontImage URL if it exists
    new_card_front_url = None
    if card_front_exists:
        # Pattern for cardFrontImage
        card_pattern = r"((?:local\s+cardFrontImage\s*=\s*|cardFrontImage\s*=\s*)['\"])https?://[^'\"]+(['\"])"
        modified_script = re.sub(card_pattern, r"\1" + card_front_url + r"\2", modified_script)
        print(f"Updated CardFrontImage URL for {ship_name} to {card_front_url}")
        new_card_front_url = card_front_url
    
    # Update the modelImage URL if it exists
    new_model_url = None
    if model_exists:
        # Pattern for modelImage
        model_pattern = r"((?:local\s+modelImage\s*=\s*|modelImage\s*=\s*)['\"])https?://[^'\"]+(['\"])"
        modified_script = re.sub(model_pattern, r"\1" + model_url + r"\2", modified_script)
        print(f"Updated ModelImage URL for {ship_name} to {model_url}")
        new_model_url = model_url
    
    # Update the faction if necessary
    # Pattern for faction in different formats
    faction_patterns = [
        r"(local\s+faction\s*=\s*['\"]).+?(['\"])",  # local faction = "UCM"
        r"(faction\s*=\s*data\.faction\s+or\s+['\"]).+?(['\"])"  # faction = data.faction or "UCM"
    ]
    
    for pattern in faction_patterns:
        if re.search(pattern, modified_script):
            modified_script = re.sub(pattern, r"\1" + determined_faction + r"\2", modified_script)
            print(f"Updated faction to {determined_faction} for {ship_name}")
            break
    
    return modified_script, new_model_url, new_card_front_url

def should_skip_container(container_path):
    """
    Check if we should skip this container and its contents
    """
    for container_name in container_path:
        if container_name in IGNORED_CONTAINERS:
            return True
    return False

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

def process_object_states(object_states, container_hierarchy, parent_path=None, depth=0, modified=False):
    """
    Recursively process object states to find and update ship card scripts
    Returns True if any modifications were made
    """
    if depth > 10:  # Prevent infinite recursion
        return modified
    
    if parent_path is None:
        parent_path = []
    
    # Skip this branch if it's in an ignored container
    if should_skip_container(parent_path):
        print(f"Skipping content in ignored container: {' > '.join(parent_path)}")
        return modified
    
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
                
                # Extract all parameters for the Excel file
                ship_data = {}
                for param in PARAMETERS:
                    ship_data[param] = extract_parameter(lua_script, param)
                
                # Add object info
                ship_data['object_name'] = obj_nickname
                ship_data['object_guid'] = obj_guid
                ship_data['container_path'] = " > ".join(container_path)
                
                # Extract the ship name
                ship_name = ship_data['name']
                if not ship_name or ship_name == "Unknown" or ship_name == "Unnamed Model":
                    # Use the object nickname if the name is not found or generic
                    ship_name = obj_nickname
                    ship_data['name'] = ship_name
                    print(f"Using object nickname '{ship_name}' as ship name was not found in script")
                
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
                
                print(f"Determined faction: {determined_faction} for {ship_name}")
                
                # Use the container-determined faction as the main faction field
                ship_data['faction'] = determined_faction
                
                # Update the script and get new image URLs
                updated_script, new_model_url, new_card_front_url = update_ship_card_script(
                    lua_script, determined_faction, ship_name, obj_nickname)
                
                # Store the new image URLs
                ship_data['new_model_url'] = new_model_url if new_model_url else "Not Updated"
                ship_data['new_card_front_url'] = new_card_front_url if new_card_front_url else "Not Updated"
                
                # Add to the extracted data for Excel
                extracted_data.append(ship_data)
                
                # If the script was changed, update it in the object
                if updated_script != lua_script:
                    obj["LuaScript"] = updated_script
                    modified = True
        
        # Process contained objects recursively and track if modifications were made
        if "ContainedObjects" in obj and obj["ContainedObjects"]:
            child_modified = process_object_states(obj["ContainedObjects"], container_hierarchy, current_path, depth + 1, modified)
            if child_modified:
                modified = True
    
    return modified

def create_excel_file(script_dir):
    """
    Create an Excel file with the extracted data
    """
    if not extracted_data:
        print("No data to write to Excel file.")
        return
    
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(extracted_data)
    
    # Reorder columns for better readability
    columns_order = [
        'name', 'faction', 'baseScale', 'health', 'sig', 'points',
        'modelImage', 'new_model_url', 'cardFrontImage', 'new_card_front_url',
        'object_name', 'object_guid', 'container_path'
    ]
    
    # Filter to only existing columns in the DataFrame
    existing_columns = [col for col in columns_order if col in df.columns]
    df = df[existing_columns]
    
    # Create the Excel file
    excel_file = os.path.join(script_dir, "ShipCardUpdateData.xlsx")
    try:
        df.to_excel(excel_file, index=False, engine='openpyxl')
        print(f"Saved data to Excel file: {excel_file}")
    except ImportError:
        # If openpyxl is not installed, save as CSV instead
        csv_file = os.path.join(script_dir, "ShipCardUpdateData.csv")
        df.to_csv(csv_file, index=False)
        print(f"Openpyxl not installed. Saved data to CSV file instead: {csv_file}")
        print("To save as Excel, install openpyxl: pip install openpyxl")

def process_tts_save_file(save_file_path):
    """
    Process a TTS save file (JSON), update ship card scripts with GitHub URLs,
    and save the modified file
    """
    print(f"Processing TTS save file: {save_file_path}")
    script_dir = os.path.dirname(os.path.abspath(save_file_path))
    if not script_dir:
        script_dir = os.getcwd()
    
    # Test repository URL structure
    test_repository_structure()
    
    try:
        # Clear previous extracted data
        extracted_data.clear()
        
        # Read the TTS save file
        with open(save_file_path, 'r', encoding='utf-8', errors='ignore') as file:
            save_data = json.load(file)
        
        # Process the save data
        if "ObjectStates" in save_data:
            # Create a container hierarchy to track parent containers
            container_hierarchy = {}
            
            # First pass: build the container hierarchy
            build_container_hierarchy(save_data["ObjectStates"], container_hierarchy)
            
            # Second pass: process objects using the container hierarchy
            modified = process_object_states(save_data["ObjectStates"], container_hierarchy)
            
            # Create Excel file with the extracted data
            create_excel_file(script_dir)
            
            # Save the modified file if changes were made
            if modified and SAVE_CHANGES:
                # Create a backup of the original file
                backup_path = save_file_path + ".backup"
                if not os.path.exists(backup_path):
                    with open(backup_path, 'w', encoding='utf-8') as file:
                        json.dump(save_data, file, indent=2)
                    print(f"Created backup of original file: {backup_path}")
                
                # Save the modified file
                modified_path = os.path.splitext(save_file_path)[0] + "_modified.json"
                with open(modified_path, 'w', encoding='utf-8') as file:
                    json.dump(save_data, file, indent=2)
                print(f"Saved modified file: {modified_path}")
                return True
            elif modified:
                print("Changes were detected, but not saved (SAVE_CHANGES is False)")
                return True
            else:
                print("No changes were made to the file")
                return False
        else:
            print("Invalid TTS save file format. Missing 'ObjectStates' array.")
            return False
    except json.JSONDecodeError:
        print(f"Error: {save_file_path} is not a valid JSON file.")
        return False
    except Exception as e:
        print(f"Error processing {save_file_path}: {e}")
        import traceback
        traceback.print_exc()
        return False

def write_error_log():
    """
    Write error log to a file
    """
    if error_log:
        log_file = "update_errors.log"
        with open(log_file, 'w', encoding='utf-8') as file:
            file.write("UPDATE ERRORS LOG\n")
            file.write("================\n\n")
            file.write(f"Repository: {GITHUB_REPO}\n")
            file.write(f"Image path format: {GITHUB_RAW_BASE_URL}\n")
            file.write(f"Use raw=true: {USE_RAW_TRUE}\n")
            file.write("================\n\n")
            for error in error_log:
                file.write(f"{error}\n")
        print(f"Wrote {len(error_log)} errors to {log_file}")

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
        
    json_files = [f for f in os.listdir(script_dir) if f.endswith('.json') and not f.endswith('.backup') and not f.endswith('_modified.json')]
    
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
        result = process_tts_save_file(save_file)
        
        # Write error log if needed
        write_error_log()
        
        if result:
            print(f"Process completed. Check update_errors.log for any errors.")
        else:
            print(f"Process completed without making changes.")