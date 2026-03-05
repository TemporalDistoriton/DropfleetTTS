import os
import re
import json
import requests
import pandas as pd
from urllib.parse import quote
import argparse  # Added for command line arguments

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

# Parameters for upgrades
UPGRADE_PARAMETERS = [
    'name',
    'cardImage',  # Common parameter for upgrade cards
    'points'
]

# Script modes
MODE_SHIPS_ONLY = "ships"
MODE_UPGRADES_ONLY = "upgrades"
MODE_BOTH = "both"

# Flag to actually save changes (set to False for testing)
SAVE_CHANGES = True

# Initialize error log
error_log = []

# Lists to store extracted data for Excel
extracted_data = []
extracted_upgrade_data = []

def extract_parameter(content, param_name):
    """
    Extract parameter value from the lua script
    """
    # Different patterns based on the parameter type
    if param_name in ['baseScale', 'health', 'sig', 'points']:
        # Number parameters (may have local or direct assignment)
        pattern = rf'(?:local\s+{param_name}\s*=\s*|{param_name}\s*=\s*)([0-9.]+)'
    elif param_name in ['modelImage', 'cardFrontImage', 'cardImage']:
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

def is_upgrade_card_script(content):
    """
    Check if the content contains an upgrade card script based on key indicators
    """
    # Look for specific indicators in upgrade card scripts
    indicators = [
        "rebuildUI()",
        "cardImage",
        "onLoad",
        "points",
        "onSave()"
    ]
    
    # Count how many indicators are present
    matches = sum(1 for indicator in indicators if indicator in content)
    
    # If upgrade has cardImage but no modelImage, it's probably an upgrade
    has_card_image = "cardImage" in content or "CardImage" in content
    no_model_image = "modelImage" not in content and "ModelImage" not in content
    
    # If at least 3 indicators are found AND it has cardImage but no modelImage, consider it an upgrade card script
    return (matches >= 3 and has_card_image and no_model_image)

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

def get_github_path(faction, item_name, image_type=None, is_upgrade=False):
    """
    Get the appropriate GitHub path based on the configured format
    """
    sanitized_name = sanitize_filename(item_name)
    
    if is_upgrade:
        # For upgrades, use the Upgrades subfolder
        return f"{faction}/Upgrades/{sanitized_name}.png"
    else:
        # For ships, use previous format with image type
        return f"{faction}/{sanitized_name}_{image_type}.png"

def check_image_exists(faction, item_name, image_type=None, is_upgrade=False):
    """
    Check if an image exists in the GitHub repository
    First checks using the API, then tries a direct HTTP request if needed
    """
    # Get the path for the image
    path = get_github_path(faction, item_name, image_type, is_upgrade)
    
    # URL encode the path for API request
    encoded_path = quote(path)
    api_url = f"{GITHUB_API_BASE_URL}/{encoded_path}"
    
    try:
        # Try the API first
        response = requests.get(api_url)
        if response.status_code == 200:
            return True, path
        
        # If API fails, try a direct HTTP request to the raw URL
        raw_url = get_github_image_url(faction, item_name, image_type, is_upgrade)
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

def get_github_image_url(faction, item_name, image_type=None, is_upgrade=False):
    """
    Get the raw GitHub URL for an image
    """
    # Get the path for the image
    sanitized_name = sanitized_filename(item_name)
    
    if is_upgrade:
        # For upgrades
        filename = f"{sanitized_name}.png"
        # URL encode the path components for the raw URL
        encoded_path = f"{faction}/Upgrades/{quote(filename)}"
    else:
        # For ships
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

def test_upgrade_structure():
    """
    Test the upgrade repository structure with the example URL provided
    """
    print("Testing upgrade repository structure...")
    
    # Test the example URL provided
    test_url = "https://github.com/TemporalDistoriton/DropfleetTTS/blob/main/UCM/Upgrades/Cobra%20Heavy%20Laser%20Pair.png?raw=true"
    
    try:
        response = requests.head(test_url)
        if response.status_code == 200:
            print(f"SUCCESS! Upgrade URL format works: {test_url}")
            
            # Set the global constants based on the format
            global GITHUB_RAW_BASE_URL, USE_RAW_TRUE
            GITHUB_RAW_BASE_URL = f"https://github.com/{GITHUB_REPO}/blob/main"
            USE_RAW_TRUE = True
            
            return True
        else:
            print(f"Upgrade URL format failed with status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"Error testing upgrade URL format: {e}")
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

def update_upgrade_card_script(lua_script, determined_faction, upgrade_name, object_name):
    """
    Update the upgrade card script with GitHub image URLs
    Returns:
        - modified script
        - new card image URL (or None if not updated)
    """
    # Check if upgrade image exists in GitHub
    upgrade_exists, upgrade_path = check_image_exists(determined_faction, upgrade_name, is_upgrade=True)
    
    # Get current image URL
    current_card_image = extract_parameter(lua_script, 'cardImage')
    
    # Log errors if image doesn't exist
    if not upgrade_exists:
        error_log.append(f"ERROR: Upgrade image not found for {upgrade_name} in {determined_faction} faction (tried path: {upgrade_path})")
        return lua_script, None
    
    # Get GitHub URL for image
    upgrade_url = get_github_image_url(determined_faction, upgrade_name, is_upgrade=True)
    
    # Create a modified script
    modified_script = lua_script
    
    # Update the cardImage URL
    card_pattern = r"((?:local\s+cardImage\s*=\s*|cardImage\s*=\s*)['\"])https?://[^'\"]+(['\"])"
    modified_script = re.sub(card_pattern, r"\1" + upgrade_url + r"\2", modified_script)
    print(f"Updated Upgrade image URL for {upgrade_name} to {upgrade_url}")
    
    # Update the faction if necessary
    faction_patterns = [
        r"(local\s+faction\s*=\s*['\"]).+?(['\"])",  # local faction = "UCM"
        r"(faction\s*=\s*data\.faction\s+or\s+['\"]).+?(['\"])"  # faction = data.faction or "UCM"
    ]
    
    for pattern in faction_patterns:
        if re.search(pattern, modified_script):
            modified_script = re.sub(pattern, r"\1" + determined_faction + r"\2", modified_script)
            print(f"Updated faction to {determined_faction} for upgrade {upgrade_name}")
            break
    
    return modified_script, upgrade_url

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

def process_object_states(object_states, container_hierarchy, parent_path=None, depth=0, modified=False, processing_mode=MODE_BOTH):
    """
    Recursively process object states to find and update ship card scripts and upgrade cards
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
            
            # Find container path for this object
            container_path = find_container_path(container_hierarchy, obj_guid)
            
            # Skip this object if it's in an ignored container
            if should_skip_container(container_path):
                print(f"Skipping object in ignored container: {' > '.join(container_path)}")
                continue
            
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
            
            # Process ship cards
            if (processing_mode == MODE_SHIPS_ONLY or processing_mode == MODE_BOTH) and is_ship_card_script(lua_script):
                print(f"Found ship card script in object: {obj_nickname}")
                
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
                
                print(f"Determined faction: {determined_faction} for ship {ship_name}")
                
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
            
            # Process upgrade cards
            elif (processing_mode == MODE_UPGRADES_ONLY or processing_mode == MODE_BOTH) and is_upgrade_card_script(lua_script):
                print(f"Found upgrade card script in object: {obj_nickname}")
                
                # Extract parameters for the Excel file
                upgrade_data = {}
                for param in UPGRADE_PARAMETERS:
                    upgrade_data[param] = extract_parameter(lua_script, param)
                
                # Add object info
                upgrade_data['object_name'] = obj_nickname
                upgrade_data['object_guid'] = obj_guid
                upgrade_data['container_path'] = " > ".join(container_path)
                
                # Extract the upgrade name
                upgrade_name = upgrade_data['name']
                if not upgrade_name or upgrade_name == "Unknown":
                    # Use the object nickname if the name is not found
                    upgrade_name = obj_nickname
                    upgrade_data['name'] = upgrade_name
                    print(f"Using object nickname '{upgrade_name}' as upgrade name was not found in script")
                
                print(f"Determined faction: {determined_faction} for upgrade {upgrade_name}")
                
                # Use the container-determined faction
                upgrade_data['faction'] = determined_faction
                
                # Update the script and get new image URL
                updated_script, new_card_url = update_upgrade_card_script(
                    lua_script, determined_faction, upgrade_name, obj_nickname)
                
                # Store the new image URL
                upgrade_data['new_card_url'] = new_card_url if new_card_url else "Not Updated"
                
                # Add to the extracted upgrade data for Excel
                extracted_upgrade_data.append(upgrade_data)
                
                # If the script was changed, update it in the object
                if updated_script != lua_script:
                    obj["LuaScript"] = updated_script
                    modified = True
        
        # Process contained objects recursively and track if modifications were made
        if "ContainedObjects" in obj and obj["ContainedObjects"]:
            child_modified = process_object_states(
                obj["ContainedObjects"], 
                container_hierarchy, 
                current_path, 
                depth + 1, 
                modified,
                processing_mode
            )
            if child_modified:
                modified = True
    
    return modified

def create_excel_file(script_dir, mode):
    """
    Create Excel file(s) with the extracted data
    """
    # Create ships Excel file if needed
    if (mode == MODE_SHIPS_ONLY or mode == MODE_BOTH) and extracted_data:
        # Create a DataFrame from the extracted data
        df_ships = pd.DataFrame(extracted_data)
        
        # Reorder columns for better readability
        ship_columns_order = [
            'name', 'faction', 'baseScale', 'health', 'sig', 'points',
            'modelImage', 'new_model_url', 'cardFrontImage', 'new_card_front_url',
            'object_name', 'object_guid', 'container_path'
        ]
        
        # Filter to only existing columns in the DataFrame
        existing_ship_columns = [col for col in ship_columns_order if col in df_ships.columns]
        df_ships = df_ships[existing_ship_columns]
        
        # Create the Excel file
        ships_excel_file = os.path.join(script_dir, "ShipCardUpdateData.xlsx")
        try:
            df_ships.to_excel(ships_excel_file, index=False, engine='openpyxl')
            print(f"Saved ship data to Excel file: {ships_excel_file}")
        except ImportError:
            # If openpyxl is not installed, save as CSV instead
            ships_csv_file = os.path.join(script_dir, "ShipCardUpdateData.csv")
            df_ships.to_csv(ships_csv_file, index=False)
            print(f"Openpyxl not installed. Saved ship data to CSV file instead: {ships_csv_file}")
    
    # Create upgrades Excel file if needed
    if (mode == MODE_UPGRADES_ONLY or mode == MODE_BOTH) and extracted_upgrade_data:
        # Create a DataFrame from the extracted upgrade data
        df_upgrades = pd.DataFrame(extracted_upgrade_data)
        
        # Reorder columns for better readability
        upgrade_columns_order = [
            'name', 'faction', 'points',
            'cardImage', 'new_card_url',
            'object_name', 'object_guid', 'container_path'
        ]
        
        # Filter to only existing columns in the DataFrame
        existing_upgrade_columns = [col for col in upgrade_columns_order if col in df_upgrades.columns]
        df_upgrades = df_upgrades[existing_upgrade_columns]
        
        # Create the Excel file
        upgrades_excel_file = os.path.join(script_dir, "UpgradeCardUpdateData.xlsx")
        try:
            df_upgrades.to_excel(upgrades_excel_file, index=False, engine='openpyxl')
            print(f"Saved upgrade data to Excel file: {upgrades_excel_file}")
        except ImportError:
            # If openpyxl is not installed, save as CSV instead
            upgrades_csv_file = os.path.join(script_dir, "UpgradeCardUpdateData.csv")
            df_upgrades.to_csv(upgrades_csv_file, index=False)
            print(f"Openpyxl not installed. Saved upgrade data to CSV file instead: {upgrades_csv_file}")
            print("To save as Excel, install openpyxl: pip install openpyxl")