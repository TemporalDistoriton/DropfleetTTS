import os
import re
import json
import requests
import pandas as pd
from urllib.parse import quote

# GitHub repository information
GITHUB_REPO = "TemporalDistoriton/DropfleetTTS"
# Use raw.githubusercontent for direct file access
GITHUB_RAW_BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main"
USE_RAW_TRUE = False  # raw.githubusercontent URLs serve raw content

# Valid factions list
VALID_FACTIONS = ['UCM', 'PHR', 'Shaltari', 'Scourge', 'Resistance', 'Bioficers', 'Neutral']

# Containers to ignore
IGNORED_CONTAINERS = ["Old 2.0 Content"]

# Extraction parameters (upgrades have no Lua card images)
PARAMETERS = ['name']

# Flag to save changes
SAVE_CHANGES = True

# Logging
error_log = []
extracted_data = []


def select_mode():
    valid = {'ships', 'upgrades'}
    while True:
        choice = input("Select mode ('Ships' or 'Upgrades'): ").strip().lower()
        if choice in valid:
            print(f"Mode selected: {choice.title()}")
            return choice
        print("Invalid selection. Please enter 'Ships' or 'Upgrades'.")

MODE = select_mode()


def sanitize_filename(name):
    # Replace invalid file/path chars
    for c in '<>:"/\\|?*':
        name = name.replace(c, '_')
    return name.strip()


def get_github_image_url(faction, name, image_type=None):
    base = sanitize_filename(name)
    if MODE == 'ships':
        if not image_type:
            raise ValueError("image_type required for ships mode")
        path = f"{faction}/{base}_{image_type}.png"
    else:
        path = f"{faction}/Upgrades/{base}.png"
    return f"{GITHUB_RAW_BASE_URL}/{quote(path)}"


def check_image_exists(faction, name, image_type=None):
    url = get_github_image_url(faction, name, image_type)
    try:
        resp = requests.head(url)
        exists = resp.status_code == 200
        print(f"DEBUG: HEAD {url} -> {resp.status_code}")
        return exists, url
    except Exception as e:
        print(f"DEBUG: HEAD failed for {url}: {e}")
        return False, url


def extract_parameter(content, param):
    if param == 'name':
        m = re.search(r"(?:local\s+name\s*=\s*|name\s*=\s*)['\"](.+?)['\"]", content)
        return m.group(1).strip() if m else None
    return None


def is_ship_card_script(content):
    indicators = ["rebuildUI()","createModel","cardFrontImage","modelImage","baseScale","onSave()"]
    return sum(i in content for i in indicators) >= 3


def should_skip(path_list):
    return any(ign in path for path in path_list for ign in IGNORED_CONTAINERS)


def determine_faction(name):
    low = name.lower()
    for f in VALID_FACTIONS:
        if f.lower() in low:
            return f
    return 'Neutral'


def build_hierarchy(states, hier, parent=None, depth=0):
    if depth > 10:
        return
    for o in states:
        guid = o.get('GUID')
        if guid:
            hier[guid] = {'parent': parent, 'nickname': o.get('Nickname', '')}
            if 'ContainedObjects' in o:
                build_hierarchy(o['ContainedObjects'], hier, guid, depth + 1)


def get_container_path(hier, guid):
    path, cur = [], guid
    for _ in range(10):
        info = hier.get(cur)
        if not info or not info['parent']:
            break
        parent = hier[info['parent']]
        if parent['nickname']:
            path.append(parent['nickname'])
        cur = info['parent']
    return list(reversed(path))


def update_ships(states, hier):
    for o in states:
        lua = o.get('LuaScript','')
        if lua and is_ship_card_script(lua):
            name_val = extract_parameter(lua, 'name') or o.get('Nickname','')
            cont_path = get_container_path(hier, o.get('GUID'))
            if should_skip(cont_path):
                continue
            faction = next((determine_faction(c) for c in cont_path if determine_faction(c)!='Neutral'), 'Neutral')
            print(f"Processing ship '{name_val}' in faction '{faction}'")
            # CardFront
            ok, url = check_image_exists(faction, name_val, 'CardFrontImage')
            if ok:
                lua = re.sub(r"((?:cardFrontImage\s*=\s*['\"]))[^'\"]+(['\"])", rf"\1{url}\2", lua)
            else:
                error_log.append(f"Missing ship cardFrontImage: {name_val}")
            # ModelImage
            ok2, murl = check_image_exists(faction, name_val, 'ModelImage')
            if ok2:
                lua = re.sub(r"((?:modelImage\s*=\s*['\"]))[^'\"]+(['\"])", rf"\1{murl}\2", lua)
            else:
                error_log.append(f"Missing ship modelImage: {name_val}")
            o['LuaScript'] = lua
            extracted_data.append({'name':name_val,'faction':faction,'card_url':url,'model_url':murl})
        if 'ContainedObjects' in o:
            update_ships(o['ContainedObjects'], hier)


def update_upgrades(states, hier):
    for o in states:
        nick = o.get('Nickname','')
        if 'upgrade' in nick.lower() and 'ContainedObjects' in o:
            faction = determine_faction(nick)
            print(f"Found upgrade container: '{nick}' -> faction '{faction}'")
            for u in o['ContainedObjects']:
                if u.get('Name')=='Custom_Tile' and 'CustomImage' in u:
                    name_val = u.get('Nickname','')
                    print(f"Processing upgrade '{name_val}' in faction '{faction}'")
                    ok, url = check_image_exists(faction, name_val)
                    if ok:
                        print(f"Updating URLs to: {url}")
                        ci = u['CustomImage']
                        ci['ImageURL'] = url
                        ci['ImageSecondaryURL'] = url
                        extracted_data.append({'name':name_val,'faction':faction,'new_url':url})
                    else:
                        error_log.append(f"Missing upgrade image: {name_val}")
        if 'ContainedObjects' in o:
            update_upgrades(o['ContainedObjects'], hier)


def create_report(sdir):
    if not extracted_data:
        print("No data to write.")
        return
    df = pd.DataFrame(extracted_data)
    excel_path = os.path.join(sdir, 'UpdateData.xlsx')
    try:
        df.to_excel(excel_path, index=False)
        print(f"Excel saved: {excel_path}")
    except Exception:
        csv_path = excel_path.replace('.xlsx','.csv')
        df.to_csv(csv_path, index=False)
        print(f"CSV saved: {csv_path}")


def write_errors(sdir):
    if not error_log:
        return
    path = os.path.join(sdir, 'update_errors.log')
    with open(path, 'w') as f:
        f.write("Errors:\n")
        for e in error_log:
            f.write(e + "\n")
    print(f"Errors logged: {path}")


def process_file(path):
    print(f"Processing save file: {path}")
    sdir = os.path.dirname(os.path.abspath(path))
    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
        data = json.load(f)
    states = data.get('ObjectStates', [])
    hier = {}
    build_hierarchy(states, hier)
    if MODE == 'ships':
        update_ships(states, hier)
    else:
        update_upgrades(states, hier)
    create_report(sdir)
    write_errors(sdir)
    if SAVE_CHANGES:
        out = path.replace('.json','_modified.json')
        with open(out,'w') as f:
            json.dump(data, f, indent=2)
        print(f"Modified file written: {out}")


if __name__ == '__main__':
    try:
        import openpyxl
        print("Excel output enabled.")
    except ImportError:
        print("CSV output only.")
    json_files = [f for f in os.listdir('.') if f.endswith('.json') and not f.endswith('_modified.json')]
    if not json_files:
        path = input("Enter path to TTS save file (.json): ").strip()
    else:
        print("Found JSON files:")
        for i, f in enumerate(json_files, 1):
            print(f"{i}. {f}")
        sel = input("Select file number or enter path: ").strip()
        try:
            path = json_files[int(sel)-1]
        except Exception:
            path = sel
    process_file(path)
