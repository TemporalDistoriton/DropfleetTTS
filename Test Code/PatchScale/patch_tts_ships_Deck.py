#!/usr/bin/env python3
"""
TTS Ship Script Patcher - Scale Compensation Fix
=================================================
Patches generateScanLines, generateSignatureLines, and
generateFiringArcLines to divide by getScaleFromBaseSize(ShipbaseSize).

Supports:
  - TTS save files (.json) - walks ObjectStates/ContainedObjects/States
  - Lua ship data files (.lua) - finds ship JSON inside savedShipCards
    Structure: savedShipCards -> [====[spawner JSON]====]
               spawner LuaScript -> objectJSON = [=[ship JSON]=]
               ship LuaScript -> actual ship code to patch

Usage:
    python patch_tts_ships.py <save_file.json or data_file.lua>
"""

import json
import sys
import re
import shutil
from pathlib import Path

SCALE_LINE = "local scale = getScaleFromBaseSize(ShipbaseSize)"

PATCHES = [
    (
        "generateScanLines",
        re.compile(r'([ \t]*)(local scanRadius = ShipScan \* UNIT_SCALE)\b(?!\s*/\s*scale)'),
        lambda m: f"{m.group(1)}{SCALE_LINE}\n{m.group(1)}{m.group(2)} / scale",
    ),
    (
        "generateSignatureLines",
        re.compile(r'([ \t]*)(local sigRadius = state\.sig \* UNIT_SCALE)\b(?!\s*/\s*scale)'),
        lambda m: f"{m.group(1)}{SCALE_LINE}\n{m.group(1)}{m.group(2)} / scale",
    ),
    (
        "generateFiringArcLines",
        re.compile(r'([ \t]*)(local arcLineLength = 16 \* UNIT_SCALE)\b(?!\s*/\s*scale)'),
        lambda m: f"{m.group(1)}{SCALE_LINE}\n{m.group(1)}{m.group(2)} / scale",
    ),
]

CLEANUP_PATTERNS = [
    re.compile(r'[ \t]*local scale = self\.getScale\(\)\.x[^\n]*\n'),
    re.compile(r'[ \t]*local scale = getScaleFromBaseSize\(ShipbaseSize\)[^\n]*\n'),
    re.compile(r'[ \t]*local _ok, _sv = pcall\(function\(\) return self\.getScale\(\) end\)[^\n]*\n'),
    re.compile(r'[ \t]*local scale = \(_ok and _sv\) and _sv\.x or getScaleFromBaseSize\(ShipbaseSize\)[^\n]*\n'),
]

REVERT_PATTERNS = [
    (re.compile(r'(local scanRadius = ShipScan \* UNIT_SCALE)\s*/\s*scale'), r'\1'),
    (re.compile(r'(local sigRadius = state\.sig \* UNIT_SCALE)\s*/\s*scale'), r'\1'),
    (re.compile(r'(local arcLineLength = 16 \* UNIT_SCALE)\s*/\s*scale'), r'\1'),
]

LUA_LONG_STRING = re.compile(r'\[([=]*)\[(.*?)\]\1\]', re.DOTALL)


def is_ship_script(lua):
    """Check if this is an actual ship script (not a spawner containing embedded ship JSON)."""
    # Must have all key functions
    if not all(x in lua for x in [
        "generateScanLines", "generateSignatureLines",
        "generateFiringArcLines", "ShipScan"
    ]):
        return False
    # If it contains objectJSON, it's a spawner - the ship code is embedded, not direct
    if "objectJSON" in lua:
        return False
    return True


def is_already_patched(lua):
    return bool(
        re.search(r'local scale = getScaleFromBaseSize\(ShipbaseSize\)\s*\n\s*local scanRadius = ShipScan \* UNIT_SCALE / scale', lua) and
        re.search(r'local scale = getScaleFromBaseSize\(ShipbaseSize\)\s*\n\s*local sigRadius = state\.sig \* UNIT_SCALE / scale', lua) and
        re.search(r'local scale = getScaleFromBaseSize\(ShipbaseSize\)\s*\n\s*local arcLineLength = 16 \* UNIT_SCALE / scale', lua)
    )


def patch_lua_script(lua):
    had_crlf = '\r\n' in lua
    lua = lua.replace('\r\n', '\n')

    if is_already_patched(lua):
        return (lua.replace('\n', '\r\n') if had_crlf else lua), 0

    for p in CLEANUP_PATTERNS:
        lua = p.sub('', lua)
    for p, r in REVERT_PATTERNS:
        lua = p.sub(r, lua)
    while '\n\n\n' in lua:
        lua = lua.replace('\n\n\n', '\n\n')

    count = 0
    for fn_name, pattern, repl_fn in PATCHES:
        new_lua, n = pattern.subn(repl_fn, lua, count=1)
        if n > 0:
            lua = new_lua
            count += 1

    if had_crlf:
        lua = lua.replace('\n', '\r\n')
    return lua, count


# ======= LUA FILE (savedShipCards) =======================================

def process_lua_file(file_path):
    """Process a Lua file containing savedShipCards with embedded JSON.
    Uses the same recursive patcher as JSON saves."""
    print(f"Loading Lua file: {file_path}")
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    new_content, ships, patches = patch_embedded_json_in_lua(content, depth=0)
    return new_content, ships, patches


# ======= TTS JSON SAVE ===================================================

def patch_embedded_json_in_lua(lua_script, depth=0):
    """Recursively find and patch ship scripts inside Lua long strings.

    Handles arbitrary nesting:
      faction tile LuaScript -> [====[spawner JSON]====]
        spawner LuaScript -> objectJSON = [=[ship JSON]=]
          ship LuaScript -> actual code to patch
    """
    ships = 0
    patches = 0
    replacements = []

    for match in LUA_LONG_STRING.finditer(lua_script):
        eq = match.group(1)
        json_str = match.group(2).strip()
        if not json_str.startswith('{'):
            continue
        try:
            obj = json.loads(json_str)
        except (json.JSONDecodeError, ValueError):
            continue
        if not isinstance(obj, dict):
            continue

        inner_lua = obj.get("LuaScript", "")
        if not inner_lua:
            continue

        obj_modified = False
        name = obj.get("Nickname", obj.get("Name", "embedded"))

        # Case 1: This IS a ship script - patch it directly
        if is_ship_script(inner_lua):
            ships += 1
            patched, n = patch_lua_script(inner_lua)
            if n > 0:
                print(f"{'  ' * depth}[Ship #{ships}] {name}")
                for fn, _, _ in PATCHES[:n]:
                    print(f"{'  ' * depth}  \u2713 {fn}")
                obj["LuaScript"] = patched
                obj_modified = True
                patches += n
            else:
                print(f"{'  ' * depth}[Ship] {name}: already correct")

        # Case 2: Contains more embedded JSON (spawner with objectJSON, etc.)
        elif '[=[' in inner_lua or '[====[' in inner_lua:
            new_inner, s, p = patch_embedded_json_in_lua(inner_lua, depth + 1)
            if p > 0:
                obj["LuaScript"] = new_inner
                obj_modified = True
            ships += s
            patches += p

        if obj_modified:
            new_json = json.dumps(obj, ensure_ascii=False, separators=(', ', ': '))
            old_block = match.group(0)
            new_block = f"[{eq}[{new_json}]{eq}]"
            replacements.append((old_block, new_block))

    for old_b, new_b in replacements:
        lua_script = lua_script.replace(old_b, new_b, 1)

    return lua_script, ships, patches


def walk_objects(obj_list, depth=0):
    ships = 0
    patches = 0
    for obj in (obj_list or []):
        if not isinstance(obj, dict):
            continue

        lua = obj.get("LuaScript", "")
        name = obj.get("Nickname", obj.get("Name", "unnamed"))

        if lua and is_ship_script(lua):
            ships += 1
            print(f"\n{'  ' * depth}[Ship #{ships}] {name}")
            patched, n = patch_lua_script(lua)
            if n > 0:
                obj["LuaScript"] = patched
                patches += n
                for fn, _, _ in PATCHES[:n]:
                    print(f"  {'  ' * depth}\u2713 {fn}")
            else:
                print(f"  {'  ' * depth}(already correct)")

        if lua and ('objectJSON' in lua or '[=[' in lua or '[====[' in lua):
            new_lua, s, p = patch_embedded_json_in_lua(lua, depth)
            if p > 0:
                obj["LuaScript"] = new_lua
                ships += s
                patches += p
            elif s > 0:
                ships += s

        s, p = walk_objects(obj.get("ContainedObjects", []), depth + 1)
        ships += s
        patches += p

        for sk, sv in (obj.get("States") or {}).items():
            if not isinstance(sv, dict):
                continue
            sl = sv.get("LuaScript", "")
            if sl and is_ship_script(sl):
                ships += 1
                patched, n = patch_lua_script(sl)
                if n > 0:
                    sv["LuaScript"] = patched
                    patches += n
            if sl and ('objectJSON' in sl or '[=[' in sl or '[====[' in sl):
                new_sl, s2, p2 = patch_embedded_json_in_lua(sl, depth + 1)
                if p2 > 0:
                    sv["LuaScript"] = new_sl
                    ships += s2
                    patches += p2
            s2, p2 = walk_objects(sv.get("ContainedObjects", []), depth + 1)
            ships += s2
            patches += p2

    return ships, patches


def process_json_save(save_path):
    print(f"Loading JSON save: {save_path}")
    with open(save_path, "r", encoding="utf-8") as f:
        save_data = json.load(f)
    ships, patches = walk_objects(save_data.get("ObjectStates", []))
    return save_data, ships, patches


# ======= MAIN ============================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python patch_tts_ships.py <save_file.json or data_file.lua>")
        sys.exit(1)

    file_path = Path(sys.argv[1])
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)

    backup_path = file_path.with_name(file_path.stem + ".backup" + file_path.suffix)
    print(f"Creating backup: {backup_path}")
    shutil.copy2(file_path, backup_path)

    print("\n" + "=" * 60)
    print("Scanning for ship scripts...")
    print("=" * 60)

    is_lua = file_path.suffix.lower() == '.lua'

    if is_lua:
        content, ships, patches = process_lua_file(file_path)
    else:
        save_data, ships, patches = process_json_save(file_path)

    print("\n" + "=" * 60)
    print(f"Ships found:     {ships}")
    print(f"Patches applied: {patches}")
    print("=" * 60)

    if patches > 0:
        print(f"\nWriting patched file: {file_path}")
        with open(file_path, "w", encoding="utf-8") as f:
            if is_lua:
                f.write(content)
            else:
                json.dump(save_data, f, ensure_ascii=False)
        print(f"Done! Backup at: {backup_path}")
    else:
        if ships == 0:
            print("\nWARNING: No ship scripts found!")
        else:
            print("\nNo changes needed - all ships already patched.")
        backup_path.unlink()


if __name__ == "__main__":
    main()
