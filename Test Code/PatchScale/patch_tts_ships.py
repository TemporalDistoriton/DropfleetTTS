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
    print(f"Loading Lua file: {file_path}")
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()

    ships_found = 0
    total_patches = 0
    outer_replacements = []

    for outer_match in LUA_LONG_STRING.finditer(content):
        outer_eq = outer_match.group(1)
        outer_text = outer_match.group(2).strip()

        try:
            spawner_obj = json.loads(outer_text)
        except (json.JSONDecodeError, ValueError):
            continue
        if not isinstance(spawner_obj, dict):
            continue

        spawner_lua = spawner_obj.get("LuaScript", "")
        spawner_modified = False

        # Check spawner itself (rare but possible)
        if spawner_lua and is_ship_script(spawner_lua):
            ships_found += 1
            name = spawner_obj.get("Nickname", "direct")
            print(f"\n[Ship #{ships_found}] {name} (direct in spawner)")
            patched, n = patch_lua_script(spawner_lua)
            if n > 0:
                spawner_obj["LuaScript"] = patched
                spawner_modified = True
                total_patches += n
                for fn, _, _ in PATCHES[:n]:
                    print(f"  \u2713 {fn}")
            else:
                print(f"  (already correct)")

        # Find objectJSON = [=*[{ship JSON}]=*] inside spawner's LuaScript
        if spawner_lua and ("objectJSON" in spawner_lua or "[=[" in spawner_lua):
            inner_replacements = []

            for inner_match in LUA_LONG_STRING.finditer(spawner_lua):
                inner_eq = inner_match.group(1)
                inner_text = inner_match.group(2).strip()

                try:
                    ship_obj = json.loads(inner_text)
                except (json.JSONDecodeError, ValueError):
                    continue
                if not isinstance(ship_obj, dict):
                    continue

                ship_lua = ship_obj.get("LuaScript", "")
                if not ship_lua or not is_ship_script(ship_lua):
                    continue

                ships_found += 1
                name = ship_obj.get("Nickname", f"ship #{ships_found}")
                print(f"\n[Ship #{ships_found}] {name}")

                patched, n = patch_lua_script(ship_lua)
                if n > 0:
                    ship_obj["LuaScript"] = patched
                    new_json = json.dumps(ship_obj, ensure_ascii=False, separators=(', ', ': '))
                    old_block = inner_match.group(0)
                    new_block = f"[{inner_eq}[{new_json}]{inner_eq}]"
                    inner_replacements.append((old_block, new_block))
                    total_patches += n
                    for fn, _, _ in PATCHES[:n]:
                        print(f"  \u2713 {fn}")
                else:
                    print(f"  (already correct)")

            if inner_replacements:
                modified_lua = spawner_lua
                for old_b, new_b in inner_replacements:
                    modified_lua = modified_lua.replace(old_b, new_b, 1)
                spawner_obj["LuaScript"] = modified_lua
                spawner_modified = True

        if spawner_modified:
            new_outer_json = json.dumps(spawner_obj, ensure_ascii=False, separators=(', ', ': '))
            old_outer_block = outer_match.group(0)
            new_outer_block = f"[{outer_eq}[{new_outer_json}]{outer_eq}]"
            outer_replacements.append((old_outer_block, new_outer_block))

    for old_b, new_b in outer_replacements:
        content = content.replace(old_b, new_b, 1)

    return content, ships_found, total_patches


# ======= TTS JSON SAVE ===================================================

def patch_embedded_json_in_lua(lua_script, depth=0):
    ships = 0
    patches = 0
    inner_replacements = []

    for match in LUA_LONG_STRING.finditer(lua_script):
        eq = match.group(1)
        json_str = match.group(2).strip()
        try:
            obj = json.loads(json_str)
        except (json.JSONDecodeError, ValueError):
            continue
        if not isinstance(obj, dict):
            continue

        inner_lua = obj.get("LuaScript", "")
        if inner_lua and is_ship_script(inner_lua):
            ships += 1
            name = obj.get("Nickname", obj.get("Name", "embedded"))
            patched, n = patch_lua_script(inner_lua)
            if n > 0:
                print(f"{'  ' * depth}  [Embedded] {name}: {n} patches")
                obj["LuaScript"] = patched
                new_json = json.dumps(obj, ensure_ascii=False, separators=(', ', ': '))
                old_block = match.group(0)
                new_block = f"[{eq}[{new_json}]{eq}]"
                inner_replacements.append((old_block, new_block))
                patches += n
            else:
                print(f"{'  ' * depth}  [Embedded] {name}: already correct")

    for old_b, new_b in inner_replacements:
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
