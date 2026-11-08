#!/usr/bin/env python3
"""Generate a Tabletop Simulator Upgrades & Admirals object from Excel.

Inputs (by default in the same directory as this script):
    Upgrades.xlsx                  - Excel containing Sheet1 / Name, Effect, Ships,
                                     Points, Faction (other sheets are ignored)
    Upgrades  Admirals.json        - Existing TTS saved-object template

Outputs:
    Upgrades_Admirals_Generated.json - importable TTS saved object
    Upgrades_TTS_Manifest.csv        - one line per generated upgrade; image URL,
                                       stat modifiers, source template, etc.

The 30 existing Admiral cards are retained by default. Old upgrade cards are
replaced using their existing tile objects/scripts whenever a match is found.
The picker's Lua UI is preserved; only the embedded savedUpgradeCards table is
reconstructed. New cards clone an existing tile and its Lua apply-button logic.

This script intentionally does not download images or publish to GitHub. It
constructs the raw GitHub image URLs, which become live when you upload PNGs.

Usage:
    pip install openpyxl
    python GenerateUpgradeTTS.py
    python GenerateUpgradeTTS.py --excel "Dropfleet_Upgrades_Audited_2026-11-08.xlsx"
    python GenerateUpgradeTTS.py --overrides upgrade_overrides.json

Optional Excel columns: Thrust, Scan, Admiral. These override extracted stats.
An optional JSON overrides file can set any combination of these four fields,
using keys such as "Resistance|Cruiser Drive Refit" or
"Scourge|Minelayer|Kulshedra"; example: {"...": {"thrust": 3}}.
"""

from __future__ import annotations

import argparse
import copy
import csv
import hashlib
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path
from urllib.parse import quote

try:
    from openpyxl import load_workbook
except ImportError as exc:
    raise SystemExit("Missing dependency. Install with: pip install openpyxl") from exc


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_GITHUB_RAW_BASE = (
    "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main"
)
FACTION_MAP = {
    "ucm": "UCM", "united colonies of mankind": "UCM",
    "phr": "PHR", "post human republic": "PHR", "post-human republic": "PHR",
    "scourge": "Scourge", "shaltari": "Shaltari",
    "resistance": "Resistance", "bioficer": "Bioficer", "bioficers": "Bioficer",
    "civilian": "Civilian",
}
REQUIRED_COLUMNS = ("Name", "Effect", "Ships", "Points", "Faction")
DATA_FIELDS = ("thrust", "points", "scan", "admiral")
EMBED_MARKER = "--EMBEDDED_DATA_START"
EMBED_ENTRY = re.compile(
    r"\{\s*name\s*=\s*\[====\[(.*?)\]====\],\s*"
    r"imageURL\s*=\s*\[====\[(.*?)\]====\],\s*"
    r"faction\s*=\s*\[====\[(.*?)\]====\],\s*"
    r"json\s*=\s*\[====\[(.*?)\]====\],\s*\},",
    re.DOTALL,
)

# Compatibility aliases from the previous picker to the new spreadsheet titles.
# Keys are (faction, exact Excel Name), values are old card Nicknames. These
# allow the generator to reuse custom card Lua rather than silently losing it.
OLD_NAME_ALIASES = {
    ("UCM", "Drive Refit"): "UCM Drive Refit",
    ("Scourge", "Keel Refit (Cloaking Keel)"): "Scourge Keel Refit Cloaking Keel",
    ("Scourge", "Keel Refit (Engine Upgrade)"): "Scourge Keel Refit Engine Upgrade",
    ("Scourge", "Keel Options (Cloaking Keel)"): "Scourge Keel Options Cloaking Keel",
    ("Scourge", "Keel Options (Engine Upgrade)"): "Scourge Keel Options Engine Upgrade",
    ("Resistance", "10K Mass Driver Turret"): "Dreadnought 10K Mass Driver Turret",
    ("Resistance", "Cruiser Bulk Landers & Fire Ships"): "Cruiser Bulk Landers And Fire Ships",
    ("Resistance", "Cruiser Fighters & Bombers"): "Cruiser Fighters And Bombers",
    ("Resistance", "N-31 Hybrid Gun Bank"): "Cruiser N-31 Hybrid Gun Bank",
    ("Resistance", "Dreadnought Bulk Landers Fire Ships"): "Dreadnought Bulk Landers And Fire Ships",
    ("Resistance", "Dreadnought Fighters & Bombers"): "Dreadnought Fighters And Bombers",
    ("Resistance", "Battleship Engine Refit"): "Grand Battleship Engine Refit",
    ("Resistance", "Heavy Dual Vent Cannon Turret"): "Dreadnought Heavy Dual Vent Cannon Turret",
    ("Resistance", "Light Vent Cannon Turret"): "Frigate Light Vent Cannon Turret",
    ("Resistance", "N-8 Artillery Cannon Bank"): "Cruiser N-8 Artillery Cannon Bank",
    ("Resistance", "N-11 Artillery Cannon Turret"): "Cruiser N-11 Artillery Cannon Turret",
    ("Resistance", "N-31 Hybrid Gun Battery"): "Dreadnought NC-31 Hybrid Gun Battery",
    ("Resistance", "N-31 Hybrid Gun Turret"): "Frigate N-31 Hybrid Gun Turret",
    ("Resistance", "N-109 Bombardment Mortar Turret"): "Cruiser N-109 Bombardment Mortar Turret",
    ("Resistance", "NC-16 Missile Bank"): "Cruiser NC-16 Missile Bank",
    ("Resistance", "NC-16 Missile Battery"): "Dreadnought NC-16 Missile Battery",
    ("Resistance", "NC-16 Missile Turret"): "Frigate NC-16 Missile Turret",
    ("Resistance", "Sensor Dome"): "Frigate Sensor Dome",
    ("Resistance", "Vent Cannon Turret"): "Cruiser Vent Cannon Turret",
    ("Resistance", "XN-31 Mass Driver Turret"): "Cruiser XN-31 Mass Driver Turret",
    ("Shaltari", "Drive Augmentation"): "Shaltari Drive Augmentation",
}


def sanitise_filename(name: str) -> str:
    """Exactly matches the sanitise_filename convention in your UpgradeGen.py."""
    keep = set(" abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-.")
    return "".join(c if c in keep else "_" for c in name).strip()


def unique_image_name(name: str, faction: str, ships: str) -> str:
    """The image generator uses this same convention, including Minelayers."""
    if name.strip().casefold() == "minelayer":
        first_ship = ships.split(",", 1)[0].strip()
        if not first_ship:
            raise ValueError(f"{faction} Minelayer must specify Ships for a unique image")
        return f"Minelayer - {first_ship}"
    return name.strip()


def image_filename(name: str, faction: str, ships: str) -> str:
    return sanitise_filename(unique_image_name(name, faction, ships)) + ".png"


def normalize_faction(raw: object) -> str:
    key = str(raw or "").strip().casefold()
    if key not in FACTION_MAP:
        raise ValueError(f"Unrecognized faction {raw!r}")
    return FACTION_MAP[key]


def norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text)).casefold()
    text = text.replace("&", " and ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def select_excel(path: str | None) -> Path:
    if path:
        selected = Path(path)
        if not selected.is_absolute():
            selected = SCRIPT_DIR / selected
        if not selected.is_file():
            raise FileNotFoundError(f"Excel workbook not found: {selected}")
        return selected
    for candidate in ("Upgrades.xlsx", "Dropfleet_Upgrades_Audited_2026-11-08.xlsx", "Upgrades - Full.xlsx"):
        selected = SCRIPT_DIR / candidate
        if selected.is_file():
            return selected
    raise FileNotFoundError("No workbook found next to script (expected Upgrades.xlsx).")


def read_upgrade_rows(excel_path: Path, sheet_name: str = "Sheet1") -> list[dict]:
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Worksheet {sheet_name!r} not found. Sheets: {wb.sheetnames}")
        ws = wb[sheet_name]
        it = ws.iter_rows(values_only=True)
        raw_headers = next(it)
        headers = {str(value).strip().casefold(): i for i, value in enumerate(raw_headers) if value is not None}
        missing = [name for name in REQUIRED_COLUMNS if name.casefold() not in headers]
        if missing:
            raise ValueError(f"Worksheet {sheet_name} missing columns: {missing}")
        def get(row: tuple, key: str):
            i = headers.get(key.casefold())
            return row[i] if i is not None and i < len(row) else None
        rows = []
        for row_no, values in enumerate(it, start=2):
            if all(v is None for v in values):
                continue
            name = str(get(values, "Name") or "").strip()
            if not name:
                raise ValueError(f"Excel row {row_no} is populated but Name is empty")
            faction = normalize_faction(get(values, "Faction"))
            points = nonnegative_int(get(values, "Points"), f"Excel row {row_no} Points")
            item = {
                "row": row_no, "name": name, "faction": faction,
                "points": points,
                "effect": str(get(values, "Effect") or "").strip(),
                "ships": str(get(values, "Ships") or "").strip(),
                "excel_mods": {},
            }
            for stat in ("thrust", "scan", "admiral"):
                raw = get(values, stat)
                if raw is not None and str(raw).strip() != "":
                    item["excel_mods"][stat] = nonnegative_int(raw, f"Excel row {row_no} {stat}")
            rows.append(item)
    finally:
        wb.close()
    if not rows:
        raise ValueError("No upgrade records found in the Excel worksheet")
    unique = set()
    for item in rows:
        k = (item["faction"].casefold(), norm(item["name"]), norm(item["ships"]))
        if k in unique:
            raise ValueError(f"Duplicate row/upgrade combination: {item}")
        unique.add(k)
    return rows


def nonnegative_int(raw: object, label: str) -> int:
    if isinstance(raw, bool):
        raise ValueError(f"{label}: expected a number, got boolean")
    try:
        value = float(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label}: invalid value {raw!r}") from exc
    if not value.is_integer() or value < 0:
        raise ValueError(f"{label}: must be a nonnegative integer, got {raw!r}")
    return int(value)


def extract_saved_cards(lua: str) -> tuple[str, list[dict]]:
    marker = re.search(r"(?m)^--EMBEDDED_DATA_START\s*$", lua)
    if not marker:
        raise ValueError("Template picker has no --EMBEDDED_DATA_START marker")
    head = lua[:marker.start()]
    tail = lua[marker.end():]
    entries = []
    for m in EMBED_ENTRY.finditer(tail):
        name, url, faction, raw_json = m.groups()
        entries.append({"name": name, "imageURL": url, "faction": faction,
                        "json": json.loads(raw_json)})
    if not entries:
        raise ValueError("Template picker contains no embedded card objects")
    return head, entries


def get_old_modifiers(lua: str) -> dict[str, int]:
    match = re.search(r"\blocal\s+UPGRADE_DATA\s*=\s*\{([^{}]*)\}", lua, re.DOTALL)
    if not match:
        raise ValueError("Card LuaScript is missing local UPGRADE_DATA")
    modifiers = {}
    for field in DATA_FIELDS:
        m = re.search(r"\b" + field + r"\s*=\s*([+-]?\d+)\s*,?", match.group(1))
        modifiers[field] = int(m.group(1)) if m else 0
    return modifiers


def replace_lua_data(lua: str, upgrade_name: str, faction: str, mods: dict) -> str:
    # Keep the original card's apply-button code, including any custom behaviour.
    name_line = 'local UPGRADE_NAME = ' + json.dumps(upgrade_name, ensure_ascii=False)
    lua, n = re.subn(r"\blocal\s+UPGRADE_NAME\s*=\s*[^\r\n]*", lambda _: name_line, lua, count=1)
    if n != 1:
        raise ValueError("Could not set UPGRADE_NAME in card Lua")
    data = "local UPGRADE_DATA = {\n" + "".join(
        f"    {field:<7} = {int(mods[field])},\n" for field in DATA_FIELDS
    ) + "}"
    lua, n = re.subn(r"\blocal\s+UPGRADE_DATA\s*=\s*\{[^{}]*\}", lambda _: data,
                     lua, count=1, flags=re.DOTALL)
    if n != 1:
        raise ValueError("Could not set UPGRADE_DATA in card Lua")
    f_line = "faction = " + json.dumps(faction, ensure_ascii=False) + ";"
    lua, n = re.subn(r"(?m)^\s*faction\s*=\s*['\"][^'\"]*['\"]\s*;?", lambda _: f_line,
                     lua, count=1)
    if n != 1:
        raise ValueError("Could not set faction in card Lua")
    # Older buildDescription() does not print anything when points = 0 and
    # there are no other stat modifiers. Make that case visible in TTS.
    lua = lua.replace('return "Upgrade: " .. table.concat(parts, ", ")',
                      'if #parts == 0 then return "Upgrade: Points +0" end\n    return "Upgrade: " .. table.concat(parts, ", ")')
    return lua


def choose_old_card(item: dict, old_upgrades: list[dict], used: set[int]) -> tuple[dict | None, str]:
    fac, name, ships = item["faction"], item["name"], item["ships"]
    if norm(name) == "minelayer":
        wanted = "Minelayer - " + ships.split(",", 1)[0].strip()
    else:
        wanted = OLD_NAME_ALIASES.get((fac, name), name)
    options = [(i, card) for i, card in enumerate(old_upgrades)
               if i not in used and card["faction"] == fac]
    exact = [(i, c) for i, c in options if norm(c["name"]) == norm(wanted)]
    if len(exact) == 1:
        used.add(exact[0][0])
        return exact[0][1], "matched"
    if len(exact) > 1:
        raise ValueError(f"Ambiguous old upgrade mapping for {fac} {name}: {[v[1]['name'] for v in exact]}")
    # Conservative fallback for workbook revisions that rename an item but
    # retain the distinctive old title. Never silently fuzzy-match duplicates.
    substring = [(i, c) for i, c in options if len(norm(name)) >= 8 and
                 (norm(c["name"]).endswith(norm(name)) or norm(name) == norm(c["name"]))]
    if len(substring) == 1:
        used.add(substring[0][0])
        return substring[0][1], "suffix-matched"
    return None, "new"


def infer_modifier(effect: str, name: str, field: str) -> int:
    if "cloaking keel" in name.casefold():
        return 0  # The Excel effect also describes a mutually exclusive engine option.
    primary = re.split(r"\b(?:alternatively|or\s+cloaking\s+keel|or\s+engine\s+upgrade)\b",
                       effect, maxsplit=1, flags=re.IGNORECASE)[0]
    m = re.search(
        r"\b(?:increas(?:es|ed|ing|e)|improv(?:es|ed|ing|e))\s+"
        r"(?:its\s+|their\s+)?" + re.escape(field) + r"\s+by\s+(\d+)",
        primary, re.IGNORECASE,
    )
    return int(m.group(1)) if m else 0


def derive_modifiers(item: dict, old: dict | None, overrides: dict) -> dict[str, int]:
    mods = {field: 0 for field in DATA_FIELDS}
    if old:
        mods.update(get_old_modifiers(old["json"].get("LuaScript", "")))
    mods["points"] = item["points"]  # Excel always authoritative for points.
    for field in ("thrust", "scan"):
        inferred = infer_modifier(item["effect"], item["name"], field)
        if inferred:
            mods[field] = inferred
    mods.update(item["excel_mods"])
    ships = item["ships"].split(",", 1)[0].strip()
    specific_key = f"{item['faction']}|{item['name']}|{ships}"
    general_key = f"{item['faction']}|{item['name']}"
    config = overrides.get(specific_key, overrides.get(general_key, {}))
    if not isinstance(config, dict):
        raise ValueError(f"Overrides for {general_key} must be an object")
    for key, val in config.items():
        if key not in DATA_FIELDS:
            raise ValueError(f"Unknown override field {key!r} for {general_key}")
        mods[key] = nonnegative_int(val, f"override {general_key} {key}")
    return mods


def classify(item: dict) -> str:
    name = item["name"].casefold()
    if name == "minelayer":
        ship = item["ships"].casefold().strip()
        return {"hydra": "Cruiser", "kulshedra": "Heavy"}.get(ship, "General")
    if "grand cruiser" in name: return "Cruiser"
    if "dreadnought" in name or "10k mass" in name or "heavy dual vent" in name: return "Dreadnought"
    if "battleship" in name: return "Heavy"
    if "frigate" in name or name in {"sensor dome", "light vent cannon turret", "n-31 hybrid gun turret", "nc-16 missile turret"}: return "Frigate"
    if "cruiser" in name or item["faction"] == "Resistance": return "Cruiser"
    return "General"


def make_unique_guid(label: str, allocated: set[str]) -> str:
    for sequence in range(10000):
        h = hashlib.sha256(f"DFC-upgrade|{label}|{sequence}".encode("utf-8")).hexdigest()[:6].upper()
        if h not in allocated:
            allocated.add(h)
            return h
    raise RuntimeError("Could not generate unique 6-digit TTS GUID")


def lua_literal(text: str) -> str:
    if "]====]" in text:
        raise ValueError("Text cannot contain reserved Lua long-string closing delimiter ]====]")
    return "[====[" + text + "]====]"


def encode_cards(cards: list[dict]) -> str:
    out = [EMBED_MARKER + "\nsavedUpgradeCards = {\n"]
    for card in cards:
        obj_json = json.dumps(card["json"], ensure_ascii=False, separators=(",", ":"))
        out += [
            "  {\n",
            "    name = " + lua_literal(card["name"]) + ",\n",
            "    imageURL = " + lua_literal(card["imageURL"]) + ",\n",
            "    faction = " + lua_literal(card["faction"]) + ",\n",
            "    json = " + lua_literal(obj_json) + ",\n",
            "  },\n",
        ]
    out += ["}\n\nfresh = false\n"]
    return "".join(out)


def build(args: argparse.Namespace) -> tuple[Path, Path]:
    excel_path = select_excel(args.excel)
    template_path = Path(args.template)
    if not template_path.is_absolute():
        template_path = SCRIPT_DIR / template_path
    save = json.loads(template_path.read_text(encoding="utf-8-sig"))
    if len(save.get("ObjectStates", [])) != 1:
        raise ValueError("Expected a one-object TTS save containing the Upgrades & Admirals picker")
    picker = save["ObjectStates"][0]
    head, entries = extract_saved_cards(picker["LuaScript"])
    admirals = [c for c in entries if c["name"].strip().casefold().startswith("lvl ")]
    old_upgrades = [c for c in entries if c not in admirals]
    rows = read_upgrade_rows(excel_path, args.sheet)
    template_tile = old_upgrades[0]["json"] if old_upgrades else None
    if not template_tile:
        raise ValueError("Existing picker contains no usable upgrade tile template")
    overrides = {}
    if args.overrides:
        override_path = Path(args.overrides)
        if not override_path.is_absolute():
            override_path = SCRIPT_DIR / override_path
        overrides = json.loads(override_path.read_text(encoding="utf-8"))
        if not isinstance(overrides, dict):
            raise ValueError("Overrides file must contain a JSON object")

    allocated = {picker.get("GUID", "").upper()}
    for c in entries:
        allocated.add(str(c["json"].get("GUID", "")).upper())
    used = set()
    seen_image_paths = set()
    new_cards = []
    manifest = []
    matched = 0
    for item in rows:
        fac, name, ships, points = (item[key] for key in ("faction", "name", "ships", "points"))
        old, status = choose_old_card(item, old_upgrades, used)
        matched += old is not None
        card = copy.deepcopy(old["json"] if old else template_tile)
        display_name = old["name"] if old else unique_image_name(name, fac, ships)
        filename = image_filename(name, fac, ships)
        relative = "/".join((fac, "Upgrades", filename))
        if relative.casefold() in seen_image_paths:
            raise ValueError(f"Duplicate image destination {relative} (row {item['row']})")
        seen_image_paths.add(relative.casefold())
        img_url = args.github_base.rstrip("/") + "/" + "/".join(
            quote(part, safe="") for part in (fac, "Upgrades", filename)
        )
        mods = derive_modifiers(item, old, overrides)
        # For an existing card keep its original UPGRADE_NAME identifier so
        # already-generated ships can recognize the upgrade as the same one.
        if old:
            m = re.search(r"\blocal\s+UPGRADE_NAME\s*=\s*['\"]([^'\"]+)['\"]",
                          card["LuaScript"])
            upgrade_id = m.group(1) if m else fac + " " + display_name
        else:
            upgrade_id = fac + " " + display_name
            card["GUID"] = make_unique_guid(upgrade_id, allocated)
        card["LuaScript"] = replace_lua_data(card["LuaScript"], upgrade_id, fac, mods)
        card["LuaScriptState"] = ""
        card["XmlUI"] = ""  # reconstructed by onLoad()
        card["Nickname"] = display_name
        card["Description"] = (
            f"[{classify(item)}] Upgrade: " + ", ".join(
                f"{k.title()} +{v}" for k, v in mods.items() if v and k != "admiral"
            ) if any(v for k, v in mods.items() if k != "admiral") else
            f"[{classify(item)}] Upgrade: Points +0"
        )
        card["GMNotes"] = (
            f"Excel upgrade: {name}\nFaction: {fac}\nShips: {ships}\n"
            f"Points: {points}\nEffect:\n{item['effect']}\n"
            f"Source workbook: {excel_path.name}, {args.sheet}!A{item['row']}"
        )
        image_info = card.setdefault("CustomImage", {})
        image_info["ImageURL"] = img_url
        image_info["ImageSecondaryURL"] = img_url
        new_cards.append({"name": display_name, "imageURL": img_url, "faction": fac,
                          "json": card})
        manifest.append({
            "Excel row": item["row"], "Faction": fac,
            "Excel Name": name, "TTS Card Name": display_name,
            "Ships": ships, "Points": mods["points"], "Thrust": mods["thrust"],
            "Scan": mods["scan"], "Admiral": mods["admiral"],
            "Image Repository Path": relative, "Image URL": img_url,
            "Template Status": status,
        })
    if args.no_admirals:
        all_cards = new_cards
    else:
        all_cards = new_cards + admirals
    all_cards.sort(key=lambda c: (c["name"].casefold(), c["faction"].casefold()))
    # Picker still allows its Setup button; do not modify any of its callbacks,
    # player spawn positions, existing faction tabs, or list-spawning UI.
    picker["LuaScript"] = head + encode_cards(all_cards)
    picker["LuaScriptState"] = ""
    picker["XmlUI"] = ""
    output_path = Path(args.output)
    if not output_path.is_absolute():
        output_path = SCRIPT_DIR / output_path
    manifest_path = Path(args.manifest)
    if not manifest_path.is_absolute():
        manifest_path = SCRIPT_DIR / manifest_path
    if output_path.resolve() == template_path.resolve():
        raise ValueError("Refusing to overwrite the original TTS template object")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(save, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    with manifest_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(manifest[0]))
        writer.writeheader()
        writer.writerows(manifest)
    # Re-open and decode the embedded table to test JSON/Lua delimiter integrity.
    verify = json.loads(output_path.read_text(encoding="utf-8"))
    _, final_cards = extract_saved_cards(verify["ObjectStates"][0]["LuaScript"])
    if len(final_cards) != len(all_cards):
        raise AssertionError(f"Self-check failed: expected {len(all_cards)}, got {len(final_cards)} cards")
    for entry in final_cards:
        if not entry["json"].get("CustomImage", {}).get("ImageURL"):
            raise AssertionError(f"Missing image on {entry['name']}")
        get_old_modifiers(entry["json"].get("LuaScript", ""))
    print(f"Workbook          : {excel_path.name} / {args.sheet}")
    print(f"Existing template : {len(old_upgrades)} upgrades and {len(admirals)} admirals")
    print(f"Generated         : {len(new_cards)} upgrades ({matched} updated, {len(new_cards)-matched} new)")
    print(f"Preserved admirals: {0 if args.no_admirals else len(admirals)}")
    print(f"Final card count  : {len(final_cards)}")
    print(f"TTS object        : {output_path}")
    print(f"Image manifest    : {manifest_path}")
    print("Note: GitHub image URLs are constructed, not checked. Upload the listed PNGs first.")
    return output_path, manifest_path


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--excel", help="Workbook path (default: Upgrades.xlsx next to script)")
    p.add_argument("--sheet", default="Sheet1", help="Upgrade worksheet (default: Sheet1)")
    p.add_argument("--template", default="Upgrades  Admirals.json", help="Existing TTS picker object")
    p.add_argument("--output", default="Upgrades_Admirals_Generated.json", help="Generated TTS JSON output")
    p.add_argument("--manifest", default="Upgrades_TTS_Manifest.csv", help="Image/stat audit CSV")
    p.add_argument("--github-base", default=DEFAULT_GITHUB_RAW_BASE,
                   help="Raw GitHub root URL ending in /branch")
    p.add_argument("--overrides", help="Optional JSON with manual per-upgrade UPGRADE_DATA fields")
    p.add_argument("--no-admirals", action="store_true", help="Exclude old Admiral cards")
    args = p.parse_args()
    try:
        build(args)
    except (OSError, ValueError, KeyError, AssertionError, json.JSONDecodeError) as exc:
        p.exit(1, f"ERROR: {exc}\n")


if __name__ == "__main__":
    main()
