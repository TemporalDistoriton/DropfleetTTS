#!/usr/bin/env python3
"""
TTS Ship Generator v8 — Template-Optimised (with Spawner Tiles)
================================================================
Key optimisation over v3/v4: instead of storing the full ~190KB spawner-tile
JSON per card entry (which itself contains the ~170KB ship model with ~146KB
of duplicated LuaScript), this version:

  1. Extracts the ship LuaScript into two parts:
       - Per-ship variable header (~1KB, unique per ship)
       - Template body (~145KB, identical for all ships)
  2. Stores THREE templates ONCE per faction container:
       - shipScriptTemplate   — the shared ship Lua body (~145KB)
       - spawnerScriptTemplate — the spawner-tile Lua (~1.2KB)
       - spawnerTileTemplate  — the spawner-tile TTS JSON shell (~800B)
  3. Each card stores only:
       - shipVars    — compact Lua table of per-ship variables (~600B)
       - shipObjJson — ship model JSON with scripts stripped (~13KB)
       - shipState   — the LuaScriptState JSON string (~2.5KB)
  4. At spawn time (when player picks ships in the UI), the faction container:
       a) Reconstructs the full ship LuaScript from shipVars + template
       b) Injects it into the stripped ship JSON via JSON.decode/encode
       c) Wraps the ship JSON inside a spawner-tile Lua as objectJSON
       d) Builds the spawner tile TTS JSON and spawns it on the sideboard
     The spawner tile on the sideboard is identical to what v4 produces.

Result: ~89% reduction in embedded data size per faction container.
  e.g. UCM with 67 ships: 12.8MB -> 1.4MB

Backward compatible: cards with legacy `json` field still spawn as before.

Usage:
    python Gen8_DFC_Data.py

The script uses the following files beside itself:
    MT.json                         Empty TTS save with faction containers
    OneAsMany.json                  Ship template save
    dfc-data-main/Data              Ship data repository

The generated save is written directly to the configured Tabletop Simulator
Saves folder.
"""

import json
import copy
import re
import sys
import random
from pathlib import Path
from urllib.parse import quote


SCRIPT_DIRECTORY = Path(__file__).resolve().parent

EMPTY_SAVE_FILE = SCRIPT_DIRECTORY / "MT.json"
SHIP_TEMPLATE_FILE = SCRIPT_DIRECTORY / "OneAsMany.json"
DFC_DATA_DIRECTORY = SCRIPT_DIRECTORY / "dfc-data-main" / "Data"
OUTPUT_FILE = Path(
    r"C:\Users\Gregg\Documents\My Games\Tabletop Simulator\Saves\Dropfleet_Python_Generated.json"
)

SCRIPT_VERSION = "Gen8 URL-encoded art + robust XML patch build 2026-08-02"
DEBUG_DATA_IMPORT = True


# --- Helpers ----------------------------------------------------------------

def generate_guid(existing_guids: set) -> str:
    """Generate a unique 6-char hex GUID not already in use."""
    while True:
        guid = ''.join(random.choices('0123456789abcdef', k=6))
        if guid not in existing_guids:
            existing_guids.add(guid)
            return guid


def collect_all_guids(save_data: dict) -> set:
    """Walk the entire save and collect every GUID already in use."""
    guids = set()
    def _walk(obj):
        if isinstance(obj, dict):
            if 'GUID' in obj:
                guids.add(obj['GUID'])
            for v in obj.values():
                _walk(v)
        elif isinstance(obj, list):
            for item in obj:
                _walk(item)
    _walk(save_data)
    return guids


def to_num(val):
    """Convert a value to int or float; pass through if already numeric."""
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return val if not isinstance(val, float) or val != int(val) else int(val)
    try:
        s = str(val).strip()
        if '.' in s:
            return float(s)
        return int(s)
    except (ValueError, TypeError):
        return 0


def parse_special_flags(row: dict) -> tuple:
    """Parse Special column for isVectored and isMonitored flags."""
    special = str(row.get('Special', '') or '').lower()
    is_vectored  = 'vectored' in special
    is_monitored = 'monitor'  in special
    return is_vectored, is_monitored


# --- Tonnage classification -------------------------------------------------

TONNAGE_PATTERNS = [
    ("dreadnought",     "Dreadnaught"),
    ("dreadnaught",     "Dreadnaught"),
    ("super battleship","Super Battleship"),
    ("superbattleship", "Super Battleship"),
    ("supercarrier",    "Supercarrier"),
    ("battleship",      "Battleship"),
    ("battlecruiser",   "Battlecruiser"),
    ("troopship",       "Troopship"),
    ("heavy cruiser",   "Heavy Cruiser"),
    ("light cruiser",   "Light Cruiser"),
    ("cruiser",         "Cruiser"),
    ("runner",          "Runner"),
    ("destroyer",       "Destroyer"),
    ("cutter",          "Cutter"),
    ("monitor",         "Monitor"),
    ("carrier",         "Carrier"),
    ("frigate",         "Frigate"),
    ("lighter",         "Lighter"),
    ("corvette",        "Corvette"),
    ("cell",            "Cell"),
]

SHIPBASE_TO_TONNAGE = {
    20: "Light",
    30: "Light",
    40: "Medium",
    50: "Heavy",
    60: "Colossal",
}


def detect_class_tonnage(name: str, description: str) -> str:
    text = (description.lower() + " " + name.lower())
    for pattern, tonnage in TONNAGE_PATTERNS:
        if pattern in text:
            return tonnage
    return "Other"


def detect_ship_tonnage(base_size: int) -> str:
    return SHIPBASE_TO_TONNAGE.get(base_size, "Light")


# --- Template extraction ----------------------------------------------------

TEMPLATE_SPLIT_MARKER = "local playerColor"


def extract_script_template(lua: str) -> str:
    """
    Extract the template body from the ship LuaScript.
    Returns everything from 'local playerColor' onward — the shared code
    that is identical for every ship.
    """
    idx = lua.find(TEMPLATE_SPLIT_MARKER)
    if idx == -1:
        raise RuntimeError(
            f"Could not find '{TEMPLATE_SPLIT_MARKER}' split point in template ship script. "
            "Is the ship script V8.2 or compatible?"
        )
    body = lua[idx:]

    # Apply the ShipThrust=0 station fix to the template once
    for nl in ('\r\n', '\n'):
        body = body.replace(
            f'if not ShipThrust or ShipThrust == 0 then{nl}        ShipThrust = 9  -- Default value',
            f'if not ShipThrust then{nl}        ShipThrust = 9  -- Default value'
        )

    return body


# --- Spawner tile templates (Python constants) ------------------------------

# The spawner-tile Lua script WITHOUT the objectJSON assignment.
# Uses self.getName() instead of a hardcoded ship name so the template is
# reusable across all ships.
SPAWNER_LUA_TEMPLATE = """\
-- SPAWNER TILE
-- Click the button to spawn the model on this tile

function onLoad()
  local pos = self.getPosition()
  rebuildUI()
end

function rebuildUI()
    local ui = {
        {tag='Defaults', children={
            {tag='Text', attributes={color='#cccccc', fontSize='18', alignment='MiddleLeft'}},
            {tag='InputField', attributes={fontSize='24', preferredHeight='40'}},
            {tag='ToggleButton', attributes={fontSize='18', preferredHeight='40', colors='#ffcc33|#ffffff|#808080|#606060', selectedBackgroundColor='#dddddd', deselectedBackgroundColor='#999999'}},
            {tag='Button', attributes={fontSize='12',textColor='#111111', preferredHeight='40', colors='#dddddd|#ffffff|#808080|#f6f6f6'}},
            {tag='Toggle', attributes={textColor='#cccccc'}},
        }},

        {tag='button', attributes={onClick='ui_createModel',text='Spawn Ship',  colors='#ccccccff|#ffffffff|#404040ff|#808080ff', width='120', height='20', position='0 110 -5', rotation='0 0 180' }}
    }

    self.UI.setXmlTable(ui)
end

function ui_createModel(player, value, id)
    local color = player.color
    local pos = self.getPosition()
    local rot = {x = 0, y = 0, z = 0}
    if color == "Blue" then
        rot.y = 180
    end
    spawnObjectJSON({
        json = objectJSON,
        position = {x = pos.x, y = pos.y + 2, z = pos.z},
        rotation = rot,
        callback_function = function(spawned_object)
            Wait.frames(function()
                spawned_object.call("SetPlayerColor", {playerColor = color})
            end, 30)
        end
    })
    broadcastToAll("Spawned: " .. self.getName() .. " (" .. color .. ")", {0, 1, 0})
end
"""

# The spawner-tile TTS object JSON shell — everything except Nickname,
# Description, CustomImage URLs, and LuaScript (filled in at spawn time).
SPAWNER_TILE_JSON_TEMPLATE = {
    "Name": "Custom_Tile",
    "Transform": {
        "posX": 0, "posY": 1.0, "posZ": 0,
        "rotX": 0, "rotY": 0, "rotZ": 0,
        "scaleX": 2.0, "scaleY": 1.0, "scaleZ": 2.0,
    },
    "Nickname": "",
    "Description": "",
    "GMNotes": "",
    "AltLookAngle": {"x": 0.0, "y": 0.0, "z": 0.0},
    "ColorDiffuse": {"r": 1.0, "g": 1.0, "b": 1.0},
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
    "CustomImage": {
        "ImageURL": "",
        "ImageSecondaryURL": "",
        "ImageScalar": 1.0,
        "WidthScale": 0.0,
        "CustomTile": {
            "Type": 0,
            "Thickness": 0.1,
            "Stackable": False,
            "Stretch": True,
        },
    },
    "LuaScript": "",
    "LuaScriptState": "",
    "XmlUI": "",
}


# --- Deployable Feature Lua script template ---------------------------------
# A stripped-down self-contained script: health, conditions (Activated, Spikes,
# Fire, Status submenu), ShipCard, player colour, save/load.
# No movement, no arc/signature/aura, no 3D model toggle.
# HUD pivot raised to -170 (vs -340 for ships) since there is no stalk.

FEATURE_LUA_TEMPLATE = """\
--- FEATURE SCRIPT V1.0 (Deployable Feature - auto-generated by Gen5.py)
-------------------------------------------------------------------------------
-- Variables That Change Per Feature
local SHIP_ID = "{SHIP_ID}"
local Shiphealth = {Shiphealth}
local ShipbaseSize = {ShipbaseSize}
local Shipcost = {Shipcost}
local SHIPCardURL = "{SHIPCardURL}"
local FIXED_IMAGE_URL = "{FIXED_IMAGE_URL}"

-- Player Color
local playerColor = "Blue"

local BACKDROP_URLS = {{
    Red  = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/MainMenuRed.png",
    Blue = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/MainMenuBlue.png",
}}

local MenuManager = nil
local MenuManagerGUID = "15fc7f"
local originalData = nil
local state = {{
    conditions = {{
        Fire = 0, Spikes = 0, Activated = 0, Mode = 0,
        Navigation_Offline = 0, Weapons_Offline = 0,
        Scanners_Offline = 0, Defence_Systems_Offline = 0,
        Orbital_Decay = 0, ShipCard = 0, Status = 0,
    }},
    health = {{ current = Shiphealth, max = Shiphealth }},
    base = {{ size = ShipbaseSize, color = {{r = 1, g = 1, b = 1}} }},
    statusMenuOpen = false,
    shipCardEditMode = false,
    upgradeLog = {{}},
}}

local UIStatus = {{ Blue = {{}}, Red = {{}}, Black = {{}}, Grey = {{}} }}
local currentUIMode = 0
local UI_SCALE = 0.5
local TEXT_COLOR    = "#FFFFFFFF"
local OUTLINE_COLOR = "#000000"
local OUTLINE_SIZE  = "3 3"
local BROADCAST_CHANNEL = "Ship_Activation_Reset"

local HiVis_BACKPLATE_WIDTH  = 527 * UI_SCALE
local HiVis_BACKPLATE_HEIGHT = 194 * UI_SCALE
local FEATURE_HUD_Z = -170   -- half the ship stalk height

-- Cache connected player colours
local cachedPlayerColors = {{'Blue', 'Red', 'Grey', 'Black'}}
local lastPlayerColorCheck = 0

function getConnectedColors()
    local now = os.clock()
    if now - lastPlayerColorCheck < 1 then return cachedPlayerColors end
    lastPlayerColorCheck = now
    local colors = {{}}
    for _, player in ipairs(Player.getPlayers()) do
        if UIStatus[player.color] then table.insert(colors, player.color) end
    end
    cachedPlayerColors = (#colors > 0) and colors or {{'Blue', 'Red', 'Grey', 'Black'}}
    return cachedPlayerColors
end

function IsPlayerSuscribed(color) return UIStatus[color] ~= nil end

-------------------------------------------------------------------------------
-- Conditions table (subset: only what features need)
Conditions = {{
    Fire = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/C7DDA05380D46C229FAC1DBE053A4DC1DE03B8D6_fire.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/C7DDA05380D46C229FAC1DBE053A4DC1DE03B8D6_fire.png",
        color = "#FFFFFF", stacks = true
    }},
    Spikes = {{
        url        = "https://steamusercontent-a.akamaihd.net/ugc/2450611726219379467/42D837366992A9D296B3BF03A2F99D391E6E8369/",
        active_url = "https://steamusercontent-a.akamaihd.net/ugc/2450611726219379467/42D837366992A9D296B3BF03A2F99D391E6E8369/",
        color = "#FFFFFF", stacks = true
    }},
    Activated = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/UnActivated.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/Activated.png",
        color = "#FFFFFF", stacks = false
    }},
    Mode = {{
        url        = "https://raw.githubusercontent.com/RobMayer/TTSLibrary/master/ui/gear.png",
        active_url = "https://raw.githubusercontent.com/RobMayer/TTSLibrary/master/ui/gear.png",
        color = "#FFFFFF", stacks = false, loop = 3
    }},
    Navigation_Offline = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/AA01C58BFE617F260DDDC185A5676E209B3AF6EE_Navonline_Active.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/F48FF9B41423F535D3E460A309C87C7DCB510D58_Navoffline_Active.png",
        color = "#FFFFFF", stacks = false
    }},
    Weapons_Offline = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/84543C99A3402B12FF6CAF824DAA8DEBCA25CAFC_WeaponsOnline.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/WeaponsOffline.png",
        color = "#FFFFFF", stacks = false
    }},
    Scanners_Offline = {{
        url        = "https://steamusercontent-a.akamaihd.net/ugc/2456242493963428157/A67BEAE60D330A86665ACFF07DDCE88F86D814DA/",
        active_url = "https://steamusercontent-a.akamaihd.net/ugc/1005936482532212349/B7AEBA523F75B3360AFEB439BF3C1B5983892557/",
        color = "#FFFFFF", stacks = false
    }},
    Defence_Systems_Offline = {{
        url        = "https://steamusercontent-a.akamaihd.net/ugc/2456242493963427532/0C313035666B6D233B09D5251E04E303CE521A42/",
        active_url = "https://steamusercontent-a.akamaihd.net/ugc/1005936482532207697/960058AA3C70751790AE88C667738D1E1CD95F21/",
        color = "#FFFFFF", stacks = false
    }},
    Orbital_Decay = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/118379CBFCC0095FC1B0C0007A6AF53E3D7F3ECA_DecayStable.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/A251D38D2A437456EB2FF1034C1CA064DF1D42A5_decay.png",
        color = "#FFFFFF", stacks = false
    }},
    ShipCard = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/ShipButton.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/ShipButton.png",
        color = "#FFFFFF", stacks = false
    }},
    Status = {{
        url        = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/StatusButton.png",
        active_url = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/StatusButtonActive.png",
        color = "#FFFFFF", stacks = false
    }},
}}

local TOOLTIPS = {{
    ShipCard   = "Show/Hide Feature Card",
    Activated  = "Toggle Activated Status",
    Mode       = "Switch UI Mode",
    Status     = "Open Status Effects Menu",
    Spikes     = "Energy Spikes (+/-)",
    Fire       = "Fire Markers (+/-)",
    Navigation_Offline       = "Toggle Navigation Offline",
    Weapons_Offline          = "Toggle Weapons Offline",
    Scanners_Offline         = "Toggle Scanners Offline",
    Defence_Systems_Offline  = "Toggle Defence Systems Offline",
    Orbital_Decay            = "Toggle Orbital Decay",
}}

-------------------------------------------------------------------------------
-- Scale helper (same as ship script, keeps base sizing consistent)
function getScaleFromBaseSize(baseSize)
    if baseSize == 30 then return 1.0
    elseif baseSize == 40 then return 1.3
    elseif baseSize == 50 then return 1.7
    elseif baseSize == 60 then return 2.15
    else return 1.0 end
end

-------------------------------------------------------------------------------
-- State helpers

function HasCriticalStatus()
    return (state.conditions.Fire               and state.conditions.Fire               > 0) or
           (state.conditions.Navigation_Offline  and state.conditions.Navigation_Offline  > 0) or
           (state.conditions.Orbital_Decay       and state.conditions.Orbital_Decay       > 0) or
           (state.conditions.Weapons_Offline     and state.conditions.Weapons_Offline     > 0) or
           (state.conditions.Defence_Systems_Offline and state.conditions.Defence_Systems_Offline > 0) or
           (state.conditions.Scanners_Offline    and state.conditions.Scanners_Offline    > 0)
end

function ApplyActivatedLockState()
    local activated = (state and state.conditions and (state.conditions.Activated or 0) > 0)
    if activated then
        if not self.getLock() then self.setLock(true) end
    else
        if self.getLock() then self.setLock(false) end
    end
end

function ModifyHealth(params)
    state.health = state.health or {{ current = Shiphealth, max = Shiphealth }}
    state.health.current = math.max(0, math.min(state.health.max,
        (state.health.current or 0) + (params.amount or 0)))
    SyncHealth()
end

function ModifyCondition(params)
    local prev = state.conditions[params.name] or 0
    if params.amount == 0 then
        state.conditions[params.name] = math.max(0, 1 - prev)
    elseif Conditions[params.name] and Conditions[params.name].loop then
        state.conditions[params.name] = math.max(0,
            (prev + params.amount + Conditions[params.name].loop) % Conditions[params.name].loop)
    else
        state.conditions[params.name] = math.max(0, prev + params.amount)
    end
    if params.name == 'Mode' then Sync()
    else SyncCondition(params.name) end
end

function RefreshBaseColor()
    local baseColor = state.base and state.base.color
    if type(baseColor) == "table" and baseColor.r and baseColor.g and baseColor.b then
        self.setColorTint({{ r=baseColor.r, g=baseColor.g, b=baseColor.b, a=baseColor.a or 1.0 }})
    end
end

-------------------------------------------------------------------------------
-- Sync helpers

function Sync() self.UI.setXml(ui()) end

function SyncCondition(name)
    if not state or not state.conditions or not Conditions then return end
    local conditionData = Conditions[name]
    if not conditionData then return end
    for _, color in ipairs(getConnectedColors()) do
        local conditionValue = state.conditions[name] or 0
        local isActive
        if name == "Status" then
            isActive = HasCriticalStatus()
            state.conditions.Status = isActive and 1 or 0
        else
            isActive = conditionValue > 0
        end
        local imageUrl = isActive and (conditionData.active_url or conditionData.url) or conditionData.url
        local colorBlock
        if currentUIMode == 0 then
            colorBlock = (conditionData.color or "#FFFFFF") .. "FF"
        else
            colorBlock = (conditionData.color or "#FFFFFF") .. (isActive and "FF" or "22")
        end
        self.UI.setAttributes(color .. "_ConditionImage_" .. name, {{ color=colorBlock, image=imageUrl }})
        self.UI.setAttributes(color .. "_ConditionText_"  .. name, {{
            active = (conditionData.stacks and isActive) and "true" or "false",
            text   = tostring(conditionValue)
        }})
    end
    if name == "Activated" then ApplyActivatedLockState() end
end

function SyncHealth()
    for _, color in ipairs(getConnectedColors()) do
        self.UI.setAttributes(color .. "_HealthBar_Text", {{
            text = state.health.current .. "/" .. state.health.max
        }})
        self.UI.setAttributes(color .. "_HealthBar", {{
            percentage = (state.health.current / state.health.max * 100)
        }})
    end
end

-------------------------------------------------------------------------------
-- UI button handlers

function UI_ModifyHealth(p, alt)
    if alt ~= '-3' then ModifyHealth({{ amount = (alt == '-1' and 1 or (alt == '-2' and -1) or 0) }}) end
end
function UI_ShowHealthTooltip(p, v, id)
    local tip = "HP: " .. state.health.current .. "/" .. state.health.max
    for _, c in ipairs(getConnectedColors()) do
        self.UI.setAttribute(c .. "_StaticTooltip", "text",   tip)
        self.UI.setAttribute(c .. "_StaticTooltip", "active", "true")
    end
end
function UI_HideTooltip(p, v, id)
    for _, c in ipairs(getConnectedColors()) do
        self.UI.setAttribute(c .. "_StaticTooltip", "active", "false")
    end
end
function UI_ShowTooltip(p, v, id)
    local name = id and id:match("ConditionFrame_(.+)")
    if name then
        local text = TOOLTIPS[name] or name
        for _, c in ipairs(getConnectedColors()) do
            self.UI.setAttribute(c .. "_StaticTooltip", "text",   text)
            self.UI.setAttribute(c .. "_StaticTooltip", "active", "true")
        end
    end
end
function UI_ModifyCondition(alt, name)
    if alt ~= '-3' then
        ModifyCondition({{ name = name, amount = (alt == '-1' and 1 or (alt == '-2' and -1) or 0) }})
    end
end
function UI_ModifyActivated(p, alt) UI_ModifyCondition("0", "Activated") end
function UI_ModifyMode(p, alt)
    if alt ~= '-3' then ModifyCondition({{ name="Mode", amount=(alt=='-1' and 1 or -1) }}) end
end
function UI_ModifySpikes(p, alt)
    if alt ~= '-3' then ModifyCondition({{ name="Spikes", amount=(alt=='-1' and 1 or (alt=='-2' and -1) or 0) }}) end
end
function UI_ModifyFire(p, alt)
    if alt ~= '-3' then ModifyCondition({{ name="Fire", amount=(alt=='-1' and 1 or (alt=='-2' and -1) or 0) }}) end
end
function UI_ModifyStatus(p, alt)
    if alt ~= '-3' then
        state.statusMenuOpen = not state.statusMenuOpen
        Sync()
    end
end
function UI_ModifyNavigation_Offline(p, alt)   UI_ModifyCondition("0", "Navigation_Offline")  end
function UI_ModifyWeapons_Offline(p, alt)       UI_ModifyCondition("0", "Weapons_Offline")     end
function UI_ModifyScanners_Offline(p, alt)      UI_ModifyCondition("0", "Scanners_Offline")    end
function UI_ModifyDefence_Systems_Offline(p, a) UI_ModifyCondition("0", "Defence_Systems_Offline") end
function UI_ModifyOrbital_Decay(p, alt)         UI_ModifyCondition("0", "Orbital_Decay")       end
function UI_ModifyShipCard(p, alt)
    if alt ~= '-3' then UI_ModifyCondition("0", "ShipCard"); toggleShipCard() end
end

-------------------------------------------------------------------------------
-- Ship card (feature card) panel  -- reused verbatim from ship script logic

local FEATURE_CARD_WIDTH  = 200
local FEATURE_CARD_HEIGHT = 280

function toggleShipCard()
    local isShown = (state.conditions.ShipCard or 0) > 0
    if isShown then
        print("Showing Feature Card")
    else
        print("Hiding Feature Card")
    end
    Sync()
end

function generateShipCardPanel()
    return string.format([[
        <Panel id="FeatureCardPanel" width="%d" height="%d"
               position="0 0 -340" rotation="0 270 90"
               rectAlignment="MiddleCenter">
            <Image width="100%%" height="100%%" image="%s" preserveAspect="true"/>
            <Button id="CloseFeatureCard" onClick="UI_CloseShipCard"
                    text="X" fontSize="20" fontStyle="Bold"
                    textColor="#FFFFFF" colors="#AA2222FF|#CC4444FF|#882222FF|#AA2222FF"
                    width="40" height="40"
                    rectAlignment="UpperRight" offsetXY="-5 -5"/>
        </Panel>
    ]], FEATURE_CARD_WIDTH, FEATURE_CARD_HEIGHT, SHIPCardURL)
end

function UI_CloseShipCard(p, alt)
    state.conditions.ShipCard = 0
    Sync()
end

-------------------------------------------------------------------------------
-- Status submenu (crippling effects)

function generateStatusSubmenu(color)
    local size      = 60 * UI_SCALE
    local menuWidth  = 280 * UI_SCALE
    local menuHeight = 65  * UI_SCALE
    local btns = {{
        HUDSingleCondition(color, "Navigation_Offline",      0,  0, size),
        HUDSingleCondition(color, "Orbital_Decay",           1,  0, size),
        HUDSingleCondition(color, "Fire",                    2,  0, size),
        HUDSingleCondition(color, "Weapons_Offline",         3,  0, size),
        HUDSingleCondition(color, "Scanners_Offline",       -1,  0, size),
        HUDSingleCondition(color, "Defence_Systems_Offline", 4,  0, size),
    }}
    local panelColor = currentUIMode == 0 and "#FFFFFFFF" or "#FFFFFF22"
    return string.format([[
        <Panel id="StatusSubmenu_%s" width="%d" height="%d"
               position="-23 -65 0" rectAlignment="MiddleCenter">
            <Image image="StatusDropdown" width="100%%" height="100%%" color="%s" preserveAspect="false"/>
            <Panel width="100%%" height="100%%" position="0 %d 0" color="#00000000">%s</Panel>
        </Panel>
    ]], color, menuWidth, menuHeight, panelColor, math.floor(-10*UI_SCALE), table.concat(btns))
end

-------------------------------------------------------------------------------
-- Condition button helpers (identical logic to ship script)

function HUDSingleCondition(color, name, x, y, size)
    if not color or not name or not x or not y or not size then return "" end
    local id  = "ConditionFrame_" .. name
    local pos = string.format("%s %s 0",
        ((x * (size+4) * UI_SCALE) - (1.5*size+4) * UI_SCALE),
        y * ((size+4) * UI_SCALE))
    return string.format([[
        <Panel id="%s" width="%d" height="%d" alignment="LowerLeft" position="%s"
               onClick="UI_Modify%s()" onMouseEnter="UI_ShowTooltip" onMouseExit="UI_HideTooltip">
        %s
        </Panel>
    ]], id, size*UI_SCALE, size*UI_SCALE, pos, name, HUDSingleConditionBody(color, name, size*UI_SCALE) or "")
end

function HUDSingleConditionBody(color, name, size)
    if not color or not name or not size or not state.conditions then return "" end
    local condition = Conditions[name]
    if not condition then return "" end
    local conditionValue = state.conditions[name] or 0
    local isActive       = conditionValue > 0
    local imageUrl = isActive and (condition.active_url or condition.url) or condition.url
    local colorBlock
    if currentUIMode == 0 then
        colorBlock = (condition.color or "#FFFFFF") .. "FF"
    else
        colorBlock = (condition.color or "#FFFFFF") .. (isActive and "FF" or "22")
    end
    local fontSize = math.floor(size * 0.85 * 2)
    return string.format([[
        <Image id="%s_ConditionImage_%s" image="%s" color="%s"
               rectAlignment="MiddleCenter" width="%d" height="%d"/>
        <Text  id="%s_ConditionText_%s"  active="%s" fontSize="%d"
               scale="0.5 0.5 0.5" width="%d" height="%d"
               text="%s" color="%s" fontStyle="Bold"
               rectAlignment="LowerRight" outline="%s" outlineSize="%s"/>
    ]],
    color, name, imageUrl, colorBlock, size, size,
    color, name, tostring(condition.stacks and isActive),
    fontSize, size*2, size*2, tostring(conditionValue),
    TEXT_COLOR, OUTLINE_COLOR, OUTLINE_SIZE)
end

function HUDSingleConditionRectangular(color, name, x, y, width, height)
    if not color or not name then return "" end
    local id  = "ConditionFrame_" .. name
    local pos = string.format("%s %s 0",
        ((x * (width+4) * UI_SCALE) - (1.5*width+4) * UI_SCALE),
        y * ((height+4) * UI_SCALE))
    return string.format([[
        <Panel id="%s" width="%d" height="%d" alignment="LowerLeft" position="%s"
               onClick="UI_Modify%s()" onMouseEnter="UI_ShowTooltip" onMouseExit="UI_HideTooltip">
        %s
        </Panel>
    ]], id, width*UI_SCALE, height*UI_SCALE, pos, name,
    HUDSingleConditionBody(color, name, math.min(width,height)*UI_SCALE) or "")
end

function FeatureHUDConditions(color)
    local size      = 40
    local TopRow    = 1.6
    local LeftAlign = 0.2
    local conds = {{
        HUDSingleConditionRectangular(color, "ShipCard",  LeftAlign-1.2, TopRow-0.2, size+30, size+10),
        HUDSingleCondition(color, "Spikes",    LeftAlign+1.5, TopRow,     size),
        HUDSingleCondition(color, "Activated", LeftAlign+2.75,TopRow,     size),
        HUDSingleCondition(color, "Mode",      LeftAlign+5.5, TopRow-3.5, size/2),
        HUDSingleConditionRectangular(color, "Status", LeftAlign+1.25, TopRow-2.6, 120, 60),
    }}
    return '<Panel width="100%" rectAlignment="MiddleLeft" position="0 0 0">'
        .. table.concat(conds) .. '</Panel>'
end

-------------------------------------------------------------------------------
-- Main UI builder

function rebuildAssets()
    local assets = {{}}
    local backdropURL = BACKDROP_URLS[playerColor] or BACKDROP_URLS["Blue"]
    assets[#assets+1] = {{ name="UIBackdrop",    url=backdropURL }}
    assets[#assets+1] = {{ name="StatusDropdown",
        url="https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/Assets/UI/SHIPUI/DropdownMenu.png" }}
    for condName, value in pairs(Conditions) do
        assets[#assets+1] = {{ name=condName,            url=value.url }}
        if value.active_url then
            assets[#assets+1] = {{ name=condName.."_active", url=value.active_url }}
        end
    end
    self.UI.setCustomAssets(assets)
end

function PlayerHUDPivot(color, additionalContent)
    if not UIStatus[color] then return "" end
    local content      = FeatureHUDUI(color)
    if not content then return "" end
    local extraContent = additionalContent or ""
    return string.format([[
        <Panel id="%s_PlayerHUDPivot" visibility="%s"
               height="260" width="100"
               position="0 0 %d"
               rotation="0 0 270"
               rectAlignment="MiddleCenter" childForceExpandWidth="false">
        %s
        %s
        </Panel>
    ]], color, color, FEATURE_HUD_Z, content, extraContent)
end

function FeatureHUDUI(color)
    if not state.health then return "" end
    local healthCurrent = state.health.current or 0
    local healthMax     = state.health.max     or 1
    local percentage    = (healthCurrent / healthMax * 100)
    local conditions    = FeatureHUDConditions(color) or ""
    local statusMenu    = state.statusMenuOpen and generateStatusSubmenu(color) or ""
    return string.format([[
        <Panel id="PlayerHUD_Container" active="true"
               height="%d" width="%d"
               rectAlignment="MiddleCenter"
               rotation="-35 0 0"
               position="0 50 %d"
               childForceExpandWidth="false">
            %s
            <Image image="UIBackdrop" width="100%%" height="100%%" color="#FFFFFFFF" preserveAspect="false"/>
            <Text id="%s_StaticTooltip" text="" color="#FFFFFF"
                  scale="0.25 0.25 0.25" fontSize="60" fontStyle="Bold"
                  rectAlignment="UpperCenter" offsetXY="-10 21"
                  height="90" width="950" active="false"
                  outline="#000000" outlineSize="2 2"/>
            <Panel width="100%%" height="100%%" color="#00000000">
                <Panel height="185" width="140" rectAlignment="MiddleCenter" position="-25 -10 0">
                    %s
                    <ProgressBar width="100%%" height="%d"
                        id="%s_HealthBar"
                        color="#00000080" fillImageColor="#44AA22FF"
                        percentage="%s" textColor="#00000000"/>
                    <Text id="%s_HealthBar_Text"
                          fontSize="%d" height="%d"
                          scale="0.25 0.25 0.25"
                          onClick="UI_ModifyHealth"
                          text="%s/%s" color="%s"
                          fontStyle="Bold"
                          outline="%s" outlineSize="%s"
                          onMouseEnter="UI_ShowHealthTooltip"
                          onMouseExit="UI_HideTooltip"/>
                </Panel>
            </Panel>
        </Panel>
    ]],
    HiVis_BACKPLATE_HEIGHT, HiVis_BACKPLATE_WIDTH, FEATURE_HUD_Z,
    statusMenu,
    color,
    conditions,
    math.floor(30 * UI_SCALE), color, percentage,
    color, math.floor(110 * UI_SCALE), math.floor(160 * UI_SCALE),
    healthCurrent, healthMax, TEXT_COLOR, OUTLINE_COLOR, OUTLINE_SIZE)
end

function ui()
    local panelContent = ""
    for _, color in ipairs({{'Blue', 'Red', 'Grey', 'Black'}}) do
        local cardPanel = (state.conditions.ShipCard and state.conditions.ShipCard > 0)
                          and generateShipCardPanel() or ""
        local pivot = PlayerHUDPivot(color, cardPanel)
        if pivot then panelContent = panelContent .. pivot end
    end
    return [[<Panel color="#FFFFFFff" height="0" width="0" rectAlignment="MiddleCenter" childForceExpandWidth="true">]]
        .. panelContent .. [[</Panel>]]
end

function RefreshShip()
    local saved = {{
        conditions = state.conditions,
        health     = state.health,
        base       = state.base,
    }}
    rebuildAssets()
    Wait.frames(function()
        for k, v in pairs(saved) do state[k] = v end
        self.UI.setXml(ui())
        RefreshBaseColor()
    end, 2)
end

-------------------------------------------------------------------------------
-- Spawner / colour entry points

function InitFromSpawner(params)
    if not params then return end
    state       = state       or {{}}
    state.health = state.health or {{}}
    state.base   = state.base   or {{}}

    if params.shipID   then SHIP_ID      = tostring(params.shipID)             end
    if params.baseSize then ShipbaseSize = tonumber(params.baseSize) or ShipbaseSize end
    if params.health   then Shiphealth   = tonumber(params.health)   or Shiphealth   end
    if params.points   then Shipcost     = tonumber(params.points)   or Shipcost     end
    if params.cardFrontImage then SHIPCardURL     = tostring(params.cardFrontImage) end
    if params.modelImage     then FIXED_IMAGE_URL = tostring(params.modelImage)     end
    if params.playerColor    then playerColor     = tostring(params.playerColor)    end
    if params.name           then self.setName(tostring(params.name)) end
    if params.faction        then self.setDescription("Faction: " .. tostring(params.faction)) end

    state.health.current = Shiphealth
    state.health.max     = Shiphealth
    state.base.size      = ShipbaseSize

    local targetScale = getScaleFromBaseSize(ShipbaseSize)
    self.setScale({{targetScale, targetScale, targetScale}})

    rebuildAssets()
    self.UI.setXml(ui())
    print("InitFromSpawner (Feature) OK: " .. self.getName())
end

function SetPlayerColor(params)
    if params and params.playerColor then
        playerColor = tostring(params.playerColor)
        rebuildAssets()
        self.UI.setXml(ui())
    end
end

-------------------------------------------------------------------------------
-- Save / load

function recoverState(save)
    if save.state then
        state = save.state
    end
    state.conditions             = state.conditions or {{}}
    state.conditions.Fire        = state.conditions.Fire        or 0
    state.conditions.Spikes      = state.conditions.Spikes      or 0
    state.conditions.Activated   = state.conditions.Activated   or 0
    state.conditions.Mode        = state.conditions.Mode        or 0
    state.conditions.ShipCard    = state.conditions.ShipCard    or 0
    state.conditions.Status      = state.conditions.Status      or 0
    state.conditions.Navigation_Offline      = state.conditions.Navigation_Offline      or 0
    state.conditions.Weapons_Offline         = state.conditions.Weapons_Offline         or 0
    state.conditions.Scanners_Offline        = state.conditions.Scanners_Offline        or 0
    state.conditions.Defence_Systems_Offline = state.conditions.Defence_Systems_Offline or 0
    state.conditions.Orbital_Decay           = state.conditions.Orbital_Decay           or 0
    state.base        = state.base        or {{ size=ShipbaseSize, color={{r=1,g=1,b=1}} }}
    state.base.color  = state.base.color  or {{r=1,g=1,b=1}}
    state.statusMenuOpen  = state.statusMenuOpen  or false
    state.shipCardEditMode = state.shipCardEditMode or false
    state.upgradeLog  = state.upgradeLog  or {{}}
end

function onLoad(save)
    local data = JSON.decode(save) or {{}}
    local isNew = (data.shipID or "") ~= SHIP_ID

    if not isNew then
        Shiphealth   = data.Shiphealth   or Shiphealth
        ShipbaseSize = data.ShipbaseSize or ShipbaseSize
        Shipcost     = data.Shipcost     or Shipcost
        if data.SHIPCardURL     then SHIPCardURL     = data.SHIPCardURL     end
        if data.FIXED_IMAGE_URL then FIXED_IMAGE_URL = data.FIXED_IMAGE_URL end
        if data.playerColor     then playerColor     = data.playerColor     end
    end

    local targetScale = getScaleFromBaseSize(ShipbaseSize)
    local cur = self.getScale()
    if cur.x ~= targetScale then
        self.setScale({{targetScale, targetScale, targetScale}})
    end

    recoverState(data)

    if isNew then
        state.health = {{ current=Shiphealth, max=Shiphealth }}
        state.base   = {{ size=ShipbaseSize, color={{r=1,g=1,b=1}} }}
    end

    MenuManager = getObjectFromGUID(MenuManagerGUID)

    rebuildAssets()
    self.UI.setXml(ui())
    RefreshBaseColor()
end

function onSave()
    local data = {{}}
    data.shipID        = SHIP_ID
    data.state         = state
    data.Shiphealth    = Shiphealth
    data.ShipbaseSize  = ShipbaseSize
    data.Shipcost      = Shipcost
    data.SHIPCardURL   = SHIPCardURL
    data.FIXED_IMAGE_URL = FIXED_IMAGE_URL
    data.playerColor   = playerColor
    return JSON.encode(data)
end

function onDestroy()
    -- nothing extra to clean up for a feature
end

function onUpdate()
    -- Features have no per-frame logic
end

function onCustomMessage(data)
    if data.messageID == BROADCAST_CHANNEL then
        state.conditions.Activated = 0
        SyncCondition("Activated")
        RefreshBaseColor()
        Sync()
    end
end

function RoundReset()
    state.conditions.Activated = 0
    SyncCondition("Activated")
    RefreshBaseColor()
    Sync()
end
"""


def build_feature_script(row: dict) -> str:
    """Build the complete self-contained Lua script for a Deployable Feature."""
    model_file   = str(row.get('3D MODEL FILE', 'None')).strip()
    texture_file = str(row.get('3D MODEL TEXTURE', 'None')).strip()
    model_3d_exists = model_file not in ('None', '', 'none')

    return FEATURE_LUA_TEMPLATE.format(
        SHIP_ID      = str(row['Name']),
        Shiphealth   = to_num(row.get('Hull', 1)),
        ShipbaseSize = to_num(row.get('Size', 30)),
        Shipcost     = to_num(row.get('Points', 0)),
        SHIPCardURL  = str(row.get('Card Image', '')),
        FIXED_IMAGE_URL = str(row.get('Model Image', '')),
    )


# --- Per-ship data builders -------------------------------------------------

def build_ship_vars(row: dict) -> dict:
    """Build the compact per-ship variables dict from one Excel row."""
    model_file   = str(row.get('3D MODEL FILE', 'None')).strip()
    texture_file = str(row.get('3D MODEL TEXTURE', 'None')).strip()
    model_3d_exists = model_file not in ('None', '', 'none')
    is_vectored, is_monitored = parse_special_flags(row)

    return {
        'SHIP_ID':           str(row['Name']),
        'Shiphealth':        to_num(row['Hull']),
        'ShipbaseSize':      to_num(row['Size']),
        'Shipcost':          to_num(row['Points']),
        'ShipScan':          to_num(row['Scan']),
        'ShipThrust':        to_num(row['Thrust']),
        'Signature':         to_num(row['Sig']),
        'SHIPCardURL':       str(row['Card Image']),
        'FIXED_IMAGE_URL':   str(row['Model Image']),
        'Model3DExists':     model_3d_exists,
        'TARGET_MESH_URL':   model_file if model_3d_exists else '',
        'TARGET_DIFFUSE_URL':texture_file if model_3d_exists else '',
        'Vectored':          is_vectored,
        'Monitor':           is_monitored,
    }


def build_lua_script_state(template_state_json: str, row: dict) -> str:
    """
    Build the LuaScriptState JSON for one ship, starting from the template's
    state and patching in per-ship values.
    """
    if not template_state_json or not template_state_json.strip():
        return ''
    try:
        state = json.loads(template_state_json)
    except json.JSONDecodeError:
        return template_state_json

    is_vectored, is_monitored = parse_special_flags(row)

    state['shipID']          = row['Name']
    state['Shiphealth']      = to_num(row['Hull'])
    state['ShipbaseSize']    = to_num(row['Size'])
    state['Shipcost']        = to_num(row['Points'])
    state['ShipScan']        = to_num(row['Scan'])
    state['ShipThrust']      = to_num(row['Thrust'])
    state['Signature']       = to_num(row['Sig'])
    state['SHIPCardURL']     = row['Card Image']
    state['FIXED_IMAGE_URL'] = row['Model Image']
    state['baseSignature']   = to_num(row['Sig'])
    state['Vectored']        = is_vectored
    state['Monitor']         = is_monitored

    model_3d_exists = str(row.get('3D MODEL FILE', 'None')).strip() not in ('None', '', 'none')
    state['Model3DAvailable'] = model_3d_exists

    for block_key in ('originalData', 'state'):
        block = state.get(block_key)
        if isinstance(block, dict):
            if 'health' in block:
                hp = to_num(row['Hull'])
                block['health']['current'] = hp
                block['health']['max']     = hp
            if 'base' in block:
                block['base']['size'] = to_num(row['Size'])
            if 'sig' in block:
                block['sig'] = to_num(row['Sig'])

    return json.dumps(state)


def build_stripped_ship_obj(template: dict, row: dict, existing_guids: set) -> tuple:
    """
    Create a ship model object with LuaScript="" and LuaScriptState="".
    Returns (stripped_json_str, lua_script_state_str).
    """
    ship = copy.deepcopy(template)

    # Unique GUIDs
    ship['GUID'] = generate_guid(existing_guids)
    for child in ship.get('ChildObjects', []):
        child['GUID'] = generate_guid(existing_guids)

    # Per-ship metadata
    ship['Nickname']    = str(row['Name'])
    ship['Description'] = str(row.get('Type', ''))

    # --- Deployable Feature: override base mesh, collider, and tint colour ---
    DEPLOYABLE_MESH_URL = "https://steamusercontent-a.akamaihd.net/ugc/10789802620685510/A59749C4A41CBE16BBAE3BD69B441EBEC20180CA/"
    DEPLOYABLE_COLLIDER_URL = "https://steamusercontent-a.akamaihd.net/ugc/10789802620685510/A59749C4A41CBE16BBAE3BD69B441EBEC20180CA/"
    DEPLOYABLE_COLOR = {"r": 162/255, "g": 162/255, "b": 162/255, "a": 41/255}

    is_deployable = str(row.get('Type', '')).strip() == 'Deployable Feature'

    if is_deployable:
        if 'CustomMesh' in ship:
            ship['CustomMesh']['MeshURL']     = DEPLOYABLE_MESH_URL
            ship['CustomMesh']['ColliderURL'] = DEPLOYABLE_COLLIDER_URL
        ship['ColorDiffuse']     = DEPLOYABLE_COLOR
        ship['ChildObjects']     = []   # remove stalk and 3D model child
        ship['AttachedDecals']   = []   # remove arc indicator decal
    else:
        # Update 3D model child (normal ships only)
        model_file   = str(row.get('3D MODEL FILE', 'None')).strip()
        texture_file = str(row.get('3D MODEL TEXTURE', 'None')).strip()
        model_3d_exists = model_file not in ('None', '', 'none')

        for child in ship.get('ChildObjects', []):
            if child.get('Nickname') == 'SHIP3dmodel':
                if model_3d_exists:
                    child['CustomMesh']['MeshURL']    = model_file
                    child['CustomMesh']['DiffuseURL'] = texture_file
                break

    # Build state before stripping
    ship_state = build_lua_script_state(ship.get('LuaScriptState', ''), row)

    # Strip scripts - these get reconstructed at spawn time
    ship['LuaScript']      = ''
    ship['LuaScriptState'] = ''

    return json.dumps(ship, ensure_ascii=False), ship_state


# --- Lua long-string delimiter safety --------------------------------------

def pick_safe_delimiter(text: str, start_level: int = 4) -> str:
    """Find a [=...=[ delimiter level that doesn't appear in the text."""
    level = '=' * start_level
    while f']={level}]' in text:
        level += '='
    return level


# --- Lua injection: functions added to the embedded data section ------------

LUA_HELPER_FUNCTIONS = r"""
--==============================================================================
-- TEMPLATE RECONSTRUCTION ENGINE (auto-generated by Gen5.py)
--==============================================================================

--- Build the per-ship variable declaration header from a shipVars table.
function buildScriptHeader(vars)
    if not vars then return "" end
    local lines = {}
    table.insert(lines, "--- SHIP SCRIPT V8.2 (Template-Reconstructed)")
    table.insert(lines, "-------------------------------------------------------------------------------")
    table.insert(lines, "-- Variables That Change Per ship")
    table.insert(lines, 'local SHIP_ID = "' .. tostring(vars.SHIP_ID or "Unknown") .. '"')
    table.insert(lines, "local Shiphealth = " .. tostring(vars.Shiphealth or 2))
    table.insert(lines, "local ShipbaseSize = " .. tostring(vars.ShipbaseSize or 30))
    table.insert(lines, "local Shipcost = " .. tostring(vars.Shipcost or 2))
    table.insert(lines, "local ShipScan = " .. tostring(vars.ShipScan or 2))
    table.insert(lines, "local ShipThrust = " .. tostring(vars.ShipThrust or 2))
    table.insert(lines, "local Signature = " .. tostring(vars.Signature or 2))
    table.insert(lines, 'local SHIPCardURL = "' .. tostring(vars.SHIPCardURL or "") .. '"')
    table.insert(lines, 'local FIXED_IMAGE_URL = "' .. tostring(vars.FIXED_IMAGE_URL or "") .. '"')
    table.insert(lines, "local Model3DExists = " .. tostring(vars.Model3DExists or false))
    table.insert(lines, 'local TARGET_MESH_URL = "' .. tostring(vars.TARGET_MESH_URL or "") .. '"')
    table.insert(lines, 'local TARGET_DIFFUSE_URL = "' .. tostring(vars.TARGET_DIFFUSE_URL or "") .. '"')
    table.insert(lines, "")
    table.insert(lines, "-- Ship-type flags")
    table.insert(lines, "local Vectored = " .. tostring(vars.Vectored or false))
    table.insert(lines, "local Monitor  = " .. tostring(vars.Monitor or false))
    table.insert(lines, "")
    table.insert(lines, "")
    return table.concat(lines, "\n")
end

--- Reconstruct the full ship JSON from template + per-ship data.
--- Uses TTS's built-in JSON module to safely inject the LuaScript.
function buildFullShipJSON(card)
    if not card or not card.shipVars or not card.shipObjJson then return nil end
    if not shipScriptTemplate or shipScriptTemplate == "" then return nil end

    local fullScript = buildScriptHeader(card.shipVars) .. shipScriptTemplate
    local shipObj = JSON.decode(card.shipObjJson)
    if not shipObj then
        broadcastToAll("ERROR: Failed to decode ship JSON for " .. (card.name or "?"), {1, 0, 0})
        return nil
    end

    shipObj.LuaScript      = fullScript
    shipObj.LuaScriptState = card.shipState or ""

    return JSON.encode(shipObj)
end

--- Find a safe Lua long-string delimiter level for the given text.
--- Returns the equals string (e.g. "======" for [======[...]======]).
function findSafeDelimiter(text)
    local level = 6
    while true do
        local closeDelim = "]" .. string.rep("=", level) .. "]"
        if not text:find(closeDelim, 1, true) then
            return string.rep("=", level)
        end
        level = level + 1
    end
end

--- Build the full spawner-tile JSON string from a card entry.
--- Reconstructs the ship -> wraps it in a spawner tile -> returns spawn-ready JSON.
function buildSpawnerTileJSON(card)
    -- 1. Reconstruct the full ship JSON string
    local shipJson = buildFullShipJSON(card)
    if not shipJson then return nil end

    -- 2. Build the spawner-tile Lua: script template + objectJSON assignment
    local eq = findSafeDelimiter(shipJson)
    local spawnerLua = spawnerScriptTemplate
        .. "\nobjectJSON = [" .. eq .. "[" .. shipJson .. "]" .. eq .. "]"

    -- 3. Clone the spawner tile JSON template and fill in per-ship values
    local tile = JSON.decode(spawnerTileTemplate)
    if not tile then
        broadcastToAll("ERROR: Failed to decode spawner tile template", {1, 0, 0})
        return nil
    end

    tile.Nickname    = card.name or "Unknown"
    tile.Description = card.description or ""
    tile.LuaScript   = spawnerLua

    if tile.CustomImage then
        tile.CustomImage.ImageURL          = card.imageURL or ""
        tile.CustomImage.ImageSecondaryURL = card.imageURL or ""
    end

    return JSON.encode(tile)
end

--==============================================================================
-- OVERRIDES: spawnCardsForPlayer (handles both new template + legacy formats)
--==============================================================================

function spawnCardsForPlayer(playerColor, cardIndices)
    if not playerSpawnCounts then playerSpawnCounts = {} end

    local spawnConfig = CONFIG.PLAYER_SPAWN_POSITIONS[playerColor]
    if not spawnConfig then
        spawnConfig = {
            position  = {x = 0, y = 1, z = 0},
            rotation  = {x = 0, y = 0, z = 0},
            direction = {x = 1, y = 0, z = 0},
        }
        broadcastToColor(
            "Warning: No spawn position configured for " .. playerColor .. ", using center.",
            playerColor, {1, 1, 0}
        )
    end

    local startPos     = spawnConfig.position
    local rotation     = spawnConfig.rotation
    local direction    = spawnConfig.direction
    local spacing      = CONFIG.CARD_SPAWN_SPACING
    local cardsPerRow  = CONFIG.CARDS_PER_ROW or 6

    local rowDirection = {
        x = direction.z,
        y = 0,
        z = direction.x,
    }

    local currentCount = playerSpawnCounts[playerColor] or 0

    for i, cardIndex in ipairs(cardIndices) do
        local card = savedShipCards[cardIndex]
        if card then
            local spawnJson = nil

            -- New template format: reconstruct spawner tile on-the-fly
            if card.shipObjJson and card.shipVars then
                spawnJson = buildSpawnerTileJSON(card)
            -- Legacy format: pre-built spawner tile JSON
            elseif card.json then
                spawnJson = card.json
            end

            if spawnJson then
                local col = currentCount % cardsPerRow
                local row = math.floor(currentCount / cardsPerRow)
                local colOffset = col * spacing
                local rowOffset = row * spacing
                local spawnPos = {
                    x = startPos.x + (direction.x * colOffset) + (rowDirection.x * rowOffset),
                    y = startPos.y,
                    z = startPos.z + (direction.z * colOffset) + (rowDirection.z * rowOffset),
                }

                spawnObjectJSON({
                    json     = spawnJson,
                    position = spawnPos,
                    rotation = rotation,
                    sound    = true,
                })

                currentCount = currentCount + 1
            end
        end
    end

    playerSpawnCounts[playerColor] = currentCount
end

--==============================================================================
-- OVERRIDE: onLoad (handles both new template fields + legacy stripping)
--==============================================================================

function onLoad(saved_data)
    selfGUID = self.getGUID()

    if savedShipCards == nil then savedShipCards = {} end
    if fresh == nil then fresh = true end
    if #savedShipCards > 0 then fresh = false end

    -- Initialise player state tables
    playerSelections       = {}
    playerCardIndices      = {}
    playerSearchFilters    = {}
    playerActiveTabs       = {}
    playerActiveTonnageTabs = {}
    playerSpawnCounts      = {}

    -- Clean loaded data (strip long-string delimiters if present)
    for _, card in ipairs(savedShipCards) do
        -- Legacy format
        if card.json        then card.json        = stripBrackets(card.json) end
        if card.imageURL    then card.imageURL    = stripBrackets(card.imageURL) end
        -- New template format
        if card.shipObjJson then card.shipObjJson = stripBrackets(card.shipObjJson) end
        if card.shipState   then card.shipState   = stripBrackets(card.shipState) end
        if card.description then card.description = stripBrackets(card.description) end
        if card.shipVars then
            for k, v in pairs(card.shipVars) do
                if type(v) == "string" then
                    card.shipVars[k] = stripBrackets(v)
                end
            end
        end
    end

    -- Strip templates too (safety)
    if shipScriptTemplate    then shipScriptTemplate    = stripBrackets(shipScriptTemplate) end
    if spawnerScriptTemplate then spawnerScriptTemplate = stripBrackets(spawnerScriptTemplate) end
    if spawnerTileTemplate   then spawnerTileTemplate   = stripBrackets(spawnerTileTemplate) end

    ensureShipTonnage()

    if fresh then
        createSetupButton()
    else
        createMainButtonsWithList()
    end
end

--==============================================================================
-- OVERRIDE: updateSave (writes new template format + legacy compat)
--==============================================================================

function updateSave()
    local script = self.getLuaScript()

    local dataStartPattern = "\n%-%-EMBEDDED_DATA_START"
    local dataStart = script:find(dataStartPattern)
    if dataStart then
        script = script:sub(1, dataStart - 1)
    end

    local embeddedData = "\n--EMBEDDED_DATA_START\n"

    -- Write the ship script template (if available)
    if shipScriptTemplate and shipScriptTemplate ~= "" then
        embeddedData = embeddedData .. "shipScriptTemplate = [====[" .. shipScriptTemplate .. "]====]\n\n"
    end

    -- Write the spawner templates (if available)
    if spawnerScriptTemplate and spawnerScriptTemplate ~= "" then
        embeddedData = embeddedData .. "spawnerScriptTemplate = [====[" .. spawnerScriptTemplate .. "]====]\n\n"
    end

    if spawnerTileTemplate and spawnerTileTemplate ~= "" then
        embeddedData = embeddedData .. "spawnerTileTemplate = [====[" .. spawnerTileTemplate .. "]====]\n\n"
    end

    embeddedData = embeddedData .. "savedShipCards = {\n"

    for i, card in ipairs(savedShipCards) do
        embeddedData = embeddedData .. "  {\n"
        embeddedData = embeddedData .. "    name = [====[" .. (card.name or "Unknown") .. "]====],\n"
        embeddedData = embeddedData .. "    imageURL = [====[" .. (card.imageURL or "") .. "]====],\n"
        embeddedData = embeddedData .. "    tonnage = [====[" .. (card.tonnage or "Light") .. "]====],\n"
        embeddedData = embeddedData .. "    shipTonnage = [====[" .. (card.shipTonnage or "Light") .. "]====],\n"
        embeddedData = embeddedData .. "    baseScale = " .. (card.baseScale or 1) .. ",\n"

        if card.shipVars and card.shipObjJson then
            -- New template format
            embeddedData = embeddedData .. "    description = [====[" .. (card.description or "") .. "]====],\n"
            embeddedData = embeddedData .. "    shipVars = {\n"
            for _, key in ipairs({"SHIP_ID", "SHIPCardURL", "FIXED_IMAGE_URL", "TARGET_MESH_URL", "TARGET_DIFFUSE_URL"}) do
                local v = card.shipVars[key]
                if v then
                    embeddedData = embeddedData .. "      " .. key .. " = [====[" .. tostring(v) .. "]====],\n"
                end
            end
            for _, key in ipairs({"Shiphealth", "ShipbaseSize", "Shipcost", "ShipScan", "ShipThrust", "Signature"}) do
                local v = card.shipVars[key]
                if v ~= nil then
                    embeddedData = embeddedData .. "      " .. key .. " = " .. tostring(v) .. ",\n"
                end
            end
            for _, key in ipairs({"Model3DExists", "Vectored", "Monitor"}) do
                local v = card.shipVars[key]
                if v ~= nil then
                    embeddedData = embeddedData .. "      " .. key .. " = " .. tostring(v) .. ",\n"
                end
            end
            embeddedData = embeddedData .. "    },\n"
            embeddedData = embeddedData .. "    shipState = [====[" .. (card.shipState or "") .. "]====],\n"
            embeddedData = embeddedData .. "    shipObjJson = [====[" .. (card.shipObjJson or "") .. "]====],\n"
        elseif card.json then
            -- Legacy format
            embeddedData = embeddedData .. "    json = [====[" .. card.json .. "]====],\n"
        end

        embeddedData = embeddedData .. "  },\n"
    end

    embeddedData = embeddedData .. "}\n\n"
    embeddedData = embeddedData .. "fresh = " .. tostring(fresh) .. "\n"

    script = script .. embeddedData
    self.setLuaScript(script)
end
"""


# --- Embedded data generation -----------------------------------------------

def build_embedded_data(card_entries: list, script_template: str) -> str:
    """
    Build the embedded Lua data section with:
      - Helper Lua functions (override spawnCardsForPlayer, onLoad, updateSave)
      - Ship script template (stored once)
      - Spawner tile templates (stored once)
      - Compact card entries (shipVars + stripped JSON per ship)
    """
    lines = []
    lines.append("\n--EMBEDDED_DATA_START")

    # 1. Lua helper functions and overrides
    lines.append(LUA_HELPER_FUNCTIONS)

    # 2. Ship script template (stored once, shared by all cards)
    delim = pick_safe_delimiter(script_template)
    lines.append(f"shipScriptTemplate = [={delim}[{script_template}]={delim}]")
    lines.append("")

    # 3. Spawner script template (stored once)
    sp_delim = pick_safe_delimiter(SPAWNER_LUA_TEMPLATE)
    lines.append(f"spawnerScriptTemplate = [={sp_delim}[{SPAWNER_LUA_TEMPLATE}]={sp_delim}]")
    lines.append("")

    # 4. Spawner tile JSON template (stored once)
    tile_json = json.dumps(SPAWNER_TILE_JSON_TEMPLATE, ensure_ascii=False)
    tile_delim = pick_safe_delimiter(tile_json)
    lines.append(f"spawnerTileTemplate = [={tile_delim}[{tile_json}]={tile_delim}]")
    lines.append("")

    # 5. Card data
    lines.append("savedShipCards = {")

    for card in card_entries:
        lines.append("  {")
        lines.append(f"    name = [====[{card['name']}]====],")
        lines.append(f"    description = [====[{card['description']}]====],")
        lines.append(f"    imageURL = [====[{card['imageURL']}]====],")
        lines.append(f"    tonnage = [====[{card['tonnage']}]====],")
        lines.append(f"    shipTonnage = [====[{card['shipTonnage']}]====],")
        lines.append(f"    baseScale = {card['baseScale']},")

        if 'json' in card:
            # Deployable Features use the legacy pre-built spawner-tile format.
            # The embedded Lua loader already supports this field.
            legacy_json = card.get('json', '')
            legacy_delim = pick_safe_delimiter(legacy_json)
            lines.append(f"    json = [={legacy_delim}[{legacy_json}]={legacy_delim}],")
        else:
            # Per-ship variables (compact Lua table)
            v = card['shipVars']
            lines.append("    shipVars = {")
            for key in ('SHIP_ID', 'SHIPCardURL', 'FIXED_IMAGE_URL', 'TARGET_MESH_URL', 'TARGET_DIFFUSE_URL'):
                lines.append(f"      {key} = [====[{v.get(key, '')}]====],")
            for key in ('Shiphealth', 'ShipbaseSize', 'Shipcost', 'ShipScan', 'ShipThrust', 'Signature'):
                lines.append(f"      {key} = {v.get(key, 0)},")
            for key in ('Model3DExists', 'Vectored', 'Monitor'):
                lines.append(f"      {key} = {str(v.get(key, False)).lower()},")
            lines.append("    },")

            # LuaScriptState and stripped ship JSON
            lines.append(f"    shipState = [====[{card['shipState']}]====],")
            lines.append(f"    shipObjJson = [====[{card['shipObjJson']}]====],")

        lines.append("  },")

    lines.append("}")
    lines.append("")
    lines.append("fresh = false")
    return "\n".join(lines)


# --- Ship template discovery ------------------------------------------------

def find_template_object(save_data: dict) -> tuple:
    for i, obj in enumerate(save_data.get('ObjectStates', [])):
        if 'SHIP_ID' in obj.get('LuaScript', ''):
            return i, obj
    raise RuntimeError("No template object with SHIP_ID found in the save file.")


# --- DFC JSON data import ----------------------------------------------------

ART_BASE_URL = (
    "https://raw.githubusercontent.com/TemporalDistoriton/"
    "DropfleetTTS/main/RemasterShips"
)

# The data repository uses full faction names, while the TTS save and art
# folders use the abbreviated PDF IDs below.
FACTION_TO_PDF_ID = {
    "bioficer":   "BIO",
    "bio":        "BIO",
    "civilian":   "CIV",
    "civ":        "CIV",
    "phr":        "PHR",
    "resistance": "RES",
    "res":        "RES",
    "scourge":    "SCO",
    "sco":        "SCO",
    "shaltari":   "SHA",
    "sha":        "SHA",
    "ucm":        "UCM",
}

# Any record whose requested faction container is absent from MT.json is placed
# in the Civilian container rather than being discarded.
CIVILIAN_FALLBACK_ID = "CIV"

# The published Deployable Feature schema does not include BaseSize. Keep the
# one non-30 mm value present in the former spreadsheet until the data source
# exposes it directly. Explicit JSON BaseSize/Size values always take priority.
DEPLOYABLE_BASE_SIZE_OVERRIDES = {
    "Genitor Tower": 20,
}


class DataImportError(RuntimeError):
    """Raised when the local dfc-data repository cannot be read safely."""


def load_json(path: Path):
    """Load one UTF-8 JSON file and report its path in any error."""
    try:
        with path.open('r', encoding='utf-8-sig') as f:
            return json.load(f)
    except FileNotFoundError as exc:
        raise DataImportError(f"Referenced data file does not exist: {path}") from exc
    except json.JSONDecodeError as exc:
        raise DataImportError(
            f"Invalid JSON in {path} at line {exc.lineno}, column {exc.colno}: {exc.msg}"
        ) from exc


def locate_dfc_data_directory() -> Path:
    """Return the hard-coded data repository beside this script."""
    candidate = DFC_DATA_DIRECTORY.resolve()
    if candidate.is_dir():
        return candidate

    raise DataImportError(
        "Could not find the required data folder:\n"
        f"  {candidate}\n"
        "Place dfc-data-main beside Gen5_DFC_Data.py."
    )


def faction_pdf_id(name: str) -> str:
    """Convert a repository faction/fleet name to the TTS PDF ID."""
    cleaned = str(name or '').strip()
    mapped = FACTION_TO_PDF_ID.get(cleaned.casefold())
    if mapped:
        return mapped
    # Preserve support for any future faction folder/container without silently
    # discarding it.  A missing matching TTS container is reported later.
    return re.sub(r'[^A-Za-z0-9]+', '', cleaned).upper() or 'UNKNOWN'


def special_text(value) -> str:
    """Convert the JSON Special list into the former spreadsheet cell format."""
    if value is None:
        return ''
    if isinstance(value, list):
        return ', '.join(str(item) for item in value if item is not None and str(item))
    return str(value)


def points_value(value):
    """Scenario-only records may use null points; TTS expects a numeric value."""
    return 0 if value is None else to_num(value)


def art_url(pdf_id: str, asset_name: str, suffix: str) -> str:
    """Build a URL-safe GitHub raw art URL.

    Ship and Admiral names may contain spaces, straight quotes, typographic
    quotes, ampersands, or other characters that are unsafe when the URL is
    inserted into a double-quoted TTS XML attribute.  Percent-encoding the
    filename preserves the real GitHub path while preventing the URL itself
    from breaking XML, even if the container template does not escape it.
    """
    faction = quote(str(pdf_id or '').strip(), safe='-_.~')
    filename = quote(f"{asset_name}{suffix}", safe='-_.~()')
    return f"{ART_BASE_URL}/{faction}/{filename}"


def normalise_quoted_nickname(text: str) -> str:
    """
    Convert paired ASCII nickname quotes to typographic quotes.

    The source JSON represents names such as ``\"Granite\" Halsey`` using
    ordinary double quotes. Those quotes break the faction container's XML
    ``image=\"...\"`` attribute when the art URL is inserted. The former Excel
    database and the GitHub art filenames use typographic quotes instead:
    ``“Granite” Halsey``.

    Only paired quotes are changed; unmatched quotes are left untouched.
    """
    value = str(text or '').strip()
    return re.sub(r'"([^"\r\n]+)"', r'“\1”', value)


def patch_container_xml_escaping(lua_script: str) -> tuple[str, bool]:
    """Ensure card image URLs are escaped wherever they enter TTS XML.

    Several MT.json revisions format the concatenation with different spacing,
    so an exact text replacement is too fragile.  This patches any remaining
    bare ``currentCard.imageURL`` expression, while leaving an existing
    ``escapeXml(currentCard.imageURL)`` call untouched.
    """
    if not isinstance(lua_script, str):
        return lua_script, False

    # Protect existing calls with a temporary marker, patch all bare uses, then
    # restore the already-correct calls.  The faction-container scripts use this
    # value only while constructing the XML preview image.
    marker = '__GEN8_ESCAPED_CURRENT_CARD_IMAGE_URL__'
    protected = lua_script.replace('escapeXml(currentCard.imageURL)', marker)
    patched, count = re.subn(
        r'\bcurrentCard\.imageURL\b',
        'escapeXml(currentCard.imageURL)',
        protected,
    )
    patched = patched.replace(marker, 'escapeXml(currentCard.imageURL)')
    return patched, count > 0


def base_row(
    data: dict,
    pdf_id: str,
    display_name: str,
    asset_name: str,
    type_text=None,
    *,
    profile=None,
    points=None,
    base_size=None,
    thrust=None,
) -> dict:
    """Convert a ship-like JSON record to the row interface used by Gen5."""
    profile = profile if isinstance(profile, dict) else data.get('Profile', {})
    if not isinstance(profile, dict):
        profile = {}

    row = {
        'Name':             str(display_name),
        'Points':           points_value(data.get('Points') if points is None else points),
        'Type':             str(data.get('Type', '') if type_text is None else type_text),
        'Size':             to_num(data.get('BaseSize', 30) if base_size is None else base_size),
        'Thrust':           to_num(profile.get('Thrust', 0) if thrust is None else thrust),
        'Scan':             to_num(profile.get('Scan', 0)),
        'Sig':              to_num(profile.get('Sig', 0)),
        'Hull':             to_num(profile.get('Hull', 0)),
        'ES':               profile.get('ES', '-'),
        'KS':               profile.get('KS', '-'),
        'BS':               profile.get('BS', '-'),
        'G':                profile.get('G', '-'),
        'Special':          special_text(profile.get('Special', [])),
        'PDF ID':           pdf_id,
        'Model Image':      art_url(pdf_id, asset_name, '_ModelImage.png'),
        'Card Image':       art_url(pdf_id, asset_name, '_CardFrontImage.png'),
        '3D MODEL FILE':    'None',
        '3D MODEL TEXTURE': 'None',
    }
    return row


def ship_row(data: dict, pdf_id: str) -> dict:
    """Build a normal fleet or civilian ship row."""
    name = str(data.get('Class', '')).strip()
    if not name:
        raise DataImportError("Ship record is missing its 'Class' field.")
    return base_row(data, pdf_id, name, name)


def famous_admiral_row(data: dict, pdf_id: str) -> dict:
    """Build a Famous Admiral row, retaining the old sheet's naming layout."""
    raw_admiral = str(data.get('Admiral', '')).strip()
    admiral = normalise_quoted_nickname(raw_admiral)
    ship_class = str(data.get('Class', '')).strip()
    if not admiral:
        raise DataImportError("Famous Admiral record is missing its 'Admiral' field.")

    if DEBUG_DATA_IMPORT and raw_admiral != admiral:
        print(f"  [XML SAFE] Admiral name: {raw_admiral!r} -> {admiral!r}")

    asset_name = f"{admiral} - {ship_class}" if ship_class else admiral
    level = data.get('AdmiralLevel')
    level_text = f" Level {level}" if level is not None else ''
    ship_type = str(data.get('Type', '')).strip()
    description = f"Famous Admiral{level_text}"
    if ship_type:
        description += f" & {ship_type}"

    return base_row(data, pdf_id, admiral, asset_name, description)


def hero_row(data: dict, pdf_id: str) -> dict:
    """Build a Hero row as '<Hero> - <Ship>', matching the old art names."""
    hero = str(data.get('Hero', '')).strip()
    ship_name = str(data.get('Ship', '')).strip()
    if not hero:
        raise DataImportError("Hero record is missing its 'Hero' field.")
    display_name = f"{hero} - {ship_name}" if ship_name else hero
    return base_row(data, pdf_id, display_name, display_name)


def deployable_feature_row(data: dict, pdf_id: str) -> dict:
    """Build a Deployable Feature row from its specialised profile format."""
    profile = data.get('Profile', {})
    if not isinstance(profile, dict):
        profile = {}

    name = str(profile.get('Type', data.get('Name', ''))).strip()
    if not name:
        raise DataImportError("Deployable Feature record is missing Profile.Type.")

    # The documented feature format does not require BaseSize, Scan, Sig, Hull,
    # or Thrust. Honour them when a future/extended record supplies them and use
    # safe TTS defaults otherwise. Weapon Scan is a useful fallback when present.
    weapon_scans = []
    for weapon in data.get('Weapons', []) or []:
        if isinstance(weapon, dict) and weapon.get('Scan') is not None:
            weapon_scans.append(to_num(weapon.get('Scan')))

    feature_profile = {
        'Thrust': profile.get('Thrust', 0),
        'Scan': profile.get('Scan', max(weapon_scans, default=0)),
        'Sig': profile.get('Sig', 0),
        'Hull': profile.get('Hull', 0),
        'ES': profile.get('ES', '-'),
        'KS': profile.get('KS', '-'),
        'BS': profile.get('BS', '-'),
        'G': profile.get('G', '-'),
        'Special': profile.get('Special', []),
    }
    explicit_size = data.get('BaseSize', profile.get('BaseSize', data.get('Size')))
    size = (
        explicit_size
        if explicit_size is not None
        else DEPLOYABLE_BASE_SIZE_OVERRIDES.get(name, 30)
    )

    return base_row(
        data,
        pdf_id,
        name,
        name,
        'Deployable Feature',
        profile=feature_profile,
        points=profile.get('PTS'),
        base_size=size,
        thrust=feature_profile['Thrust'],
    )


def strip_fleet_prefix(station_name: str, fleet_name: str) -> str:
    """Remove the leading fleet label used in full station data names."""
    name = str(station_name or '').strip()
    fleet = str(fleet_name or '').strip()
    prefixes = [fleet, faction_pdf_id(fleet)]
    for prefix in prefixes:
        if prefix and name.casefold().startswith((prefix + ' ').casefold()):
            return name[len(prefix):].lstrip()
    return name


def station_row(data: dict) -> dict:
    """Build a fleet-specific station row and route it to its fleet container."""
    fleet = str(data.get('Fleet', '')).strip()
    if not fleet:
        raise DataImportError("Fleet Station record is missing its 'Fleet' field.")
    pdf_id = faction_pdf_id(fleet)
    name = strip_fleet_prefix(data.get('Name', ''), fleet)
    if not name:
        raise DataImportError("Fleet Station record is missing its 'Name' field.")

    profile = data.get('Profile', {})
    return base_row(
        data,
        pdf_id,
        name,
        name,
        str(data.get('Type', 'Space Station')),
        profile=profile,
        thrust=0,
    )


def _index_filenames(index: dict, key: str) -> list:
    """Return an index list using a case-insensitive key lookup."""
    value = None
    for actual_key, actual_value in index.items():
        if str(actual_key).casefold() == key.casefold():
            value = actual_value
            break

    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        return value
    raise DataImportError(
        f"Index field '{key}' must be a filename or list of filenames, "
        f"not {type(value).__name__}."
    )


def load_index_records(
    folder: Path,
    index: dict,
    key: str,
    builder,
    pdf_id=None,
    loaded_paths: set[Path] | None = None,
) -> list:
    """Load every filename listed under one index key."""
    rows = []

    for filename in _index_filenames(index, key):
        data_path = folder / str(filename)
        try:
            data = load_json(data_path)
            row = builder(data, pdf_id) if pdf_id is not None else builder(data)
            rows.append(row)
            if loaded_paths is not None:
                loaded_paths.add(data_path.resolve())
        except DataImportError as exc:
            raise DataImportError(f"{exc} (referenced by {folder.name}/{key})") from exc
    return rows


def load_fleet_folder(folder: Path, loaded_paths: set[Path] | None = None) -> list:
    """Load Ships, Heroes, Famous Admirals and Deployable Features for a fleet."""
    index_path = folder / '_fleet.json'
    index = load_json(index_path)
    pdf_id = faction_pdf_id(folder.name)

    rows = []
    rows.extend(load_index_records(
        folder, index, 'Ships', ship_row, pdf_id, loaded_paths
    ))
    rows.extend(load_index_records(
        folder, index, 'Heroes', hero_row, pdf_id, loaded_paths
    ))
    rows.extend(load_index_records(
        folder, index, 'FamousAdmirals', famous_admiral_row, pdf_id, loaded_paths
    ))
    rows.extend(load_index_records(
        folder, index, 'DeployableFeatures', deployable_feature_row, pdf_id, loaded_paths
    ))
    return rows


def _first_present(mapping: dict, *keys, default=None):
    """Return the first non-empty value found under any case-insensitive key."""
    if not isinstance(mapping, dict):
        return default

    casefolded = {str(key).casefold(): value for key, value in mapping.items()}
    for key in keys:
        value = casefolded.get(str(key).casefold())
        if value is not None and value != '':
            return value
    return default


def normalised_folder_name(name: str) -> str:
    """Return a folder label stripped to lowercase letters and digits."""
    return re.sub(r'[^a-z0-9]+', '', str(name or '').casefold())


def is_civilian_folder(folder: Path) -> bool:
    """Recognise Civilian even when the folder contains stray spaces/punctuation."""
    return normalised_folder_name(folder.name) == 'civilian'


def iter_json_files(folder: Path) -> list[Path]:
    """Find JSON files recursively using a case-insensitive suffix check."""
    return sorted(
        (path for path in folder.rglob('*')
         if path.is_file() and path.suffix.casefold() == '.json'),
        key=lambda path: str(path).casefold(),
    )


def payload_summary(payload) -> str:
    """Return a compact diagnostic description of a decoded JSON payload."""
    if isinstance(payload, dict):
        keys = ', '.join(str(key) for key in list(payload.keys())[:12])
        if len(payload) > 12:
            keys += ', ...'
        return f"object with keys: [{keys}]"
    if isinstance(payload, list):
        return f"list with {len(payload)} item(s)"
    return type(payload).__name__


def _civilian_record_dicts(payload) -> list[dict]:
    """
    Extract ship dictionaries from the variants used by civilian exports.

    Normal fleet files contain one top-level dictionary, but this also accepts
    a one-item/list export and common wrapper keys such as Ship or Data.
    """
    if isinstance(payload, list):
        records = []
        for item in payload:
            records.extend(_civilian_record_dicts(item))
        return records

    if not isinstance(payload, dict):
        return []

    # A normal ship dictionary can be used directly.
    direct_keys = {str(key).casefold() for key in payload}
    if direct_keys.intersection({
        'class', 'name', 'shipname', 'ship_name', 'profile', 'stats',
        'thrust', 'hull', 'weapons', 'load',
    }):
        return [payload]

    # Some exports wrap the actual record in one of these fields.
    records = []
    for wrapper in ('Ship', 'Ships', 'Data', 'Record', 'Records', 'Entry', 'Entries'):
        wrapped = _first_present(payload, wrapper)
        if wrapped is not None:
            records.extend(_civilian_record_dicts(wrapped))
    return records


def civilian_ship_row(data: dict, source_path: Path) -> dict:
    """
    Build a CIV row without relying on the fleet file using exactly 'Class'.

    Civilian files are explicitly trusted as ship files, so the filename is a
    final name fallback. This guarantees files such as affluence_liner.json are
    imported even if that export uses Name, ShipName, Stats, or top-level stats.
    """
    profile = _first_present(data, 'Profile', 'Stats', 'Statline', default={})
    if not isinstance(profile, dict):
        profile = {}

    # Permit civilian exporters that place the profile fields at top level.
    merged_profile = dict(profile)
    for field, aliases in {
        'Thrust': ('Thrust',),
        'Scan': ('Scan',),
        'Sig': ('Sig', 'Signature'),
        'Hull': ('Hull', 'HP', 'Health'),
        'ES': ('ES',),
        'KS': ('KS',),
        'BS': ('BS',),
        'G': ('G', 'Group'),
        'Special': ('Special', 'SpecialRules'),
    }.items():
        if field not in merged_profile:
            value = _first_present(data, *aliases)
            if value is not None:
                merged_profile[field] = value

    filename_name = source_path.stem.replace('_', ' ').strip().title()
    display_name = str(_first_present(
        data,
        'Class', 'Name', 'ShipName', 'Ship Name', 'Ship_Name',
        default=filename_name,
    )).strip() or filename_name

    type_text = str(_first_present(
        data,
        'Type', 'ShipType', 'Ship Type', 'Role',
        default='Civilian Ship',
    )).strip()

    points = _first_present(data, 'Points', 'PTS', 'Cost', default=0)
    base_size = _first_present(data, 'BaseSize', 'Base Size', 'Size', default=30)

    # Handle textual size forms such as "M/40mm" or "40 mm".
    if isinstance(base_size, str):
        match = re.search(r'(20|25|30|32|40|50|60|80)\s*mm', base_size, re.I)
        if match:
            base_size = int(match.group(1))

    normalised = dict(data)
    normalised['Points'] = points
    normalised['Type'] = type_text
    normalised['BaseSize'] = base_size

    return base_row(
        normalised,
        CIVILIAN_FALLBACK_ID,
        display_name,
        display_name,
        type_text,
        profile=merged_profile,
        points=points,
        base_size=base_size,
    )


def load_civilian_folder(folder: Path, loaded_paths: set[Path] | None = None) -> list:
    """Load every individual JSON file beneath a Civilian folder into CIV."""
    if loaded_paths is None:
        loaded_paths = set()

    print(f"  [CIV] Entering Civilian loader: {folder.resolve()}")
    print(f"  [CIV] Folder exists={folder.exists()} directory={folder.is_dir()}")

    rows = []
    ordered_paths: list[Path] = []
    seen_paths: set[Path] = set()
    index_path = folder / '_civilian.json'

    print(f"  [CIV] Index path: {index_path} (exists={index_path.is_file()})")
    if index_path.is_file():
        index = load_json(index_path)
        if not isinstance(index, dict):
            raise DataImportError(f"Civilian index must be a JSON object: {index_path}")
        indexed_names = _index_filenames(index, 'Ships')
        print(f"  [CIV] Index keys: {list(index.keys())}")
        print(f"  [CIV] Index lists {len(indexed_names)} ship file(s)")
        for filename in indexed_names:
            path = (folder / str(filename)).resolve()
            print(f"  [CIV]   indexed -> {path.name} (exists={path.is_file()})")
            if path not in seen_paths:
                ordered_paths.append(path)
                seen_paths.add(path)

    scanned_json = [path for path in iter_json_files(folder) if not path.name.startswith('_')]
    print(f"  [CIV] Recursive scan found {len(scanned_json)} non-index JSON file(s)")
    affluence_matches = [
        path for path in scanned_json
        if path.name.casefold() == 'affluence_liner.json'
    ]
    if affluence_matches:
        for match in affluence_matches:
            print(f"  [CIV] FOUND TARGET FILE: {match.resolve()}")
    else:
        print("  [CIV] WARNING: affluence_liner.json was NOT found under this folder")

    for path in scanned_json:
        resolved = path.resolve()
        if resolved not in seen_paths:
            ordered_paths.append(resolved)
            seen_paths.add(resolved)

    for number, data_path in enumerate(ordered_paths, start=1):
        print(f"  [CIV] [{number}/{len(ordered_paths)}] Reading {data_path}")
        if data_path in loaded_paths:
            print("  [CIV]   skipped: path was already loaded")
            continue
        if not data_path.is_file():
            raise DataImportError(f"Referenced civilian ship file does not exist: {data_path}")

        payload = load_json(data_path)
        print(f"  [CIV]   decoded as {payload_summary(payload)}")
        records = _civilian_record_dicts(payload)

        # A non-index JSON in Civilian is intended as ship data. If its shape is
        # unfamiliar, still construct a row from the filename rather than drop it.
        if not records and isinstance(payload, dict):
            print("  [CIV]   no recognised wrapper/profile; using top-level object directly")
            records = [payload]

        if not records:
            print(f"  [CIV] WARNING: no object records in {data_path.name}; skipped")
            loaded_paths.add(data_path)
            continue

        print(f"  [CIV]   extracted {len(records)} record(s)")
        for record_index, record in enumerate(records, start=1):
            row = civilian_ship_row(record, data_path)
            rows.append(row)
            print(
                f"  [CIV]   imported record {record_index}: "
                f"Name='{row.get('Name')}', Type='{row.get('Type')}', "
                f"PDF ID='{row.get('PDF ID')}', Size={row.get('Size')}"
            )
        loaded_paths.add(data_path)

    print(
        f"  [CIV] COMPLETE: considered {len(ordered_paths)} JSON file(s); "
        f"prepared {len(rows)} CIV record(s)"
    )
    return rows

def load_stations_folder(folder: Path, loaded_paths: set[Path] | None = None) -> list:
    """Load fleet-specific stations; generic stations have no faction container."""
    index = load_json(folder / '_stations.json')
    return load_index_records(
        folder, index, 'FleetStations', station_row, loaded_paths=loaded_paths
    )


def build_unindexed_row(data: dict, source_folder: Path) -> dict | None:
    """
    Convert an individual ship-like JSON file that was not listed by an index.

    This is deliberately conservative: rules, systems, armaments, objectives and
    other dictionaries are ignored. Any recognised record whose source folder
    does not map to an existing faction is later redirected to Faction CIV.
    """
    if not isinstance(data, dict):
        return None

    folder_pdf_id = faction_pdf_id(source_folder.name)

    # Test specialised records before normal ships because Hero and Famous
    # Admiral files also contain a Class field.
    if data.get('Admiral') and data.get('Class'):
        return famous_admiral_row(data, folder_pdf_id)

    if data.get('Hero') and data.get('Class'):
        return hero_row(data, folder_pdf_id)

    if data.get('Fleet') and data.get('Name') and isinstance(data.get('Profile'), dict):
        return station_row(data)

    if data.get('Class') and isinstance(data.get('Profile'), dict):
        return ship_row(data, folder_pdf_id)

    profile = data.get('Profile')
    if (
        isinstance(profile, dict)
        and profile.get('Type')
        and 'PTS' in profile
        and ('Weapons' in data or 'Load' in data)
    ):
        return deployable_feature_row(data, folder_pdf_id)

    return None


def load_unindexed_ship_files(folder: Path, loaded_paths: set[Path]) -> list:
    """
    Recover ship-like JSON files omitted from, or lacking, a local index.

    The repository currently exposes indexes for normal fleets, but this scan is
    important for Civilian, Misc and future folders. It also protects against a
    stale index silently hiding a newly added ship.
    """
    rows = []

    for data_path in iter_json_files(folder):
        resolved = data_path.resolve()
        if resolved in loaded_paths or data_path.name.startswith('_'):
            continue

        try:
            data = load_json(data_path)
            row = build_unindexed_row(data, folder)
        except DataImportError as exc:
            raise DataImportError(f"{exc} (while scanning {folder.name})") from exc

        if row is None:
            continue

        loaded_paths.add(resolved)
        rows.append(row)

    return rows


def read_ship_rows_from_data(data_directory: Path) -> list:
    """Build former spreadsheet rows directly from the repository JSON files."""
    rows = []
    source_counts: dict[str, int] = {}
    loaded_paths: set[Path] = set()

    print(f"  Import build: {SCRIPT_VERSION}")
    print(f"  Data root resolved to: {data_directory.resolve()}")

    entries = sorted(data_directory.iterdir(), key=lambda path: path.name.casefold())
    print(f"  Data root contains {len(entries)} immediate item(s):")
    for entry in entries:
        kind = 'DIR ' if entry.is_dir() else 'FILE'
        print(f"    [{kind}] {entry.name!r}")

    folders = [entry for entry in entries if entry.is_dir()]
    print(f"  Top-level folders discovered ({len(folders)}): " +
          ', '.join(repr(folder.name) for folder in folders))

    # Report the exact target file anywhere below Data before any parsing.
    affluence_matches = [
        path for path in data_directory.rglob('*')
        if path.is_file() and path.name.casefold() == 'affluence_liner.json'
    ]
    if affluence_matches:
        for match in affluence_matches:
            print(f"  TARGET CHECK: affluence_liner.json exists at {match.resolve()}")
    else:
        print("  TARGET CHECK WARNING: affluence_liner.json not visible anywhere below Data")

    # Load Civilian explicitly before the generic folder loop. This avoids any
    # dependency on index presence and makes the intended route unmistakable.
    civilian_folders = [folder for folder in folders if is_civilian_folder(folder)]
    if not civilian_folders:
        # Also accept a nested Civilian directory in case the checkout gained an
        # extra wrapper directory.
        civilian_folders = [
            path for path in data_directory.rglob('*')
            if path.is_dir() and is_civilian_folder(path)
        ]

    processed_civilian_folders: set[Path] = set()
    if civilian_folders:
        print(f"  Civilian folder candidate(s): {len(civilian_folders)}")
        for civilian_folder in civilian_folders:
            resolved_folder = civilian_folder.resolve()
            if resolved_folder in processed_civilian_folders:
                continue
            processed_civilian_folders.add(resolved_folder)
            civ_rows = load_civilian_folder(civilian_folder, loaded_paths)
            rows.extend(civ_rows)
            source_counts[civilian_folder.name] = (
                source_counts.get(civilian_folder.name, 0) + len(civ_rows)
            )
    else:
        print("  WARNING: No directory whose name normalises to 'Civilian' was found")

    for folder in folders:
        if folder.resolve() in processed_civilian_folders:
            print(f"  Inspecting folder {folder.name!r}: already handled by explicit CIV loader")
            continue

        json_files = iter_json_files(folder)
        fleet_index = (folder / '_fleet.json').is_file()
        civ_index = (folder / '_civilian.json').is_file()
        stations_index = (folder / '_stations.json').is_file()
        print(
            f"  Inspecting folder {folder.name!r}: "
            f"json={len(json_files)}, _fleet={fleet_index}, "
            f"_civilian={civ_index}, _stations={stations_index}"
        )

        indexed_rows = []
        if fleet_index:
            loaded = load_fleet_folder(folder, loaded_paths)
            indexed_rows.extend(loaded)
            print(f"    fleet index produced {len(loaded)} record(s)")

        # A strangely named folder can still advertise itself through an index.
        if civ_index:
            loaded = load_civilian_folder(folder, loaded_paths)
            indexed_rows.extend(loaded)
            print(f"    civilian index/scan produced {len(loaded)} record(s)")

        if stations_index:
            loaded = load_stations_folder(folder, loaded_paths)
            indexed_rows.extend(loaded)
            print(f"    stations index produced {len(loaded)} record(s)")

        rows.extend(indexed_rows)
        recovered_rows = load_unindexed_ship_files(folder, loaded_paths)
        rows.extend(recovered_rows)
        print(f"    unindexed scan recovered {len(recovered_rows)} record(s)")

        folder_total = len(indexed_rows) + len(recovered_rows)
        source_counts[folder.name] = source_counts.get(folder.name, 0) + folder_total
        print(f"    folder total: {folder_total} ship-like record(s)")

    if not rows:
        raise DataImportError(
            f"No ship-like records were found in {data_directory}. "
            "Expected fleet/civilian/station indexes or individual ship JSON files."
        )

    unique_rows = []
    seen = set()
    for row in rows:
        key = json.dumps(row, sort_keys=True, ensure_ascii=False)
        if key in seen:
            print(
                "  WARNING: Exact duplicate data record skipped: "
                f"{row.get('PDF ID')} / {row.get('Name')} / {row.get('Type')}"
            )
            continue
        seen.add(key)
        unique_rows.append(row)

    print("  Per-folder import summary:")
    for source, count in sorted(source_counts.items(), key=lambda item: item[0].casefold()):
        print(f"    {source!r}: {count} ship-like record(s)")

    pdf_counts: dict[str, int] = {}
    for row in unique_rows:
        pdf_id = str(row.get('PDF ID', '') or '').strip().upper() or 'UNKNOWN'
        pdf_counts[pdf_id] = pdf_counts.get(pdf_id, 0) + 1
    print("  Per-container routing summary before fallback:")
    for pdf_id, count in sorted(pdf_counts.items()):
        print(f"    {pdf_id}: {count} record(s)")

    civ_rows = [
        row for row in unique_rows
        if str(row.get('PDF ID', '')).strip().upper() == CIVILIAN_FALLBACK_ID
    ]
    print(f"  Civilian/CIV records prepared: {len(civ_rows)}")
    for row in civ_rows[:20]:
        print(f"    [CIV ROW] {row.get('Name')} | {row.get('Type')}")
    if len(civ_rows) > 20:
        print(f"    ... plus {len(civ_rows) - 20} more CIV row(s)")

    return unique_rows


# --- Main -------------------------------------------------------------------

def main():
    print("=" * 78)
    print(f"Running {SCRIPT_VERSION}")
    print(f"Python source file: {Path(__file__).resolve()}")
    print("=" * 78)

    # All input and output paths are hard-coded above. Any command-line
    # arguments supplied by an IDE, launcher, or file association are ignored.
    save_path = EMPTY_SAVE_FILE
    template_path = SHIP_TEMPLATE_FILE
    output_path = OUTPUT_FILE

    # -- 1. Load the empty TTS save --
    if not save_path.is_file():
        print(f"ERROR: Empty TTS save not found: {save_path}")
        sys.exit(1)

    print(f"Loading empty TTS save: {save_path}")
    with save_path.open('r', encoding='utf-8-sig') as f:
        save_data = json.load(f)

    try:
        data_directory = locate_dfc_data_directory()
        print(f"Loading ship data from: {data_directory}")
        rows = read_ship_rows_from_data(data_directory)
    except DataImportError as exc:
        print(f"ERROR: {exc}")
        sys.exit(1)

    print(f"  Found {len(rows)} ship-like records in the data repository.")

    # -- 2. Load the ship template from OneAsMany.json --
    if not template_path.is_file():
        print(f"ERROR: Ship template save not found: {template_path}")
        sys.exit(1)

    print(f"Loading ship template: {template_path}")
    with template_path.open('r', encoding='utf-8-sig') as f:
        template_data = json.load(f)

    _, template_obj = find_template_object(template_data)
    template_idx = None
    print(f"  Template found: Nickname='{template_obj.get('Nickname')}', "
          f"GUID={template_obj['GUID']}")

    # -- 3. Extract the shared script template --
    template_lua = template_obj['LuaScript']
    script_template = extract_script_template(template_lua)
    print(f"  Script template extracted: {len(script_template):,} chars "
          f"(shared across all ships)")

    # -- 4. Discover faction container tiles --
    faction_tiles = {}
    for i, obj in enumerate(save_data.get('ObjectStates', [])):
        nick = obj.get('Nickname', '')
        if nick.startswith('Faction '):
            # Container IDs are treated case-insensitively. Store them in the
            # same uppercase form used by the imported data rows.
            pdf_id = nick.split('Faction ', 1)[1].strip().upper()
            faction_tiles[pdf_id] = i
            print(f"  Found container tile: '{nick}' (PDF ID '{pdf_id}') at index {i}")

    if not faction_tiles:
        print("ERROR: No faction container tiles found "
              "(expected 'Faction BIO', 'Faction UCM', etc.)")
        sys.exit(1)

    if CIVILIAN_FALLBACK_ID not in faction_tiles:
        print(
            "ERROR: The fallback container 'Faction CIV' was not found in MT.json. "
            "Add that container before generating the save."
        )
        sys.exit(1)

    # -- 5. Collect existing GUIDs --
    existing_guids = collect_all_guids(save_data)

    # -- 6. Process each ship row --
    groups: dict[str, list] = {}
    vectored_count  = 0
    monitored_count = 0

    for i, row in enumerate(rows):
        # Blank, malformed, or differently-cased faction IDs are normalised
        # here. Unknown IDs are redirected to CIV after all cards are built.
        pdf_id    = str(row.get('PDF ID', '') or '').strip().upper() or 'UNKNOWN'
        ship_name = str(row['Name'])
        ship_type = str(row.get('Type', ''))
        card_url  = str(row.get('Card Image', ''))

        if DEBUG_DATA_IMPORT and ('Granite' in ship_name or 'Halsey' in ship_name):
            print(f"  [URL SAFE] {ship_name!r} card URL: {card_url}")
        base_size = to_num(row['Size'])

        is_vectored, is_monitored = parse_special_flags(row)
        if is_vectored:  vectored_count  += 1
        if is_monitored: monitored_count += 1

        is_deployable = ship_type.strip() == 'Deployable Feature'

        # b) Build stripped ship JSON (always needed for mesh/collider/colour overrides)
        ship_obj_json, ship_state = build_stripped_ship_obj(
            template_obj, row, existing_guids
        )

        if is_deployable:
            # Features use a self-contained Lua script; embed it directly so the
            # spawner tile uses the card.json (legacy) path — no template reconstruction.
            feature_lua  = build_feature_script(row)
            ship_obj      = json.loads(ship_obj_json)
            ship_obj['LuaScript']      = feature_lua
            ship_obj['LuaScriptState'] = ship_state
            full_ship_json = json.dumps(ship_obj, ensure_ascii=False)

            # Wrap in a spawner tile immediately (same as legacy path)
            spawner_tile = json.loads(json.dumps(SPAWNER_TILE_JSON_TEMPLATE))  # deep copy
            spawner_tile['Nickname']    = ship_name
            spawner_tile['Description'] = ship_type
            if spawner_tile.get('CustomImage'):
                spawner_tile['CustomImage']['ImageURL']          = card_url
                spawner_tile['CustomImage']['ImageSecondaryURL'] = card_url

            delim = pick_safe_delimiter(full_ship_json)
            spawner_lua = (SPAWNER_LUA_TEMPLATE
                           + f"\nobjectJSON = [={delim}[{full_ship_json}]={delim}]")
            spawner_tile['LuaScript'] = spawner_lua

            card_entry = {
                "name":        ship_name,
                "description": ship_type,
                "imageURL":    card_url,
                "tonnage":     detect_class_tonnage(ship_name, ship_type),
                "shipTonnage": detect_ship_tonnage(base_size),
                "baseScale":   1,
                # Legacy path: pre-built spawner tile JSON (no template reconstruction)
                "json":        json.dumps(spawner_tile, ensure_ascii=False),
            }
        else:
            # a) Build per-ship variables (normal ships only)
            ship_vars = build_ship_vars(row)

            # c) Build the card entry (compact template format)
            card_entry = {
                "name":        ship_name,
                "description": ship_type,
                "imageURL":    card_url,
                "tonnage":     detect_class_tonnage(ship_name, ship_type),
                "shipTonnage": detect_ship_tonnage(base_size),
                "baseScale":   1,
                "shipVars":    ship_vars,
                "shipState":   ship_state,
                "shipObjJson": ship_obj_json,
            }

        groups.setdefault(pdf_id, []).append(card_entry)

        if (i + 1) % 50 == 0:
            print(f"  Processed {i + 1}/{len(rows)} ships...")

    total_ships = len(rows)
    print(f"  Created {total_ships} ship entries across {len(groups)} factions.")
    print("  Built card groups before fallback:")
    for group_id, entries in sorted(groups.items()):
        print(f"    {group_id}: {len(entries)} card(s)")
    print(f"  Special flags detected:")
    print(f"    isVectored  = true : {vectored_count} ships")
    print(f"    isMonitored = true : {monitored_count} ships")
    print(f"    Neither flag       : {total_ships - vectored_count - monitored_count} ships")

    # -- 7. Redirect records without a matching faction container to CIV --
    # This includes unknown future factions, blank IDs, and any faction whose
    # container has been omitted from MT.json. Existing CIV records remain in
    # the same group and the redirected cards are appended to them.
    fallback_count = 0
    for pdf_id in list(groups):
        if pdf_id in faction_tiles:
            continue

        orphaned_entries = groups.pop(pdf_id)
        groups.setdefault(CIVILIAN_FALLBACK_ID, []).extend(orphaned_entries)
        fallback_count += len(orphaned_entries)
        print(
            f"  WARNING: No container tile found for PDF ID '{pdf_id}'. "
            f"Redirected {len(orphaned_entries)} ship(s) to 'Faction CIV'."
        )

    if fallback_count:
        print(f"  Redirected {fallback_count} total ship(s) to the CIV fallback container.")

    # -- 8. Sort each group alphabetically --
    for pdf_id in groups:
        groups[pdf_id].sort(key=lambda c: c['name'].casefold())

    # -- 9. Embed card data into each faction container tile --
    for pdf_id, card_entries in groups.items():
        tile_idx  = faction_tiles[pdf_id]
        container = save_data['ObjectStates'][tile_idx]
        lua_script = container['LuaScript']

        # XML safety: older MT templates insert the card image URL directly
        # into an XML attribute. Escape it at render time so quoted Admiral
        # names cannot invalidate the entire faction UI.
        lua_script, xml_patch_applied = patch_container_xml_escaping(lua_script)
        if xml_patch_applied:
            print(f"  Applied XML image-URL escaping to '{container['Nickname']}'")

        # Remove any existing embedded data section
        for marker in ("\r\n--EMBEDDED_DATA_START", "\n--EMBEDDED_DATA_START"):
            pos = lua_script.rfind(marker)
            if pos != -1:
                lua_script = lua_script[:pos]
                break

        # Build new embedded data
        embedded = build_embedded_data(card_entries, script_template)
        lua_script = lua_script + embedded

        container['LuaScript'] = lua_script
        print(f"  Embedded {len(card_entries)} ships into '{container['Nickname']}' "
              f"({len(embedded):,} chars)")

    # -- 10. Remove the template object from the save --
    if template_idx is not None:
        save_data['ObjectStates'] = [
            obj for i, obj in enumerate(save_data['ObjectStates'])
            if i != template_idx
        ]

    # -- 11. Write output --
    output_path.parent.mkdir(parents=True, exist_ok=True)
    print(f"\nWriting output to: {output_path}")
    with output_path.open('w', encoding='utf-8') as f:
        json.dump(save_data, f, ensure_ascii=False)

    print("Done!")
    print(f"  Total objects on table: {len(save_data['ObjectStates'])}")


if __name__ == '__main__':
    main()
