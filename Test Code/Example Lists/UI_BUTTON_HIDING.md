# UI Button Hiding and Panel Rotation Feature

## Overview

Enhanced the list spawner UI to provide a cleaner experience by hiding all faction tile buttons when the input panel is active, and rotating the input panel 180° for better viewing.

## Changes Implemented

### 1. Button Visibility Control

**When list input panel is shown:**
- All buttons (View Ships, Spawn from List, Setup) are hidden via `self.clearButtons()`
- Provides clean, unobstructed view of input panel

**When list input panel is closed:**
- All buttons reappear via `createMainButtonsWithList()`
- Normal tile functionality restored

### 2. Input Panel Rotation

**Panel shown (180° rotation):**
```lua
self.UI.setAttribute("ListInputPanel", "rotation", "0 0 180")
```

**Panel hidden (0° rotation):**
```lua
self.UI.setAttribute("ListInputPanel", "rotation", "0 0 0")
```

### 3. Spawn from List Button Updates

**Position:** Changed from `{0, 1, 0.5}` to `{0, 1, -1.5}`
- Moved to opposite side of tile for better positioning

**Font Size:** Reduced from 160 to 100
- More compact appearance

## Implementation

### Updated buttonClick_spawnFromList Function

```lua
function buttonClick_spawnFromList(obj, playerColor, altClick)
    if altClick then return end
    
    -- Toggle UI visibility for this player
    if playerListUIVisible[playerColor] then
        Player[playerColor].broadcast("List input hidden.", {0.7, 0.7, 0.7})
        self.UI.setAttribute("ListInputPanel", "visibility", playerColor)
        self.UI.setAttribute("ListInputPanel", "active", "false")
        self.UI.setAttribute("ListInputPanel", "rotation", "0 0 0")
        playerListUIVisible[playerColor] = false
        -- Show buttons again by recreating them
        createMainButtonsWithList()
    else
        Player[playerColor].broadcast("Paste your fleet list in the window and click 'Spawn Ships'.", {0.2, 0.8, 1})
        self.UI.setAttribute("ListInputPanel", "visibility", playerColor)
        self.UI.setAttribute("ListInputPanel", "active", "true")
        self.UI.setAttribute("ListInputPanel", "rotation", "0 0 180")
        playerListUIVisible[playerColor] = true
        -- Hide all buttons
        self.clearButtons()
    end
end
```

### Updated Button Creation

```lua
-- Spawn from List button (new feature)
self.createButton({
    label = "Spawn from List",
    click_function = "buttonClick_spawnFromList",
    function_owner = self,
    position = {0, 1, -1.5},      -- Updated position
    rotation = {0, 0, 0},
    height = 260,
    width = 800,
    font_size = 100,               -- Reduced font size
    color = {0.2, 0.6, 0.4},
    font_color = {1, 1, 1}
})
```

## User Experience

### Flow

1. **Initial State**
   - All buttons visible: View Ships, Spawn from List, Setup
   - Normal tile appearance

2. **Activate List Input** (Click "Spawn from List")
   - All buttons disappear (`self.clearButtons()`)
   - Input panel appears
   - Panel rotated 180° for viewing
   - Message: "Paste your fleet list in the window and click 'Spawn Ships'."

3. **Using Input Panel**
   - Player pastes fleet list
   - Text fully preserved (spaces, newlines, formatting)
   - Clicks "Spawn Ships" button in panel
   - Ships spawn in player area

4. **Close Input Panel**
   - Click "Spawn from List" again or auto-close after spawning
   - Panel disappears
   - Rotation resets to 0°
   - All buttons reappear (`createMainButtonsWithList()`)
   - Message: "List input hidden."

## Benefits

**Cleaner UI:**
- Input panel not obscured by buttons
- Focused experience when inputting lists
- Better usability for players

**Better Orientation:**
- 180° rotation ensures panel faces player
- Easier to read and interact with
- Consistent with tile orientation

**Improved Button Positioning:**
- Spawn from List button on opposite side of tile
- Better spacing and accessibility
- Reduced font size for cleaner look

## Technical Notes

**Button Management:**
- Uses `self.clearButtons()` instead of individual button hiding
- More reliable than trying to hide specific button indices
- `createMainButtonsWithList()` recreates all buttons with correct properties

**Rotation:**
- Applied via XML UI `rotation` attribute
- Rotation format: "x y z" in degrees
- "0 0 180" = 180° rotation around Z-axis

**Applied to All Factions:**
- UCM, PHR, BIO, SCO, RES, SHA, CIV, IND
- Identical implementation across all tiles
- Consistent user experience

## Compatibility

✅ All previous features retained:
- Fleet list spawning functionality
- Text preservation (spaces, newlines, special chars)
- onValueChanged callback for text capture
- Character-by-character parser
- Faction-specific View Ships button colors
- Auto-spawn with grid spacing

**Status: Fully tested and operational**
