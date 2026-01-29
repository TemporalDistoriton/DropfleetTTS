# Graphic Design Improvements

## Overview
This document details the improvements made based on feedback from the graphic design team.

## Changes Implemented

### 1. Enhanced User Instructions

**Before:**
```
Paste your fleet list below and click Spawn Ships
```

**After:**
```
In New Recruit, Open your List, Click Export,
Click Text, Click Copy to clipboard, then paste below
```

**Details:**
- Added step-by-step instructions for users unfamiliar with the process
- Split into two lines for better readability
- Increased font size from 12 to 14 for better visibility
- Positioned at offsetXY="0 -70" and offsetXY="0 -90"

### 2. Reduced Overall Scale

**Before:**
- No scale attribute (default 1.0)
- Panel appeared very large on screen

**After:**
- Scale: `.28 .28 .28`
- Matches the View Ships UI exactly
- Same size as Faction Box Tile

**Benefits:**
- More compact and professional appearance
- Consistent with existing UI elements
- Doesn't dominate the screen
- Better integration with game table

### 3. Fixed Input Field Positioning

**Before:**
- InputField position: `0 -180 0`
- Height: 400
- Was hanging off the bottom of the panel

**After:**
- InputField position: `0 -250 0`
- Height: 350
- Starts just after instruction text
- Button positions adjusted:
  - Spawn Ships: `0 -460 0`
  - Close: `0 -520 0`

**Visual Layout:**
```
┌─────────────────────────────┐
│ Spawn Ships from List       │ ← Title (offsetXY: 0 -30)
│                             │
│ In New Recruit, Open your   │ ← Instructions line 1 (offsetXY: 0 -70)
│ List, Click Export,         │
│ Click Text, Click Copy to   │ ← Instructions line 2 (offsetXY: 0 -90)
│ clipboard, then paste below │
│                             │
│ ┌─────────────────────────┐ │
│ │                         │ │ ← InputField (position: 0 -250 0)
│ │   Paste fleet list...   │ │   (height: 350)
│ │                         │ │
│ │                         │ │
│ └─────────────────────────┘ │
│                             │
│    [ Spawn Ships ]          │ ← Button (position: 0 -460 0)
│    [    Close    ]          │ ← Button (position: 0 -520 0)
└─────────────────────────────┘
```

### 4. Fixed View Ships Ship Type Sorting

**Problem:**
- Tonnage categories were generic: Light, Medium, Heavy, Colossal
- Didn't match actual ship types in the game
- Sorting was no longer working properly

**Solution:**
Implemented specific ship type categories that match game terminology.

**New Categories (in order):**
1. Other
2. Cell
3. Corvette
4. Lighter
5. Frigate
6. Carrier
7. Monitor
8. Cutter
9. Destroyer
10. Runner
11. Light Cruiser
12. Cruiser
13. Heavy Cruiser
14. Troopship
15. Battlecruiser
16. Battleship
17. Supercarrier
18. Super Battleship
19. Dreadnaught

**Detection Logic:**
```lua
-- Check both description and name for ship type
local text = descLower .. " " .. nameLower

-- Ordered from largest to smallest for proper detection
if string.find(text, "dreadnought") or string.find(text, "dreadnaught") then
    cardData.tonnage = "Dreadnaught"
elseif string.find(text, "super battleship") or string.find(text, "superbattleship") then
    cardData.tonnage = "Super Battleship"
elseif string.find(text, "supercarrier") then
    cardData.tonnage = "Supercarrier"
-- ... continues for all types ...
else
    cardData.tonnage = "Other" -- Default
end
```

**Key Features:**
- Checks both card description AND card name
- Handles compound names (e.g., "Super Battleship" before "Battleship")
- Ordered from largest to smallest to avoid misclassification
- Case-insensitive matching
- Defaults to "Other" if no type is found

**Benefits:**
- Precise categorization matching game terminology
- Easy to find specific ship types in UI
- Works reliably with existing card data
- Supports alternate spellings (dreadnought/dreadnaught)

## Technical Implementation

### XML UI Changes
```xml
<Panel id="ListInputPanel" 
       position="0 10 -110" 
       width="600" 
       height="650" 
       color="rgba(0.1, 0.1, 0.15, 0.95)" 
       active="false" 
       visibility="" 
       scale=".28 .28 .28"           ← Added scale
       rotation="0 0 180">
    
    <Text fontSize="28" ...>Spawn Ships from List</Text>
    
    <Text fontSize="14" offsetXY="0 -70" ...>     ← New instruction line 1
        In New Recruit, Open your List, Click Export,
    </Text>
    
    <Text fontSize="14" offsetXY="0 -90" ...>     ← New instruction line 2
        Click Text, Click Copy to clipboard, then paste below
    </Text>
    
    <InputField id="ListInputField" 
                position="0 -250 0"               ← Moved down from -180
                height="350" ... />               ← Reduced from 400
    
    <Button onClick="onSpawnShipsClick" 
            position="0 -460 0" ... />            ← Adjusted from -420
    
    <Button onClick="onCloseListUIClick" 
            position="0 -520 0" ... />            ← Adjusted from -480
</Panel>
```

### Lua Script Changes

**CONFIG.UI.TONNAGE_CATEGORIES:**
```lua
TONNAGE_CATEGORIES = {
    "Other", "Cell", "Corvette", "Lighter", "Frigate", "Carrier", "Monitor", 
    "Cutter", "Destroyer", "Runner", "Light Cruiser", "Cruiser", "Heavy Cruiser", 
    "Troopship", "Battlecruiser", "Battleship", "Supercarrier", "Super Battleship", 
    "Dreadnaught"
}
```

**Tonnage Detection Function:**
- Replaced generic Light/Medium/Heavy detection
- Added comprehensive ship type checking
- Checks both description and name fields
- Handles special cases and compound names

## Testing

### Visual Testing
- [x] Panel scales correctly to .28 .28 .28
- [x] Instructions are clearly visible and readable
- [x] Input field doesn't hang off panel
- [x] All elements fit within panel bounds
- [x] Consistent with View Ships UI appearance

### Functional Testing
- [x] Ship type detection works for all categories
- [x] View Ships tabs display correct ship types
- [x] Sorting by ship type works correctly
- [x] Default "Other" category catches unrecognized ships
- [x] List spawner still functions correctly

### Compatibility
- [x] All 8 faction tiles updated identically
- [x] No breaking changes to existing functionality
- [x] Backwards compatible with saved cards
- [x] Works with all test fleet lists

## Files Modified

- `TS_Save_13540.json` - All 8 faction tiles updated

## Summary

All graphic design feedback has been successfully implemented:
1. ✅ Detailed user instructions added
2. ✅ Panel scale reduced to match View Ships UI
3. ✅ Input field positioning fixed
4. ✅ View Ships sorting updated to use specific ship types

The list spawner now provides better user guidance and matches the visual design of existing UI elements, while View Ships offers precise ship categorization matching game terminology.
