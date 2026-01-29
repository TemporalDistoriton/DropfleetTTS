# Button Restoration Fix

## Issue
After using the "Spawn from List" feature, the buttons (View Ships, Setup, Spawn from List) would stay hidden permanently, making the list spawner unusable after the first use.

## Root Cause
Both `onSpawnShipsClick()` and `onCloseListUIClick()` functions were:
1. Hiding the input panel ✓
2. Resetting the visibility state ✓
3. **NOT** restoring the buttons ✗

## Solution
Added `createMainButtonsWithList()` call to both functions to restore all buttons when the panel closes.

### Before (Broken)
```lua
function onSpawnShipsClick(player, value, id)
    -- ... spawn ships logic ...
    
    -- Clear the input and hide UI
    playerListInputText[player.color] = ""
    self.UI.setValue("ListInputField", "")
    self.UI.setAttribute("ListInputPanel", "active", "false")
    playerListUIVisible[player.color] = false
    -- ❌ Buttons stay hidden!
end

function onCloseListUIClick(player, value, id)
    self.UI.setAttribute("ListInputPanel", "active", "false")
    playerListUIVisible[player.color] = false
    playerListInputText[player.color] = ""
    -- ❌ Buttons stay hidden!
end
```

### After (Fixed)
```lua
function onSpawnShipsClick(player, value, id)
    -- ... spawn ships logic ...
    
    -- Clear the input and hide UI
    playerListInputText[player.color] = ""
    self.UI.setValue("ListInputField", "")
    self.UI.setAttribute("ListInputPanel", "active", "false")
    self.UI.setAttribute("ListInputPanel", "rotation", "0 0 0")
    playerListUIVisible[player.color] = false
    
    -- Restore buttons
    createMainButtonsWithList()  -- ✓ Buttons reappear!
end

function onCloseListUIClick(player, value, id)
    self.UI.setAttribute("ListInputPanel", "active", "false")
    self.UI.setAttribute("ListInputPanel", "rotation", "0 0 0")
    playerListUIVisible[player.color] = false
    playerListInputText[player.color] = ""
    
    -- Restore buttons
    createMainButtonsWithList()  -- ✓ Buttons reappear!
end
```

## User Experience Flow (Fixed)

### Scenario 1: Spawn Ships
1. Click "Spawn from List" → Buttons hide, panel appears
2. Paste fleet list
3. Click "Spawn Ships" → Ships spawn, panel closes
4. **Buttons reappear** ✓
5. Can use "Spawn from List" again

### Scenario 2: Close Panel
1. Click "Spawn from List" → Buttons hide, panel appears
2. Click "Close" button
3. **Buttons reappear** ✓
4. Can use "Spawn from List" again

### Scenario 3: Toggle Panel
1. Click "Spawn from List" → Buttons hide, panel appears
2. Click "Spawn from List" again → Panel closes
3. **Buttons reappear** ✓
4. Can toggle on/off unlimited times

## Technical Details

**Button Management:**
- `self.clearButtons()` - Removes all buttons from the tile
- `createMainButtonsWithList()` - Recreates all buttons:
  - "Spawn Ship" button
  - "Spawn from List" button
  - "View Ships" button (faction-colored)
  - "Setup" button

**Panel State:**
- `playerListUIVisible[playerColor]` - Tracks whether panel is visible for each player
- When `true` → Panel is shown, buttons are hidden
- When `false` → Panel is hidden, buttons are shown

## Files Updated
- All 8 faction tiles in TS_Save_13540.json

## Testing
✅ Can use list spawner multiple times per session
✅ Buttons reappear after spawning ships
✅ Buttons reappear after closing panel
✅ No permanent button hiding
✅ All functionality works correctly

## Commit
- b0267d2 - Fix: Restore buttons when closing list input panel or after spawning ships
