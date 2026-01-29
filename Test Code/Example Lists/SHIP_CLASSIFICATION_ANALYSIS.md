# Ship Classification System - Complete Analysis

## Overview

This document provides a complete analysis of how the ship classification and filtering system works in the faction tiles.

## How Ship Classification Works

### 1. When Ships Are Added (Setup Flow)

**User Action:** Click "Setup" button → Box-select ship cards → Click "Add Selected Ships"

**Code Flow:**
```lua
function buttonClick_addSelectedCards()
    local cards = Player[playerColor].getSelectedObjects()
    
    for _, obj in ipairs(cards) do
        -- Extract card data including ship type classification
        local cardData = extractCardData(obj)
        
        -- Store in savedShipCards array
        table.insert(savedShipCards, cardData)
    end
end
```

### 2. Ship Type Detection (extractCardData)

**Function:** `extractCardData(cardObj)`

**Detection Logic:**
```lua
function extractCardData(cardObj)
    local cardData = {}
    
    -- Get description and name from card
    local desc = cardObj.getDescription() or ""
    local descLower = string.lower(desc)
    local name = cardObj.getName() or ""
    local nameLower = string.lower(name)
    
    -- Combine both for matching
    local text = descLower .. " " .. nameLower
    
    -- Check ship types in hierarchical order (most specific first)
    if string.find(text, "dreadnought") or string.find(text, "dreadnaught") then
        cardData.tonnage = "Dreadnaught"
    elseif string.find(text, "super battleship") or string.find(text, "superbattleship") then
        cardData.tonnage = "Super Battleship"
    elseif string.find(text, "supercarrier") then
        cardData.tonnage = "Supercarrier"
    elseif string.find(text, "battleship") then
        cardData.tonnage = "Battleship"
    elseif string.find(text, "battlecruiser") then
        cardData.tonnage = "Battlecruiser"
    elseif string.find(text, "troopship") then
        cardData.tonnage = "Troopship"
    elseif string.find(text, "heavy cruiser") then
        cardData.tonnage = "Heavy Cruiser"
    elseif string.find(text, "light cruiser") then
        cardData.tonnage = "Light Cruiser"
    elseif string.find(text, "cruiser") then
        cardData.tonnage = "Cruiser"
    elseif string.find(text, "runner") then
        cardData.tonnage = "Runner"
    elseif string.find(text, "destroyer") then
        cardData.tonnage = "Destroyer"
    elseif string.find(text, "cutter") then
        cardData.tonnage = "Cutter"
    elseif string.find(text, "monitor") then
        cardData.tonnage = "Monitor"
    elseif string.find(text, "carrier") then
        cardData.tonnage = "Carrier"
    elseif string.find(text, "frigate") then
        cardData.tonnage = "Frigate"
    elseif string.find(text, "lighter") then
        cardData.tonnage = "Lighter"
    elseif string.find(text, "corvette") then
        cardData.tonnage = "Corvette"
    elseif string.find(text, "cell") then
        cardData.tonnage = "Cell"
    else
        cardData.tonnage = "Other"
    end
    
    return cardData
end
```

**Key Points:**
- Checks both card description AND card name
- Uses case-insensitive matching (`string.lower()`)
- Checks most specific types BEFORE generic types
- Falls back to "Other" if no match found

### 3. View Ships Tab Filtering

**User Action:** Click "View Ships" button → Click a category tab (e.g., "Cruiser")

**Code Flow:**
```lua
function uiTabClicked(tabName)
    activeTab = tabName  -- e.g., "Cruiser", "Battleship", "All"
    rebuildCards()       -- Rebuild UI with filtered cards
end

function rebuildCards()
    local filtered = getFilteredCards(activeTab, searchTerm)
    
    -- Build UI showing only filtered cards
    for _, card in ipairs(filtered) do
        -- Display card
    end
end

function getFilteredCards(activeTab, searchFilter)
    local filtered = {}
    
    for i, card in ipairs(savedShipCards) do
        -- Check if card matches the active tab
        local matchesTab = (activeTab == "All") or (card.tonnage == activeTab)
        
        -- Check if card matches search filter (if any)
        local matchesSearch = (searchFilter == nil or searchFilter == "" or 
                              string.find(string.lower(card.name), string.lower(searchFilter)))
        
        if matchesTab and matchesSearch then
            table.insert(filtered, card)
        end
    end
    
    return filtered
end
```

**Key Points:**
- **"All" tab:** Shows all ships (no filtering by tonnage)
- **Specific tabs:** Only shows ships where `card.tonnage == activeTab`
- **Search filter:** Additional filtering by ship name (if search box used)

## Why Ships Might Not Appear in Tabs

### Issue 1: Ships Not Added Yet
**Problem:** No ships stored in `savedShipCards` array
**Solution:** Click "Setup" → Box-select cards → Click "Add Selected Ships"

### Issue 2: Ship Type Not Detected
**Problem:** Ship description/name doesn't contain recognized keywords
**Check:** What's in the card's description? Does it have "Cruiser", "Battleship", etc.?
**Result:** Ship will be classified as "Other"

### Issue 3: Tab Name Mismatch
**Problem:** Tab name doesn't exactly match the tonnage value
**Example:**
- Tab is "Cruiser" but tonnage is "Cruisers" (with 's')
- Tab is "Light Cruiser" but tonnage is "LightCruiser" (no space)

### Issue 4: Detection Order Wrong
**Problem:** Generic pattern matches before specific pattern
**Example (WRONG):**
```lua
if string.find(text, "cruiser") then  -- Matches "Light Cruiser" first!
    cardData.tonnage = "Cruiser"
elseif string.find(text, "light cruiser") then  -- Never reached
    cardData.tonnage = "Light Cruiser"
```

**Example (CORRECT):**
```lua
if string.find(text, "light cruiser") then  -- Check specific first
    cardData.tonnage = "Light Cruiser"
elseif string.find(text, "cruiser") then  -- Then generic
    cardData.tonnage = "Cruiser"
```

## Tab Categories Defined

The available tabs are defined in `TONNAGE_CATEGORIES`:

```lua
TONNAGE_CATEGORIES = {
    "Other", 
    "Cell", 
    "Corvette", 
    "Lighter", 
    "Frigate", 
    "Carrier", 
    "Monitor", 
    "Cutter", 
    "Destroyer", 
    "Runner", 
    "Light Cruiser", 
    "Cruiser", 
    "Heavy Cruiser", 
    "Troopship", 
    "Battlecruiser", 
    "Battleship", 
    "Supercarrier", 
    "Super Battleship", 
    "Dreadnaught"
}
```

## Example Ship Classifications

### UCM Ships
- **Boston Light Cruiser** → "Light Cruiser" tab
  - Description contains "Light Cruiser"
- **Madrid Cruiser** → "Cruiser" tab
  - Description contains "Cruiser" but not "Light" or "Heavy"
- **New Cairo Heavy Cruiser** → "Heavy Cruiser" tab
  - Description contains "Heavy Cruiser"
- **Beijing Battlecruiser** → "Battlecruiser" tab
  - Description contains "Battlecruiser"

### PHR Ships
- **Bellerophon Battlecruiser** → "Battlecruiser" tab
- **Orpheus Cruiser** → "Cruiser" tab
- **Theseus Light Cruiser** → "Light Cruiser" tab

### Scourge Ships
- **Yokai Cruiser** → "Cruiser" tab
- **Djinn Cruiser** → "Cruiser" tab

### Shaltari Ships
- **Selenium Cruiser** → "Cruiser" tab
- **Caesium Battlecruiser** → "Battlecruiser" tab

## Debugging Steps

### Step 1: Check if ships are stored
Add debug output after adding ships:
```lua
function buttonClick_addSelectedCards()
    -- ... add cards code ...
    
    print("Total ships stored: " .. #savedShipCards)
    for i, card in ipairs(savedShipCards) do
        print("Ship " .. i .. ": " .. card.name .. " - Type: " .. card.tonnage)
    end
end
```

### Step 2: Check filter matching
Add debug output in getFilteredCards:
```lua
function getFilteredCards(activeTab, searchFilter)
    print("Filtering for tab: " .. activeTab)
    
    for i, card in ipairs(savedShipCards) do
        local matchesTab = (activeTab == "All") or (card.tonnage == activeTab)
        print("Card: " .. card.name .. " | Tonnage: " .. card.tonnage .. " | Matches: " .. tostring(matchesTab))
    end
    
    -- ... rest of function ...
end
```

### Step 3: Check tab click
Add debug output in uiTabClicked:
```lua
function uiTabClicked(tabName)
    print("Tab clicked: " .. tabName)
    activeTab = tabName
    rebuildCards()
end
```

## Expected Behavior

**Correct Flow:**
1. User clicks "Setup" button
2. User box-selects ship card tiles (e.g., 3 cruiser cards)
3. User clicks "Add Selected Ships"
4. Ships are stored with detected tonnage:
   - Card 1: "Boston Light Cruiser" → tonnage = "Light Cruiser"
   - Card 2: "Madrid Cruiser" → tonnage = "Cruiser"
   - Card 3: "New Cairo Heavy Cruiser" → tonnage = "Heavy Cruiser"
5. User clicks "View Ships" button
6. User clicks "Cruiser" tab
7. Only "Madrid Cruiser" appears (tonnage matches "Cruiser")
8. User clicks "Light Cruiser" tab
9. Only "Boston Light Cruiser" appears (tonnage matches "Light Cruiser")
10. User clicks "All" tab
11. All 3 ships appear

## Common Issues and Solutions

### Issue: "No ships appear in any category tab"
**Likely Cause:** Ships haven't been added via Setup button
**Solution:** Click Setup → Select cards → Add Selected Ships

### Issue: "Ships appear in 'All' but not in specific tabs"
**Likely Cause:** Tonnage values don't match tab names exactly
**Solution:** Check that detection assigns exact tab name (e.g., "Cruiser" not "Cruisers")

### Issue: "All cruisers appear in 'Cruiser' tab, even Light/Heavy Cruisers"
**Likely Cause:** Detection order is wrong - checking "cruiser" before "light cruiser"
**Solution:** Reorder detection to check specific types first

### Issue: "Ships show up in 'Other' tab instead of their type"
**Likely Cause:** Ship description/name doesn't contain expected keyword
**Solution:** Check actual card description - may need to add alternate spellings or keywords
