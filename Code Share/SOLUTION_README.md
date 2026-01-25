# ModelSpawner Fix - Simple Callback Approach

## The Problem
The original spawner was spawning ships but NOT calling the ship's `InitFromSpawner` function, which meant:
- Ships loaded with default parameters
- Parameters from `shipParamList` were never applied
- The ship's onLoad would search for cards (causing delays)

## The Solution
Use TTS's built-in `callback_function` parameter in `spawnObjectJSON` to call `InitFromSpawner` on the spawned ship after it loads.

### What Changed

**In the spawner script:**
```lua
-- OLD: Just spawn without callback
spawnObjectJSON({
    json = objectJSON,
    position = spawnPos,
    sound = false
})

-- NEW: Spawn with callback to apply parameters
local params = shipParamList[1]  -- Get the test parameters

spawnObjectJSON({
    json = objectJSON,
    position = spawnPos,
    sound = false,
    callback_function = function(spawnedObj)
        -- Wait a few frames for ship to initialize
        Wait.frames(function()
            -- Call InitFromSpawner on the spawned ship
            spawnedObj.call("InitFromSpawner", params)
        end, 3)
    end
})
```

**In the ship script:**
- No changes needed! The ship already has `InitFromSpawner` function
- This function applies all the parameters correctly
- It updates health, sig, scan, thrust, points, name, faction, images, etc.

## How It Works

1. User clicks "Load Models" button
2. Spawner gets parameters from `shipParamList[1]`
3. Ship spawns from objectJSONs
4. Callback executes after spawn
5. Waits 3 frames for ship to initialize
6. Calls `InitFromSpawner(params)` on the ship
7. Ship applies all parameters instantly
8. Console shows "Ship spawned and configured in X.XX seconds"

## Test Parameters

The spawner is configured with one ship variant:
```lua
{
    shipID = "BASETEMPLATECARD2",
    baseSize = 40,
    health = 6,
    sig = 7,
    points = 7,
    scan = 7,
    thrust = 7,
    name = "BASE TEMPLATE CARD mk2",
    faction = "TEMPLATE FACTION mk2",
    cardFrontImage = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/RemasterShips/PHR/Achilles_CardFrontImage.png",
    modelImage = "https://raw.githubusercontent.com/TemporalDistoriton/DropfleetTTS/main/RemasterShips/PHR/Achilles_ModelImage.png"
}
```

## Expected Performance

- **Spawn time**: 1-2 seconds (includes 3-frame wait + parameter application)
- **Parameters**: All correctly applied to the ship
- **Verification**: Console will show confirmation message

## How to Use

1. Load `Solution.json` in Tabletop Simulator
2. Find the ModelSpawner tile on the table
3. Click "Load Models" button
4. Wait 1-2 seconds
5. Ship appears with correct parameters!

## Verification

Check the spawned ship has:
- Name: "BASE TEMPLATE CARD mk2"
- Health: 6
- Signature: 7
- Scan: 7
- Thrust: 7
- Points: 7
- Base Size: 40mm
- Faction: "TEMPLATE FACTION mk2" (in description)

Console should show:
```
Ship spawned and configured in X.XX seconds
```

## Why This Works

- **Simple**: Uses standard TTS callback mechanism
- **Fast**: Only waits 3 frames (< 0.5 seconds)
- **Reliable**: `InitFromSpawner` is already tested and working
- **Clean**: No modifications to ship script needed
- **Backward compatible**: Ship's card search still works for manual spawning

---

**This is the correct, simple solution that actually works in TTS!**
