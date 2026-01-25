# ModelSpawner Fix - Debugging Version

## Current Status
Added extensive debugging to diagnose why parameters aren't being passed correctly.

## What This Version Does

### Spawner Changes
1. **Inline parameters** - Parameters defined directly in the spawner (not using shipParamList[1])
2. **Extended wait time** - Wait 1 full second before calling InitFromSpawner
3. **Extensive logging** - Prints what parameters are being sent
4. **Debug callback** - Calls DebugOutputParams after initialization

### Ship Changes
1. **DebugOutputParams function** - New function that outputs all current ship parameters
2. **Returns parameter object** - So spawner can verify what values the ship has

## How to Test

1. Load Solution.json in Tabletop Simulator
2. Open the Lua console (` key or ~/` key)
3. Click "Load Models" button on the spawner
4. Watch the console output

## Expected Console Output

```
=== SPAWNER: Preparing to spawn ship ===
Parameters to pass:
  shipID = BASETEMPLATECARD2
  baseSize = 40
  health = 6
  sig = 7
  (... etc ...)

=== SPAWNER CALLBACK: Ship spawned, waiting to initialize ===

=== SPAWNER CALLBACK: Calling InitFromSpawner ===
InitFromSpawner called successfully

=== CHECKING SHIP PARAMETERS ===

=== SHIP DEBUG OUTPUT ===
SHIP_ID: BASETEMPLATECARD2
ShipbaseSize: 40
Shiphealth: 6
Signature: 7
(... etc ...)
=== END DEBUG OUTPUT ===
```

## What to Check

1. **Does InitFromSpawner get called?** - Look for "InitFromSpawner RECEIVED:" in console
2. **Are parameters received by InitFromSpawner?** - Should print all parameter key-value pairs
3. **Do parameters apply?** - DebugOutputParams should show updated values
4. **What does the ship actually have?** - Direct checks of name, description, scale

## Possible Issues

If parameters still don't apply, check:

1. **TTS Lua call() syntax** - Maybe `call()` doesn't work as expected in TTS
2. **Timing issue** - Ship's onLoad might override InitFromSpawner changes
3. **State persistence** - Ship might be loading from save data instead
4. **Function visibility** - InitFromSpawner might not be accessible from spawner

## Next Steps

Based on console output, we can:
- Try different calling methods (setVar, setTable, etc.)
- Adjust timing (wait longer, use different events)
- Modify ship's onLoad to check for spawner flag
- Use alternative parameter passing methods

---

**This is a diagnostic version - not the final solution!**
