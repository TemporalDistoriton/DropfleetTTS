# Button Size and Color Fix

## Issue
Previous XML UI implementation created massive buttons (width=750, height=220) that overwhelmed the interface and didn't function properly.

## Root Cause
Used standalone XML Panel approach with incorrect dimensions. TTS UI units don't scale the same way as expected, leading to oversized buttons.

## Solution
Reverted to TTS's native `self.UI.setXmlTable()` approach with proper button dimensions.

## Button Specifications

### Correct Dimensions
- **Width**: 140 pixels
- **Height**: 25 pixels  
- **Position**: Above tile with z-offset of -5
- **Rotation**: '0 0 180' (flipped to face down)

### Button Layout
```
Y-Position    Button Name          Color
─────────────────────────────────────────
  110        Spawn Ship           Grey (original)
   60        Spawn from List      Green
   10        View Ships           Faction-specific
  -40        Setup                Dark Grey
```

## Faction-Specific View Ships Colors

| Faction | Color Name | Hex Code  |
|---------|-----------|-----------|
| UCM     | GREEN     | #4CAF50   |
| PHR     | GOLD      | #FFD700   |
| BIO     | RED       | #F44336   |
| SCO     | Purple    | #9C27B0   |
| RES     | BLUE      | #2196F3   |
| SHA     | ORANGE    | #FF9800   |
| CIV     | Grey      | #808080   |
| IND     | Dark Grey | #505050   |

## Implementation Method

### setXmlTable Approach (CORRECT)
```lua
function rebuildUI()
    local ui = {
        {tag='Defaults', children={...}},
        {tag='button', attributes={
            onClick='buttonClick_viewShips',
            text='View Ships',
            colors='#4CAF50FF|#66BB6AFF|#2E7D32FF|#1B5E20FF',  -- UCM Green
            width='140',
            height='25',
            position='0 10 -5',
            rotation='0 0 180'
        }}
    }
    self.UI.setXmlTable(ui)
end
```

### Standalone Panel Approach (INCORRECT - caused massive buttons)
```xml
<!-- DON'T USE THIS APPROACH -->
<Panel position="0 0 -400" width="850" height="800">
    <Button width="750" height="220">  <!-- TOO BIG! -->
        ...
    </Button>
</Panel>
```

## Why setXmlTable Works Better

1. **Consistent with existing code**: All faction tiles already used this approach
2. **Proper scaling**: TTS handles dimensions correctly
3. **Reliable positioning**: Buttons appear in expected locations
4. **Native TTS behavior**: Works seamlessly with TTS UI system

## Color Format

TTS buttons use 4 colors in the format:
```
colors='normal|highlighted|pressed|disabled'
```

With alpha channel:
```
colors='#4CAF50FF|#66BB6AFF|#2E7D32FF|#1B5E20FF'
         ↑ Normal  ↑ Hover    ↑ Press   ↑ Disabled
```

## Testing Results

✅ All buttons properly sized (not massive)
✅ All buttons functional (onClick handlers work)
✅ Faction colors correctly applied
✅ Proper spacing between buttons
✅ No overlap or visual issues
✅ List spawner functionality unchanged

## Files Changed
- `TS_Save_13540.json` - All 8 faction tiles updated

## Commit
Fixed in commit a4c4d01
