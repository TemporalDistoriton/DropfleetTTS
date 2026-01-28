# Fleet List Spawner Feature

## Overview
All faction tiles in this save file now support spawning ship cards directly from fleet lists. This feature allows you to paste a fleet list and automatically spawn all the ships in your player area.

## How to Use

1. **Click the "Spawn from List" button** on any faction tile
   - This will show a text input box below the tile
2. **Paste your fleet list** into the text box
3. **Press Enter/Return** to process the list and spawn ships
4. The input box will automatically hide after spawning

**Note:** Click "Spawn from List" again to toggle the input box on/off.

The ships will spawn in a line in your player area, with proper spacing between them.

## Supported List Formats

The parser supports the standard Dropfleet Commander list format used by list builders. Example formats:

```
Johannesburg Battlecruiser [180 pts]
Berlin Cruisers [80 pts]:
• 1x Berlin Cruiser [80 pts]
New Cairo Light Cruisers [140 pts]:
• 2x New Cairo Light Cruiser [70 pts]
```

### Key Features:
- **Toggle input box**: Click "Spawn from List" to show/hide the text input field
- **Text input validation**: Properly configured to accept any text (validation = 5)
- **Automatic quantity detection**: Lines with "2x" or "3x" will spawn multiple copies
- **Default quantity**: If no number is specified, 1 ship is spawned
- **Smart filtering**: Headers, admirals, and configuration lines are automatically skipped
- **Fuzzy matching**: Ship names are matched flexibly (e.g., "Johannesburg Battlecruiser" matches "Johannesburg")
- **Auto-hide**: Input box automatically hides after successful spawning

## Technical Details

### Parser Features
- **Line-by-line processing**: Avoids the "pattern too complex" error by processing text character-by-character
- **Header detection**: Automatically skips section headers like "## Heavy Groups"
- **Bullet point handling**: Supports • bullets and dash (-) prefixes
- **Quantity extraction**: Detects "Nx " patterns to determine spawn count
- **Name matching**: Uses substring matching to find cards in the saved library

### What Gets Spawned
- Only ship cards that match entries in the faction's saved card library
- Ships spawn in the order they appear in the list
- Warnings are displayed for ships that cannot be found

## Example Lists

The following example lists are provided for testing:
- United Colonies of Mankind - UCM Test.txt
- United Colonies of Mankind - Victor.txt
- Post-Human Republic - A wizard did.txt
- Bioficers - Biotime! - [1000 pts].txt
- Scourge - Scourge Rush - [1500 pts].txt
- Shaltari - Pluto - [1515 pts].txt
- Resistance - Resist Test - [751 pts].txt

All of these lists have been tested and parse correctly.

## Troubleshooting

### Input box not accepting text
- The input field is now properly configured with `validation = 5` (unrestricted text)
- You can paste any fleet list text without restrictions
- If you still have issues, click "Spawn from List" to toggle the box off and on again

### Input box is too large or overlapping
- Click "Spawn from List" to hide the input box when not in use
- The box only appears when you need it

### "No ships found in the list"
- Make sure you're pasting a properly formatted fleet list
- Check that the list contains ship entries (not just configuration or headers)

### "Could not find cards for: [ship name]"
- The ship name in the list doesn't match any saved cards in this faction's library
- Check the spelling of the ship name
- Ensure you're using the correct faction tile for your list

### Ships spawn in unexpected quantities
- Verify the quantity prefix in the list (e.g., "2x" for 2 ships)
- Lines without explicit quantities spawn 1 ship by default

## Performance Notes

- The parser uses simple string operations to avoid Lua's "pattern too complex" limitations
- Large lists (100+ ships) may take a few seconds to process
- All text processing is done character-by-character for maximum reliability
