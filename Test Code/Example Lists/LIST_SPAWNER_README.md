# Fleet List Spawner Feature

## Overview
All faction tiles in this save file now support spawning ship cards directly from fleet lists. This feature allows you to paste a fleet list and automatically spawn all the ships in your player area.

## How to Use

1. **Click the "Spawn from List" button** on any faction tile
   - A Global UI panel will appear in your view
2. **Paste your fleet list** into the text input field in the panel
   - The panel preserves all formatting (spaces, newlines, bullets)
3. **Click "Spawn Ships"** button in the panel
4. Ships will spawn automatically in your player area
5. The panel will auto-close after spawning

**Note:** Click "Spawn from List" again to toggle the panel on/off, or click "Close" in the panel.

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
- **Global UI panel**: Floating window that appears when you click "Spawn from List"
- **Full text preservation**: XML InputField preserves spaces, newlines, and all special characters
- **Multi-line support**: Properly handles fleet lists with multiple lines
- **Automatic quantity detection**: Lines with "2x" or "3x" will spawn multiple copies
- **Default quantity**: If no number is specified, 1 ship is spawned
- **Smart filtering**: Headers, admirals, and configuration lines are automatically skipped
- **Fuzzy matching**: Ship names are matched flexibly (e.g., "Johannesburg Battlecruiser" matches "Johannesburg")
- **Auto-hide**: Panel automatically closes after successful spawning

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

### Text appears without spaces or newlines
- This issue has been fixed! The new Global XML UI properly preserves all formatting.
- Make sure you're using the latest version (commit 2c99fd3 or later)

### Panel not appearing
- Make sure you clicked "Spawn from List" button on the faction tile
- Check that you're looking at your screen (panel appears in Global UI space)
- Try clicking the button again to toggle

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
