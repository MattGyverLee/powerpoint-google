# PowerPoint Cleaner for Google Slides

A Python script to fix PowerPoint files that fail to import into Google Slides.

## The Problem

Some PowerPoint files (especially those edited in LibreOffice or synced via Dropbox) develop compatibility issues that prevent Google Slides from reading them:

1. **Malformed ZIP structure** - Extra padding bytes between files in the .pptx archive
2. **LibreOffice metadata** - Application metadata that Google Slides rejects

Microsoft PowerPoint can open these files without issue, but Google Slides cannot.

## The Solution

This script rewrites the PowerPoint file with:
- Clean ZIP structure (removes extra padding)
- Microsoft Office-compatible metadata

## Installation

### Requirements
- Python 3.6 or higher (comes pre-installed on most systems)

### Windows Setup

1. **Check if Python is installed:**
   - Open PowerShell or Command Prompt
   - Type: `python --version`
   - If you see a version number (3.6+), you're ready!

2. **If Python is not installed:**
   - Download from: https://www.python.org/downloads/
   - During installation, **check "Add Python to PATH"**

3. **Download the script:**
   - Save `clean_powerpoint.py` to a folder (e.g., `C:\Tools\`)

## Usage

### Option 1: Clean a Single File

**PowerShell:**
```powershell
python clean_powerpoint.py "C:\Users\YourName\Documents\presentation.pptx"
```

This creates `presentation-CLEAN.pptx` in the same folder.

**With custom output name:**
```powershell
python clean_powerpoint.py "presentation.pptx" "fixed-presentation.pptx"
```

### Option 2: Batch Clean Multiple Files

Clean all PowerPoint files in a directory:

```powershell
python clean_powerpoint.py --batch "C:\Users\YourName\Documents\Presentations"
```

Clean all files in the current directory:
```powershell
python clean_powerpoint.py --batch .
```

### Option 3: Drag and Drop (Windows)

Create a batch file for easy drag-and-drop:

1. Create a new text file named `clean_powerpoint.bat`
2. Add this content:
   ```batch
   @echo off
   python "C:\Path\To\clean_powerpoint.py" %1
   pause
   ```
3. Save it
4. Drag a .pptx file onto the .bat file to clean it

## Examples

### Example 1: Basic Usage
```powershell
PS C:\Presentations> python clean_powerpoint.py "Sales Deck.pptx"

Cleaning: Sales Deck.pptx
  ✓ Fixed LibreOffice metadata
  ✓ Original size: 921,947 bytes
  ✓ Cleaned size:  920,505 bytes
  ✓ Removed 1,442 bytes of padding/metadata
  ✓ Saved to: Sales Deck-CLEAN.pptx
```

### Example 2: Batch Processing
```powershell
PS C:\Presentations> python clean_powerpoint.py --batch .

Found 5 PowerPoint file(s)

Cleaning: Presentation1.pptx
  ✓ Fixed LibreOffice metadata
  ✓ Saved to: Presentation1-CLEAN.pptx

Cleaning: Presentation2.pptx
  ✓ Saved to: Presentation2-CLEAN.pptx

...

==================================================
Batch processing complete:
  ✓ Successfully cleaned: 5
```

### Example 3: Quiet Mode (minimal output)
```powershell
python clean_powerpoint.py --batch . --quiet
```

## Command-Line Options

```
python clean_powerpoint.py <input.pptx> [output.pptx]
python clean_powerpoint.py --batch <directory>

Options:
  --batch       Process all .pptx files in a directory
  --quiet, -q   Suppress progress messages
  --help, -h    Show help message
```

## What Gets Fixed

### 1. ZIP Structure Issues
- Removes extra padding bytes between ZIP entries
- Rebuilds the archive with standard structure
- Ensures proper ZIP file format

### 2. Metadata Issues
- Replaces LibreOffice application metadata
- Updates to Microsoft Office PowerPoint metadata
- Preserves all slide content and formatting

## Troubleshooting

### "python: command not found"
- Python is not installed or not in PATH
- Solution: Install Python and make sure "Add to PATH" is checked

### "'python' is not recognized..."
- Try `python3` instead of `python`
- Or use: `py clean_powerpoint.py ...`

### Script runs but file still won't import
- The file may have other issues beyond ZIP/metadata
- Try opening and re-saving in Microsoft PowerPoint
- Check if the file is password-protected

### Permission errors
- Make sure you have write permission in the output directory
- Don't run the script on files that are currently open

## Technical Details

**What this script does:**
1. Opens the .pptx file (which is a ZIP archive)
2. Reads all files inside the archive
3. Detects and fixes LibreOffice metadata in `docProps/app.xml`
4. Writes a new .pptx with clean ZIP structure
5. Verifies the output file integrity

**What this script does NOT do:**
- Modify slide content or formatting
- Change images or media
- Alter animations or transitions
- Edit speaker notes

**File size changes:**
- Typically 1-2KB smaller due to removed padding
- All content remains identical

## Safety

- The script **never modifies** the original file
- It always creates a new file with `-CLEAN` suffix
- Your original presentation is safe
- The cleaned file can be opened in both PowerPoint and Google Slides

## Example Files

This repository includes:
- `Vibe Coding Overview- clean.pptx` - Original problematic file
- `Vibe Coding Overview- CLEAN-FIXED.pptx` - Cleaned version

## License

MIT License - Free to use and modify

## Support

If you encounter issues:
1. Make sure you're using Python 3.6+
2. Verify the input file is a valid .pptx file
3. Check that the file isn't open in another program
4. Try running with `--help` to see usage information

## Author

Created to solve Google Slides import issues with LibreOffice-edited PowerPoint files.
