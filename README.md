# Move Files & Leave Shortcuts

Moves files with specified extensions from a source folder to a target folder, preserving the directory structure. A `.lnk` shortcut pointing to the new location is left in place of each moved file.

**Use case:** offload files to a secondary location (e.g. OneDrive) while keeping transparent access from the original folder tree.

## How it works

1. Recursively scans `SOURCE_FOLDER` for files matching `EXTENSIONS`
2. Moves each file to the mirrored path under `TARGET_FOLDER`
3. Creates a Windows shortcut (`.lnk`) at the original file location pointing to the new path

## Configuration

Edit the constants at the top of `move_files.py`:

```python
SOURCE_FOLDER = r"E:\Main\Knowledge Base"
TARGET_FOLDER = r"E:\OneDrive\Knowledge_Base_OneDrive"
DRY_RUN      = False        # True = preview only, no files moved
EXTENSIONS   = [".doc", ".docx"]
```

## Requirements

- Windows
- Python 3.10+
- [`pywin32`](https://pypi.org/project/pywin32/)

```
pip install pywin32
```

## Usage

```
python move_files.py
```

Set `DRY_RUN = True` first to preview what will be moved without making any changes.

## Notes

- Shortcut creation uses the native `IShellLink` COM interface, which correctly handles Unicode characters in paths.  
- If an individual file fails, the error is printed and the script continues. A summary of all failures is shown at the end.
