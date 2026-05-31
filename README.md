# Knowledge Base — OneDrive Sync

Keeps `E:\Main\Knowledge Base` and `E:\OneDrive\Knowledge Base` in a clean split:
OneDrive holds **only Word documents**; everything else lives locally.
A `.lnk` shortcut in the local folder provides transparent access to each Word doc in OneDrive.

## Main script

**`sync_folders__in_onedrive_keep_only_word_docs.py`** — run this to sync both directions.

It enforces the invariant in three jobs, all folder-structure-preserving:

| Job | What it does |
|-----|--------------|
| **JOB 1** OneDrive → Local | Moves any non-Word file from OneDrive back to the matching local path |
| **JOB 2** Local → OneDrive | Moves any Word doc found locally to OneDrive; leaves a `.lnk` shortcut behind |
| **JOB 3** Shortcut repair | For every Word doc in OneDrive, creates missing shortcuts and re-points stale ones |

## Configuration

Edit the constants at the top of the script:

```python
LOCAL_FOLDER    = r"E:\Main\Knowledge Base"
ONEDRIVE_FOLDER = r"E:\OneDrive\Knowledge Base"
DRY_RUN         = True          # True = preview only, no files moved
WORD_EXTENSIONS = [".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"]
CONFLICT_POLICY = "skip"        # "skip" | "overwrite" | "rename"
REPAIR_MISMATCHED = True        # re-point shortcuts that point at a wrong path
```

## Usage

```
python sync_folders__in_onedrive_keep_only_word_docs.py
```

Always run with `DRY_RUN = True` first to review what will change.  
Set `DRY_RUN = False` to apply — the script will ask for `yes` confirmation before doing anything.

Every run (dry or live) writes categorized TXT reports to a timestamped folder under `reports/`:

| Report file | Contents |
|-------------|----------|
| `00_SUMMARY.txt` | Header, counts, and every category listed in full |
| `01_j1_moved.txt` | Non-Word files moved OneDrive → Local |
| `02_j1_skipped.txt` | Files skipped due to conflicts |
| `03_j2_moved.txt` | Word docs moved Local → OneDrive |
| `04_j2_shortcut_created.txt` | Shortcuts created for newly-moved Word docs |
| `05_j3_shortcut_created.txt` | Shortcuts created (were missing) |
| `06_j3_shortcut_repointed.txt` | Shortcuts re-pointed (were stale) |
| `07_errors.txt` | Any errors encountered |

The `reports/` folder is git-ignored.

## Requirements

- Windows
- Python 3.10+
- [`pywin32`](https://pypi.org/project/pywin32/) — `pip install pywin32`

## Other scripts

`move_files.py` and `restore_nonword_files_and_repair_shortcuts.py` are the original
one-direction scripts that `sync_folders__in_onedrive_keep_only_word_docs.py` supersedes.
They are kept for reference.
