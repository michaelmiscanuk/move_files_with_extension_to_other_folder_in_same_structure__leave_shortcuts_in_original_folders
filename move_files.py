import os
import shutil
import pythoncom
from win32com.shell import shell

# ─── CONFIG ───────────────────────────────────────────────────────────────────
SOURCE_FOLDER = r"E:\Main\Knowledge Base"
TARGET_FOLDER = r"E:\OneDrive\Knowledge_Base_OneDrive"
DRY_RUN = False  # Set to True to see what would happen without making changes
EXTENSIONS = [".doc", ".docx"]
# ──────────────────────────────────────────────────────────────────────────────


def collect_files(source: str, extensions: list[str]) -> list[tuple[str, str]]:
    """Walk source folder and return list of (src_file_path, relative_path) tuples."""
    matches = []
    for dirpath, _dirnames, filenames in os.walk(source):
        for filename in filenames:
            if os.path.splitext(filename)[1].lower() in extensions:
                full_path = os.path.join(dirpath, filename)
                rel_path = os.path.relpath(full_path, source)
                matches.append((full_path, rel_path))
    return matches


def create_shortcut(shortcut_path: str, target_path: str) -> None:
    """Create a Windows .lnk shortcut pointing to target_path (Unicode-safe)."""
    shortcut = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink,
        None,
        pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IShellLink,
    )
    shortcut.SetPath(target_path)
    persist_file = shortcut.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Save(shortcut_path, 0)


def run(source: str, target: str, dry_run: bool, extensions: list[str]) -> None:
    files = collect_files(source, extensions)

    if not files:
        print("No matching files found.")
        return

    print(f"{'[DRY RUN] ' if dry_run else ''}Found {len(files)} file(s) to move:\n")

    failures = []

    for src_path, rel_path in files:
        dst_path = os.path.join(target, rel_path)
        shortcut_path = os.path.splitext(src_path)[0] + ".lnk"

        print(f"  FILE    : {src_path}")
        print(f"  -> MOVE : {dst_path}")
        print(f"  -> LNK  : {shortcut_path}")
        print()

        if not dry_run:
            try:
                # Create target directory if needed
                os.makedirs(os.path.dirname(dst_path), exist_ok=True)

                # Move the file
                shutil.move(src_path, dst_path)

                # Ensure source directory still exists before creating shortcut
                os.makedirs(os.path.dirname(shortcut_path), exist_ok=True)

                # Create shortcut in original location
                create_shortcut(shortcut_path, dst_path)
            except Exception as e:
                failures.append((src_path, dst_path, shortcut_path, str(e)))
                print(f"  [ERROR] {e}\n")

    if dry_run:
        print("[DRY RUN] No files were moved. Set DRY_RUN = False to execute.")
    else:
        succeeded = len(files) - len(failures)
        print(f"Done. {succeeded}/{len(files)} file(s) moved and shortcut(s) created.")
        if failures:
            print(f"\n{len(failures)} failure(s):")
            for src, dst, lnk, err in failures:
                print(f"  FILE : {src}")
                print(f"  DST  : {dst}")
                print(f"  LNK  : {lnk}")
                print(f"  ERR  : {err}")
                print()


if __name__ == "__main__":
    run(SOURCE_FOLDER, TARGET_FOLDER, DRY_RUN, EXTENSIONS)
