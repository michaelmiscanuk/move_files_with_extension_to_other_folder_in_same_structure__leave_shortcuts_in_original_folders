"""
Sync Knowledge Base so OneDrive holds ONLY Word documents.

Invariant after every run
--------------------------
  E:\\Main\\Knowledge Base   – contains everything EXCEPT Word docs
                               + a .lnk shortcut for every Word doc in OneDrive
  E:\\OneDrive\\Knowledge Base – contains ONLY Word docs (no other file types)

Three jobs are run in order, all folder-structure-preserving:

  JOB 1  Non-Word files OneDrive → Local
         Any file in OneDrive that is not a Word doc (and not a .lnk) is moved
         back to the matching location in Main.  Word docs are left alone.
         Conflict policy (skip / overwrite / rename) governs collisions.

  JOB 2  Word docs Local → OneDrive
         Any real Word doc found in Main (not a .lnk stub) is moved to the
         matching location in OneDrive, and a .lnk shortcut is left behind in
         Main pointing at it.

  JOB 3  Shortcut repair
         For every Word doc currently in OneDrive, the corresponding .lnk in
         Main is checked.  Missing shortcuts are created; stale ones (pointing
         at an old path, e.g. after a folder rename) are re-pointed.

DRY_RUN = True  previews all three jobs without touching disk.
Every run writes categorized TXT reports to the reports/ folder.
"""

import os
import sys
import shutil
import datetime
import pythoncom
from win32com.shell import shell

# Windows consoles default to cp1252, which crashes on non-Latin filenames.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ─── CONFIG ───────────────────────────────────────────────────────────────────
LOCAL_FOLDER = r"E:\Main\Knowledge Base"  # Main: non-Word files + shortcuts
ONEDRIVE_FOLDER = r"E:\OneDrive\Knowledge Base"  # OneDrive: Word docs only

DRY_RUN = False  # Set False to apply changes (will prompt for confirmation)

WORD_EXTENSIONS = [".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"]

# Collision policy when a non-Word file already exists at the local destination:
#   "skip"      – leave the OneDrive copy in place and warn  (safest, default)
#   "overwrite" – replace the local file
#   "rename"    – keep both with a numeric suffix, e.g. "file (1).pdf"
CONFLICT_POLICY = "skip"

# Re-point shortcuts that exist but point at the wrong path (e.g. after a
# folder rename).  Set False to leave mismatched shortcuts untouched.
REPAIR_MISMATCHED = True

REPORTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports")
# ──────────────────────────────────────────────────────────────────────────────


# ─── REPORT ───────────────────────────────────────────────────────────────────
class Report:
    """Accumulates per-category result lines and writes them to timestamped TXT files."""

    CATEGORIES = {
        "j1_moved": "JOB1  Non-Word files moved OneDrive -> Local",
        "j1_skipped": "JOB1  Non-Word files skipped (destination already exists)",
        "j2_moved": "JOB2  Word docs moved Local -> OneDrive",
        "j2_shortcut_created": "JOB2  Shortcuts created for newly-moved Word docs",
        "j3_shortcut_created": "JOB3  Shortcuts created (were missing)",
        "j3_shortcut_repointed": "JOB3  Shortcuts re-pointed (were stale)",
        "j3_shortcut_ok": "JOB3  Shortcuts already correct (no action)",
        "j3_shortcut_mismatched": "JOB3  Shortcuts pointing elsewhere (left untouched)",
        "errors": "Errors",
    }

    def __init__(self, reports_dir: str, dry_run: bool):
        self.lines = {k: [] for k in self.CATEGORIES}
        stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        mode = "DRYRUN" if dry_run else "LIVE"
        self.run_dir = os.path.join(reports_dir, f"{stamp}_{mode}")

    def add(self, category: str, line: str) -> None:
        self.lines[category].append(line)

    def count(self, category: str) -> int:
        return len(self.lines[category])

    def write(self, header_lines: list) -> tuple:
        os.makedirs(self.run_dir, exist_ok=True)

        master_path = os.path.join(self.run_dir, "00_SUMMARY.txt")
        with open(master_path, "w", encoding="utf-8") as f:
            f.write("\n".join(header_lines) + "\n")
            f.write("\n=== COUNTS ===\n")
            for key, title in self.CATEGORIES.items():
                f.write(f"{self.count(key):>7}  {title}\n")
            for key, title in self.CATEGORIES.items():
                f.write(f"\n\n=== {title} ({self.count(key)}) ===\n")
                f.write(
                    "\n".join(self.lines[key]) + "\n" if self.lines[key] else "(none)\n"
                )

        written = [master_path]
        for i, (key, title) in enumerate(self.CATEGORIES.items(), start=1):
            if not self.lines[key]:
                continue
            path = os.path.join(self.run_dir, f"{i:02d}_{key}.txt")
            with open(path, "w", encoding="utf-8") as f:
                f.write(f"{title}\nCount: {self.count(key)}\n{'=' * 60}\n")
                f.write("\n".join(self.lines[key]) + "\n")
            written.append(path)

        return master_path, written


# ─── SHORTCUT HELPERS ─────────────────────────────────────────────────────────
def create_shortcut(shortcut_path: str, target_path: str) -> None:
    """Create a Windows .lnk file pointing at target_path (Unicode-safe)."""
    lnk = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink,
        None,
        pythoncom.CLSCTX_INPROC_SERVER,
        shell.IID_IShellLink,
    )
    lnk.SetPath(target_path)
    lnk.QueryInterface(pythoncom.IID_IPersistFile).Save(shortcut_path, 0)


def read_shortcut_target(shortcut_path: str):
    """Return the target path of an existing .lnk, or None if unreadable."""
    try:
        lnk = pythoncom.CoCreateInstance(
            shell.CLSID_ShellLink,
            None,
            pythoncom.CLSCTX_INPROC_SERVER,
            shell.IID_IShellLink,
        )
        lnk.QueryInterface(pythoncom.IID_IPersistFile).Load(shortcut_path)
        return lnk.GetPath(shell.SLGP_RAWPATH)[0]
    except Exception:
        return None


def shortcut_stem(filename: str) -> str:
    """Return the .lnk filename for a given Word doc filename (drops the Word extension)."""
    return os.path.splitext(filename)[0] + ".lnk"


# ─── CONFLICT RESOLUTION ──────────────────────────────────────────────────────
def resolve_conflict(dst_path: str) -> str:
    """Return a non-colliding destination path with a numeric suffix."""
    base, ext = os.path.splitext(dst_path)
    n = 1
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


# ─── JOB 1: Non-Word files OneDrive → Local ───────────────────────────────────
def job1_move_nonword_to_local(dry_run: bool, report: Report) -> None:
    print("=== JOB 1: Non-Word files  OneDrive -> Local ===")

    for dirpath, _dirs, files in os.walk(ONEDRIVE_FOLDER):
        rel = os.path.relpath(dirpath, ONEDRIVE_FOLDER)
        local_dir = os.path.join(LOCAL_FOLDER, rel)

        for filename in files:
            if filename.lower().endswith(".lnk"):
                continue
            if os.path.splitext(filename)[1].lower() in WORD_EXTENSIONS:
                continue  # Word docs stay in OneDrive

            src = os.path.join(dirpath, filename)
            dst = os.path.normpath(os.path.join(local_dir, filename))

            if os.path.exists(dst):
                if CONFLICT_POLICY == "skip":
                    report.add("j1_skipped", f"{src}  ->  {dst}  (exists)")
                    continue
                if CONFLICT_POLICY == "rename":
                    dst = resolve_conflict(dst)
                elif CONFLICT_POLICY != "overwrite":
                    report.add(
                        "errors", f"unknown CONFLICT_POLICY '{CONFLICT_POLICY}': {src}"
                    )
                    continue

            if dry_run:
                report.add("j1_moved", f"{src}  ->  {dst}")
            else:
                try:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.move(src, dst)
                    report.add("j1_moved", f"{src}  ->  {dst}")
                except Exception as e:
                    report.add("errors", f"JOB1 MOVE failed: {src}  ->  {dst}  : {e}")

    print(
        f"  moved: {report.count('j1_moved')}   skipped: {report.count('j1_skipped')}"
    )


# ─── JOB 2: Word docs Local → OneDrive ────────────────────────────────────────
def job2_move_worddocs_to_onedrive(dry_run: bool, report: Report) -> None:
    print("=== JOB 2: Word docs  Local -> OneDrive ===")

    for dirpath, _dirs, files in os.walk(LOCAL_FOLDER):
        rel = os.path.relpath(dirpath, LOCAL_FOLDER)
        od_dir = os.path.join(ONEDRIVE_FOLDER, rel)

        for filename in files:
            if filename.lower().endswith(".lnk"):
                continue  # skip existing shortcuts
            if os.path.splitext(filename)[1].lower() not in WORD_EXTENSIONS:
                continue  # only move Word docs

            src = os.path.join(dirpath, filename)
            dst = os.path.normpath(os.path.join(od_dir, filename))
            lnk = os.path.normpath(os.path.join(dirpath, shortcut_stem(filename)))

            if dry_run:
                report.add("j2_moved", f"{src}  ->  {dst}")
                report.add("j2_shortcut_created", f"{lnk}  ->  {dst}")
            else:
                try:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.move(src, dst)
                    os.makedirs(os.path.dirname(lnk), exist_ok=True)
                    create_shortcut(lnk, dst)
                    report.add("j2_moved", f"{src}  ->  {dst}")
                    report.add("j2_shortcut_created", f"{lnk}  ->  {dst}")
                except Exception as e:
                    report.add("errors", f"JOB2 MOVE failed: {src}  ->  {dst}  : {e}")

    print(
        f"  moved: {report.count('j2_moved')}   shortcuts: {report.count('j2_shortcut_created')}"
    )


# ─── JOB 3: Shortcut repair ───────────────────────────────────────────────────
def job3_repair_shortcuts(dry_run: bool, report: Report) -> None:
    print("=== JOB 3: Shortcut repair ===")

    for dirpath, _dirs, files in os.walk(ONEDRIVE_FOLDER):
        rel = os.path.relpath(dirpath, ONEDRIVE_FOLDER)
        local_dir = os.path.join(LOCAL_FOLDER, rel)

        for filename in files:
            if os.path.splitext(filename)[1].lower() not in WORD_EXTENSIONS:
                continue

            word_doc = os.path.join(dirpath, filename)
            lnk = os.path.normpath(os.path.join(local_dir, shortcut_stem(filename)))
            detail = f"{lnk}  ->  {word_doc}"

            if os.path.exists(lnk):
                target = read_shortcut_target(lnk)
                if target and os.path.normcase(
                    os.path.normpath(target)
                ) == os.path.normcase(word_doc):
                    report.add("j3_shortcut_ok", detail)
                    continue

                # Shortcut exists but points elsewhere.
                stale_detail = (
                    f"{lnk}\n    points to : {target}\n    expected  : {word_doc}"
                )
                if not REPAIR_MISMATCHED:
                    report.add("j3_shortcut_mismatched", stale_detail)
                    continue
                if dry_run:
                    report.add("j3_shortcut_repointed", stale_detail)
                else:
                    try:
                        create_shortcut(lnk, word_doc)
                        report.add("j3_shortcut_repointed", stale_detail)
                    except Exception as e:
                        report.add(
                            "errors", f"JOB3 REPOINT failed: {stale_detail}  : {e}"
                        )
                continue

            # Shortcut is missing.
            if dry_run:
                report.add("j3_shortcut_created", detail)
            else:
                try:
                    os.makedirs(os.path.dirname(lnk), exist_ok=True)
                    create_shortcut(lnk, word_doc)
                    report.add("j3_shortcut_created", detail)
                except Exception as e:
                    report.add("errors", f"JOB3 CREATE failed: {detail}  : {e}")

    print(
        f"  created: {report.count('j3_shortcut_created')}   "
        f"re-pointed: {report.count('j3_shortcut_repointed')}   "
        f"ok: {report.count('j3_shortcut_ok')}"
    )


# ─── MAIN ─────────────────────────────────────────────────────────────────────
def run(dry_run: bool) -> None:
    header = [
        "sync_folders__in_onedrive_keep_only_word_docs.py",
        f"Run at:          {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"Mode:            {'DRY RUN (no changes made)' if dry_run else 'LIVE (changes applied)'}",
        f"Local (Main):    {LOCAL_FOLDER}",
        f"OneDrive:        {ONEDRIVE_FOLDER}",
        f"Word extensions: {WORD_EXTENSIONS}",
        f"Conflict policy: {CONFLICT_POLICY}",
        f"Repair stale shortcuts: {REPAIR_MISMATCHED}",
    ]
    for line in header:
        print(line)
    print()

    for label, path in [("Local", LOCAL_FOLDER), ("OneDrive", ONEDRIVE_FOLDER)]:
        if not os.path.isdir(path):
            print(f"ERROR: {label} folder not found: {path}")
            return

    report = Report(REPORTS_DIR, dry_run)

    job1_move_nonword_to_local(dry_run, report)
    job2_move_worddocs_to_onedrive(dry_run, report)
    job3_repair_shortcuts(dry_run, report)

    master_path, written = report.write(header)

    print("\n--- Summary ---")
    if dry_run:
        print("[DRY RUN] No changes were made. Set DRY_RUN = False to execute.")
    for key, title in Report.CATEGORIES.items():
        print(f"{report.count(key):>7}  {title}")
    print(f"\nReports ({len(written)} file(s)): {report.run_dir}")


if __name__ == "__main__":
    if not DRY_RUN:
        confirm = input(
            f"WARNING: This will MOVE files between '{LOCAL_FOLDER}' and "
            f"'{ONEDRIVE_FOLDER}'. Continue? (yes/no): "
        )
        if confirm.lower() != "yes":
            print("Operation cancelled.")
            raise SystemExit(0)
    run(DRY_RUN)
