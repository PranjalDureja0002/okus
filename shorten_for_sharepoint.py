"""
shorten_for_sharepoint.py

Makes a COPY of your folder with shortened folder/file names so it uploads to
Teams / SharePoint / OneDrive without "path too long" errors.

- Your original folder is NOT changed.
- A "_File name index.csv" is placed inside the copy, mapping every short name
  back to its original full name, so nothing is lost.
- No logins, no client ID, no admin rights. Standard library only (Python 3.8+).

How to run (Command Prompt or PowerShell):
    python shorten_for_sharepoint.py "C:\\Work\\Scoping Templates"

Optional settings:
    python shorten_for_sharepoint.py "C:\\Work\\Scoping Templates" --dest "C:\\ST_Upload" --max-folder 30 --max-file 50

Then upload the new folder (default C:\\ST_Upload\\<your folder name>) to the
Teams channel: Files tab -> Upload -> Folder, or drag and drop it in.
"""

import argparse
import csv
import os
import re
import shutil
import sys

BAD_CHARS = re.compile(r'[~"#%&*:<>?/\\{|}]')


def long_path(p: str) -> str:
    """Add the \\\\?\\ prefix on Windows so paths over 260 characters still work."""
    if os.name != "nt":
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def short_name(name: str, max_len: int, is_file: bool, used: set) -> str:
    """Tidy and truncate a name, keeping the extension and avoiding duplicates."""
    if is_file:
        base, ext = os.path.splitext(name)
    else:
        base, ext = name, ""

    base = BAD_CHARS.sub("", re.sub(r"\s+", " ", base)).strip().rstrip(".")
    if not base:
        base = "Item"

    room = max(max_len - len(ext), 5)
    if len(base) > room:
        base = base[:room].strip().rstrip(".")

    candidate = base + ext
    n = 2
    while candidate.lower() in used:  # Windows/SharePoint names are case-insensitive
        suffix = f" ({n})"
        b = base[: room - len(suffix)].strip() if len(base) + len(suffix) > room else base
        candidate = f"{b}{suffix}{ext}"
        n += 1
    used.add(candidate.lower())
    return candidate


def copy_tree(src, dst, rel_orig, rel_new, args, mapping, stats):
    used = set()
    with os.scandir(long_path(src)) as it:
        entries = sorted(it, key=lambda e: (not e.is_dir(follow_symlinks=False), e.name.lower()))

    for e in entries:
        is_dir = e.is_dir(follow_symlinks=False)
        new = short_name(e.name, args.max_folder if is_dir else args.max_file, not is_dir, used)
        o_rel = os.path.join(rel_orig, e.name) if rel_orig else e.name
        n_rel = os.path.join(rel_new, new) if rel_new else new
        renamed = new != e.name
        if renamed:
            stats["renamed"] += 1

        src_path = os.path.join(src, e.name)
        dst_path = os.path.join(dst, new)

        if is_dir:
            os.makedirs(long_path(dst_path), exist_ok=True)
            mapping.append(("Folder", n_rel, o_rel, renamed))
            copy_tree(src_path, dst_path, o_rel, n_rel, args, mapping, stats)
        else:
            try:
                shutil.copy2(long_path(src_path), long_path(dst_path))
                stats["files"] += 1
                mapping.append(("File", n_rel, o_rel, renamed))
            except OSError as err:
                stats["errors"].append((o_rel, str(err)))


def main():
    ap = argparse.ArgumentParser(description="Copy a folder with SharePoint-safe short names.")
    ap.add_argument("source", help="Your main folder, e.g. C:\\Work\\Scoping Templates")
    ap.add_argument("--dest", default=r"C:\ST_Upload", help="Where the shortened copy is made (keep it short)")
    ap.add_argument("--max-folder", type=int, default=40, help="Max characters per folder name")
    ap.add_argument("--max-file", type=int, default=70, help="Max characters per file name incl. extension")
    ap.add_argument("--max-path", type=int, default=220, help="Max path length inside the main folder")
    args = ap.parse_args()

    src_root = os.path.abspath(args.source).rstrip("\\/")
    if not os.path.isdir(long_path(src_root)):
        sys.exit(f"Folder not found: {src_root}")

    root_name = short_name(os.path.basename(src_root), args.max_folder, False, set())
    dst_root = os.path.join(args.dest, root_name)
    if os.path.exists(long_path(dst_root)):
        sys.exit(f"{dst_root} already exists. Delete or rename it, then run again.")
    os.makedirs(long_path(dst_root))

    mapping, stats = [], {"files": 0, "renamed": 0, "errors": []}
    copy_tree(src_root, dst_root, "", "", args, mapping, stats)

    index_path = os.path.join(dst_root, "_File name index.csv")
    with open(long_path(index_path), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Type", "NewPath", "OriginalPath", "Renamed"])
        w.writerows(mapping)

    too_long = [m[1] for m in mapping if len(os.path.join(root_name, m[1])) > args.max_path]

    print(f"\nDone. Copied {stats['files']} files into {dst_root}")
    print(f"{stats['renamed']} names were shortened. See '_File name index.csv' inside the folder.")
    if stats["errors"]:
        print(f"\nWARNING: {len(stats['errors'])} files could not be copied:")
        for path, err in stats["errors"][:15]:
            print(f"  {path}\n    -> {err}")
    if too_long:
        print(f"\nWARNING: {len(too_long)} paths are still longer than {args.max_path} characters:")
        for p in too_long[:15]:
            print(f"  {p}")
        print(f"Delete {dst_root} and run again with smaller --max-folder / --max-file.")
    elif not stats["errors"]:
        print("All paths are within limits. Upload the folder to the Teams channel (Files -> Upload -> Folder).")


if __name__ == "__main__":
    main()
