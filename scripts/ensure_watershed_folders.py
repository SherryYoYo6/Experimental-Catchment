#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ensure every watershed in Literature.xlsx has a library folder (with .gitkeep)."""

from __future__ import annotations

import shutil
from pathlib import Path

import pandas as pd

from build_literature_library import (
    INPUT_FILE,
    LIBRARY_ROOT,
    GITKEEP,
    build_watershed_folder_map,
    ensure_watershed_folder,
)

WORKSPACE = Path("/workspace")
CUMULATIVE_LOG = WORKSPACE / "PDF_Download_Log_cumulative.xlsx"


def migrate_pdfs_to_mapped_folders(folder_map: dict[str, str]) -> int:
    """Move PDFs into folders matching PDF_Library_Folder; update cumulative log paths."""
    if not CUMULATIVE_LOG.exists():
        return 0
    log = pd.read_excel(CUMULATIVE_LOG)
    moved = 0
    for i, row in log.iterrows():
        if row.get("Download_Result") != "Downloaded":
            continue
        path_str = str(row.get("PDF_File_Path", "") or "")
        if not path_str or path_str == "nan":
            continue
        pdf_path = WORKSPACE / path_str
        if not pdf_path.exists():
            matches = list(LIBRARY_ROOT.rglob(pdf_path.name))
            pdf_path = matches[0] if matches else None
        if pdf_path is None or not pdf_path.exists():
            continue
        ws = str(row["Watershed"])
        correct_dir = LIBRARY_ROOT / folder_map[ws]
        correct_dir.mkdir(parents=True, exist_ok=True)
        dest = correct_dir / pdf_path.name
        if pdf_path.resolve() != dest.resolve():
            if dest.exists():
                pdf_path.unlink()
            else:
                shutil.move(str(pdf_path), str(dest))
                moved += 1
            log.at[i, "PDF_File_Path"] = str(dest.relative_to(WORKSPACE))
        ensure_watershed_folder(folder_map[ws])
    log.to_excel(CUMULATIVE_LOG, index=False)
    return moved


def remove_orphan_empty_folders(valid_folders: set[str]) -> int:
    """Remove legacy folders that are empty except .gitkeep and not in the current map."""
    removed = 0
    if not LIBRARY_ROOT.exists():
        return 0
    for child in list(LIBRARY_ROOT.iterdir()):
        if not child.is_dir() or child.name in valid_folders:
            continue
        contents = [p for p in child.iterdir() if p.name != GITKEEP]
        if not contents:
            (child / GITKEEP).unlink(missing_ok=True)
            child.rmdir()
            removed += 1
    return removed


def main() -> None:
    df = pd.read_excel(INPUT_FILE)
    folder_map = build_watershed_folder_map(df["Watershed"].astype(str).tolist())
    for folder_name in folder_map.values():
        ensure_watershed_folder(folder_name)
    # .gitkeep in folders that already contain PDFs
    for pdf in LIBRARY_ROOT.rglob("*.pdf"):
        ensure_watershed_folder(pdf.parent.name)
    moved = migrate_pdfs_to_mapped_folders(folder_map)
    removed = remove_orphan_empty_folders(set(folder_map.values()))
    df["PDF_Library_Folder"] = df["Watershed"].astype(str).map(folder_map)
    df.to_excel(INPUT_FILE, index=False)
    n = len(set(folder_map.values()))
    with_pdf = sum(1 for f in set(folder_map.values()) if list((LIBRARY_ROOT / f).glob("*.pdf")))
    empty = n - with_pdf
    print(f"Folders: {n} watershed folders ({with_pdf} with PDFs, {empty} empty)")
    print(f"All folders include {GITKEEP} for git tracking")
    print(f"Moved {moved} PDFs; removed {removed} orphan empty folders")
    print(f"Updated {INPUT_FILE} (PDF_Library_Folder column)")


if __name__ == "__main__":
    main()
