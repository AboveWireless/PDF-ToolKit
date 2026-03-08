from __future__ import annotations

from dataclasses import dataclass
from hashlib import sha256
from pathlib import Path

from pdf_toolkit.errors import ValidationError


@dataclass(slots=True)
class DuplicateGroup:
    content_hash: str
    kept_file: Path
    duplicate_files: list[Path]
    file_size: int


def _iter_pdf_files(folder: Path, *, recursive: bool) -> list[Path]:
    if not folder.exists():
        raise ValidationError(f"Folder does not exist: {folder}")
    if not folder.is_dir():
        raise ValidationError(f"Expected a folder path: {folder}")
    iterator = folder.rglob("*.pdf") if recursive else folder.glob("*.pdf")
    return sorted(path for path in iterator if path.is_file())


def _hash_file(path: Path) -> tuple[str, int]:
    digest = sha256()
    size = 0
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def scan_duplicate_pdfs(folder: Path, *, recursive: bool = True) -> dict[str, object]:
    files = _iter_pdf_files(folder, recursive=recursive)
    groups: dict[str, list[Path]] = {}
    sizes: dict[str, int] = {}
    for path in files:
        content_hash, file_size = _hash_file(path)
        groups.setdefault(content_hash, []).append(path)
        sizes[content_hash] = file_size

    duplicate_groups: list[DuplicateGroup] = []
    for content_hash, grouped_paths in sorted(groups.items()):
        if len(grouped_paths) < 2:
            continue
        ordered = sorted(grouped_paths)
        duplicate_groups.append(
            DuplicateGroup(
                content_hash=content_hash,
                kept_file=ordered[0],
                duplicate_files=ordered[1:],
                file_size=sizes[content_hash],
            )
        )

    duplicate_count = sum(len(group.duplicate_files) for group in duplicate_groups)
    return {
        "folder": folder,
        "recursive": recursive,
        "scanned_file_count": len(files),
        "duplicate_group_count": len(duplicate_groups),
        "duplicate_file_count": duplicate_count,
        "groups": duplicate_groups,
    }


def remove_duplicate_pdfs(folder: Path, *, recursive: bool = True, delete_duplicates: bool = False) -> dict[str, object]:
    result = scan_duplicate_pdfs(folder, recursive=recursive)
    removed_files: list[Path] = []
    if delete_duplicates:
        for group in result["groups"]:
            assert isinstance(group, DuplicateGroup)
            for duplicate_path in group.duplicate_files:
                duplicate_path.unlink(missing_ok=False)
                removed_files.append(duplicate_path)
    return {
        "outputs": [],
        "details": {
            "folder": result["folder"],
            "recursive": result["recursive"],
            "scanned_file_count": result["scanned_file_count"],
            "duplicate_group_count": result["duplicate_group_count"],
            "duplicate_file_count": result["duplicate_file_count"],
            "groups": [
                {
                    "content_hash": group.content_hash,
                    "kept_file": str(group.kept_file),
                    "duplicate_files": [str(path) for path in group.duplicate_files],
                    "file_size": group.file_size,
                }
                for group in result["groups"]
            ],
            "removed_files": [str(path) for path in removed_files],
            "removed_count": len(removed_files),
            "delete_duplicates": delete_duplicates,
        },
    }

