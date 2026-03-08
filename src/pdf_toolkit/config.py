from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import tomllib


@dataclass(slots=True)
class ToolkitConfig:
    default_output_root: Path | None = None
    report_format: str = "json"
    ocr_language: str = "eng"
    temp_dir: Path = Path("tmp/pdfs")
    overwrite: bool = False


def load_config(start_dir: Path | None = None) -> ToolkitConfig:
    search_dir = (start_dir or Path.cwd()).resolve()
    config_path = search_dir / "pdf-toolkit.toml"
    if not config_path.exists():
        return ToolkitConfig()

    data = tomllib.loads(config_path.read_text(encoding="utf-8"))
    raw = data.get("toolkit", {})
    output_root = raw.get("default_output_root")
    temp_dir = raw.get("temp_dir", "tmp/pdfs")
    return ToolkitConfig(
        default_output_root=Path(output_root) if output_root else None,
        report_format=str(raw.get("report_format", "json")),
        ocr_language=str(raw.get("ocr_language", "eng")),
        temp_dir=Path(temp_dir),
        overwrite=bool(raw.get("overwrite", False)),
    )


def resolve_path(path: Path | None, config: ToolkitConfig) -> Path | None:
    if path is None:
        return None
    if path.is_absolute() or config.default_output_root is None:
        return path
    return config.default_output_root / path
