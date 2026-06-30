from __future__ import annotations

import subprocess
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import List, Union
from xml.sax.saxutils import escape

from bartender_label_runner import PROJECT_ROOT, TEMPLATE_ROOT, BarTenderRunnerError, resolve_template_path, validate_runtime_csv
from config import BARTENDER_EXE, BARTENDER_PRINTER


OUTPUT_DIR = PROJECT_ROOT / "final_labels" / "previews"
SCRIPT_DIR = PROJECT_ROOT / "final_labels" / "preview_scripts"
MAX_PREVIEW_MO_DIRS = 20
MAX_PREVIEW_SCRIPT_FILES = 100


def safe_name(value: str) -> str:
    allowed = []
    for char in value:
        if char.isalnum() or char in ("-", "_"):
            allowed.append(char)
        else:
            allowed.append("_")
    return "".join(allowed).strip("_") or "label"


def clear_directory(path: Union[str, Path]) -> None:
    directory = Path(path)
    if not directory.exists():
        return
    for child in directory.iterdir():
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()


def cleanup_preview_cache(
    preview_root: Union[str, Path] = OUTPUT_DIR,
    script_dir: Union[str, Path] = SCRIPT_DIR,
    keep_mo_dirs: int = MAX_PREVIEW_MO_DIRS,
    keep_script_files: int = MAX_PREVIEW_SCRIPT_FILES,
) -> None:
    root = Path(preview_root)
    if root.exists():
        mo_dirs = [path for path in root.iterdir() if path.is_dir()]
        mo_dirs.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        for old_dir in mo_dirs[max(keep_mo_dirs, 0):]:
            shutil.rmtree(old_dir)

        direct_files = [path for path in root.iterdir() if path.is_file()]
        for file_path in direct_files:
            file_path.unlink()

    scripts = Path(script_dir)
    if scripts.exists():
        script_files = [path for path in scripts.iterdir() if path.is_file()]
        script_files.sort(key=lambda path: path.stat().st_mtime, reverse=True)
        for old_file in script_files[max(keep_script_files, 0):]:
            old_file.unlink()


def build_preview_xml(
    template_path: Path,
    output_dir: Path,
    file_name_template: str,
    printer_name: str = "",
    image_format: str = "PNG",
    dpi: int = 300,
) -> str:
    printer_xml = f"\n        <Printer>{escape(printer_name)}</Printer>" if printer_name else ""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<XMLScript Version="2.0">
  <Command Name="ForceCopies">
    <FormatSetup>
      <Format CloseAtEndOfJob="false">{escape(str(template_path))}</Format>
      <PrintSetup>
        <NumberSerializedLabels>1</NumberSerializedLabels>
        <IdenticalCopiesOfLabel>1</IdenticalCopiesOfLabel>{printer_xml}
      </PrintSetup>
    </FormatSetup>
  </Command>
  <Command Name="ExportPreview">
    <ExportPrintPreviewToImage ReturnImageInResponse="false">
      <Folder>{escape(str(output_dir))}</Folder>
      <FileNameTemplate>{escape(file_name_template)}</FileNameTemplate>
      <ImageFormatType>{escape(image_format)}</ImageFormatType>
      <Colors>btColors24Bit</Colors>
      <DPI>{dpi}</DPI>
      <Overwrite>true</Overwrite>
      <IncludeMargins>false</IncludeMargins>
      <IncludeBorder>true</IncludeBorder>
      <Format CloseAtEndOfJob="true">{escape(str(template_path))}</Format>
    </ExportPrintPreviewToImage>
  </Command>
</XMLScript>
'''


def collect_preview_images(out_dir: Path, base_name: str, image_format: str) -> List[Path]:
    generated = sorted(
        path
        for path in out_dir.iterdir()
        if path.is_file() and path.name.startswith(base_name)
    )
    normalized: List[Path] = []
    for path in generated:
        if path.suffix:
            normalized.append(path)
            continue

        renamed = path.with_name(path.name + f".{image_format.lower()}")
        if renamed.exists():
            renamed.unlink()
        path.rename(renamed)
        normalized.append(renamed)
    return normalized


def wait_for_preview_images(out_dir: Path, base_name: str, image_format: str, timeout_seconds: int) -> List[Path]:
    deadline = time.time() + timeout_seconds
    while True:
        images = collect_preview_images(out_dir, base_name, image_format)
        if images:
            return images
        if time.time() >= deadline:
            return []
        time.sleep(0.5)


def generate_preview_image(
    template_path: Union[str, Path],
    csv_path: Union[str, Path],
    label_name: str,
    output_dir: Union[str, Path] = OUTPUT_DIR,
    script_dir: Union[str, Path] = SCRIPT_DIR,
    bartender_exe: Union[str, Path] = BARTENDER_EXE,
    printer_name: str = BARTENDER_PRINTER,
    image_format: str = "PNG",
    dpi: int = 300,
    run: bool = True,
    preview_timeout_seconds: int = 8,
) -> List[Path]:
    template = resolve_template_path(template_path, TEMPLATE_ROOT)
    if not template.exists():
        raise BarTenderRunnerError(f"模板文件不存在：{template}")
    validate_runtime_csv(csv_path)

    out_dir = Path(output_dir).resolve()
    xml_dir = Path(script_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    xml_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"{timestamp}_{safe_name(label_name)}"
    expected_image = out_dir / f"{base_name}.{image_format.lower()}"
    suffix = 1
    while expected_image.exists():
        base_name = f"{timestamp}_{safe_name(label_name)}_{suffix}"
        expected_image = out_dir / f"{base_name}.{image_format.lower()}"
        suffix += 1

    script_path = xml_dir / f"{base_name}.xml"
    script_path.write_text(
        build_preview_xml(
            template_path=template,
            output_dir=out_dir,
            file_name_template=base_name,
            printer_name=printer_name,
            image_format=image_format,
            dpi=dpi,
        ),
        encoding="utf-8",
    )

    if run:
        command = [str(bartender_exe), f"/XMLSCRIPT={script_path}", f"/D={Path(csv_path).resolve()}", "/X"]
        subprocess.run(command, check=True)

    if not run:
        return [expected_image]

    normalized = wait_for_preview_images(out_dir, base_name, image_format, preview_timeout_seconds)

    if not normalized:
        printer_hint = f"；当前指定打印机：{printer_name}" if printer_name else "；当前未指定打印机"
        raise BarTenderRunnerError(
            f"未生成预览图：{expected_image}{printer_hint}。"
            "如果清空打印机后可以生成，说明该打印机驱动/名称与 BarTender 模板不兼容。"
        )

    expected_count = len(validate_runtime_csv(csv_path))
    if len(normalized) > expected_count:
        raise BarTenderRunnerError(
            f"预览页数异常：CSV={expected_count} 行，预览={len(normalized)} 页。"
            "正常情况下预览页数不会大于标签行数；请检查模板副本数或页面设置。"
        )

    return normalized
