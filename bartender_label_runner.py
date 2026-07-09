from __future__ import annotations

import argparse
import csv
import json
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Union
from xml.sax.saxutils import escape

import app_paths
from app_paths import FINAL_LABELS_DIR, PROJECT_ROOT, RUNTIME_DIR, TEMPLATE_ROOT

RUNTIME_CSV = RUNTIME_DIR / "current_label.csv"
OUTPUT_DIR = FINAL_LABELS_DIR
PRINT_SCRIPT_DIR = OUTPUT_DIR / "print_scripts"


class BarTenderRunnerError(Exception):
    pass


@dataclass(frozen=True)
class LabelJob:
    template_path: Path
    csv_path: Path
    output_dir: Path
    bartender_exe: Union[Path, str] = "bartend.exe"
    printer_name: str = ""
    print_now: bool = False
    copies: int = 1


def resolve_template_path(template_path: Union[str, Path], template_root: Union[str, Path, None] = None) -> Path:
    path = Path(template_path)
    if not path.is_absolute():
        template_root = template_root or app_paths.TEMPLATE_ROOT
        path = Path(template_root) / path
    return path.resolve()


def validate_runtime_csv(csv_path: Union[str, Path]) -> List[Dict[str, str]]:
    path = Path(csv_path).resolve()
    if not path.exists():
        raise BarTenderRunnerError(f"runtime CSV 不存在：{path}")

    with path.open(newline="", encoding="utf-8-sig") as file:
        rows = list(csv.DictReader(file))

    if not rows:
        raise BarTenderRunnerError(f"runtime CSV 没有数据行：{path}")
    return rows


def build_bartender_command(job: LabelJob) -> List[str]:
    if not job.template_path.exists():
        raise BarTenderRunnerError(f"模板文件不存在：{job.template_path}")
    validate_runtime_csv(job.csv_path)

    command = [
        str(job.bartender_exe),
        f'/F="{job.template_path}"',
        f'/D="{job.csv_path}"',
    ]
    if job.printer_name:
        command.append(f'/PRN="{job.printer_name}"')
    if job.print_now:
        command.append(f"/C={job.copies}")
        command.append("/P")
    command.append("/X")
    return command


def build_print_xml(job: LabelJob) -> str:
    printer_xml = f"\n        <Printer>{escape(job.printer_name)}</Printer>" if job.printer_name else ""
    return f'''<?xml version="1.0" encoding="utf-8"?>
<XMLScript Version="2.0">
  <Command Name="PrintLabels">
    <Print>
      <Format CloseAtEndOfJob="true" SaveAtEndOfJob="true">{escape(str(job.template_path))}</Format>
      <PrintSetup>
        <NumberSerializedLabels>1</NumberSerializedLabels>
        <IdenticalCopiesOfLabel>{job.copies}</IdenticalCopiesOfLabel>{printer_xml}
      </PrintSetup>
    </Print>
  </Command>
</XMLScript>
'''


def write_print_xml(job: LabelJob) -> Path:
    PRINT_SCRIPT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    script_path = PRINT_SCRIPT_DIR / f"print_{timestamp}.xml"
    suffix = 1
    while script_path.exists():
        script_path = PRINT_SCRIPT_DIR / f"print_{timestamp}_{suffix}.xml"
        suffix += 1
    script_path.write_text(build_print_xml(job), encoding="utf-8")
    return script_path


def build_bartender_xmlscript_command(job: LabelJob, script_path: Path) -> List[str]:
    return [
        str(job.bartender_exe),
        f"/XMLSCRIPT={script_path}",
        f"/D={job.csv_path}",
        "/X",
    ]


def run_bartender_command(command: List[str]) -> None:
    completed = subprocess.run(command, check=False, capture_output=True, text=True)
    if completed.returncode == 0:
        return

    details = [
        f"BarTender 执行失败，退出码：{completed.returncode}",
        "命令：" + " ".join(command),
    ]
    if completed.stdout:
        details.append("stdout：" + completed.stdout.strip())
    if completed.stderr:
        details.append("stderr：" + completed.stderr.strip())
    raise BarTenderRunnerError("\n".join(details))


def write_label_job_manifest(job: LabelJob, command: List[str], script_path: Union[Path, None] = None) -> Path:
    job.output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    manifest_path = job.output_dir / f"label_job_{timestamp}.json"
    suffix = 1
    while manifest_path.exists():
        manifest_path = job.output_dir / f"label_job_{timestamp}_{suffix}.json"
        suffix += 1
    manifest = {
        "created_at": timestamp,
        "template_path": str(job.template_path),
        "csv_path": str(job.csv_path),
        "printer_name": job.printer_name,
        "print_now": job.print_now,
        "copies": job.copies,
        "save_template_after_print": bool(job.print_now),
        "xmlscript_path": str(script_path) if script_path else "",
        "command": command,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest_path


def generate_final_label_job(
    template_path: Union[str, Path],
    csv_path: Union[str, Path] = RUNTIME_CSV,
    output_dir: Union[str, Path] = OUTPUT_DIR,
    bartender_exe: Union[str, Path] = "bartend.exe",
    printer_name: str = "",
    print_now: bool = False,
    copies: int = 1,
    run: bool = False,
) -> Path:
    job = LabelJob(
        template_path=resolve_template_path(template_path),
        csv_path=Path(csv_path).resolve(),
        output_dir=Path(output_dir).resolve(),
        bartender_exe=bartender_exe,
        printer_name=printer_name,
        print_now=print_now,
        copies=max(int(copies), 1),
    )
    if job.print_now:
        script_path = write_print_xml(job)
        command = build_bartender_xmlscript_command(job, script_path)
    else:
        script_path = None
        command = build_bartender_command(job)
    manifest_path = write_label_job_manifest(job, command, script_path=script_path)

    if run:
        run_bartender_command(command)

    return manifest_path


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="用 BarTender 模板和 runtime CSV 生成最终标签任务。")
    parser.add_argument("--template", required=True, help="模板相对路径或绝对路径")
    parser.add_argument("--csv", default=str(RUNTIME_CSV), help="runtime CSV 路径")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="最终标签任务输出目录")
    parser.add_argument("--bartender-exe", default="bartend.exe", help="BarTender 可执行文件路径")
    parser.add_argument("--printer", default="", help="打印机名称")
    parser.add_argument("--print-now", action="store_true", help="生成命令时加入 /P 打印参数")
    parser.add_argument("--copies", type=int, default=1, help="打印份数")
    parser.add_argument("--run", action="store_true", help="实际执行 BarTender 命令")
    return parser


def main() -> int:
    args = build_arg_parser().parse_args()
    try:
        manifest_path = generate_final_label_job(
            template_path=args.template,
            csv_path=args.csv,
            output_dir=args.output_dir,
            bartender_exe=args.bartender_exe,
            printer_name=args.printer,
            print_now=args.print_now,
            copies=args.copies,
            run=args.run,
        )
    except (BarTenderRunnerError, subprocess.CalledProcessError) as error:
        print("生成失败")
        print(error)
        return 1

    print("生成成功")
    print(f"任务文件：{manifest_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
