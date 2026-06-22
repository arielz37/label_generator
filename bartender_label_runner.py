from __future__ import annotations

import argparse
import csv
import json
import subprocess
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Union


PROJECT_ROOT = Path(__file__).resolve().parent
TEMPLATE_ROOT = PROJECT_ROOT / "Templates"
RUNTIME_CSV = PROJECT_ROOT / "runtime_data" / "current_label.csv"
OUTPUT_DIR = PROJECT_ROOT / "final_labels"


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


def resolve_template_path(template_path: Union[str, Path], template_root: Union[str, Path] = TEMPLATE_ROOT) -> Path:
    path = Path(template_path)
    if not path.is_absolute():
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
        command.append("/P")
    command.append("/X")
    return command


def write_label_job_manifest(job: LabelJob, command: List[str]) -> Path:
    job.output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    manifest_path = job.output_dir / f"label_job_{timestamp}.json"
    manifest = {
        "created_at": timestamp,
        "template_path": str(job.template_path),
        "csv_path": str(job.csv_path),
        "printer_name": job.printer_name,
        "print_now": job.print_now,
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
    run: bool = False,
) -> Path:
    job = LabelJob(
        template_path=resolve_template_path(template_path),
        csv_path=Path(csv_path).resolve(),
        output_dir=Path(output_dir).resolve(),
        bartender_exe=bartender_exe,
        printer_name=printer_name,
        print_now=print_now,
    )
    command = build_bartender_command(job)
    manifest_path = write_label_job_manifest(job, command)

    if run:
        subprocess.run(command, check=True)

    return manifest_path


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="用 BarTender 模板和 runtime CSV 生成最终标签任务。")
    parser.add_argument("--template", required=True, help="模板相对路径或绝对路径")
    parser.add_argument("--csv", default=str(RUNTIME_CSV), help="runtime CSV 路径")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="最终标签任务输出目录")
    parser.add_argument("--bartender-exe", default="bartend.exe", help="BarTender 可执行文件路径")
    parser.add_argument("--printer", default="", help="打印机名称")
    parser.add_argument("--print-now", action="store_true", help="生成命令时加入 /P 打印参数")
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
