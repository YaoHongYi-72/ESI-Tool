from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Callable

from analyze_2025_esi import run_analysis
from generate_2025_esi_doc import generate_report_doc
from generate_2025_esi_process_doc import generate_process_doc
from generate_author_flow_doc import generate_author_flow_doc
from generate_esi_stats_workbook import generate_stats_workbook


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _settings_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path.home() / ".config"
    target = base / "ESI统计工具"
    try:
        target.mkdir(parents=True, exist_ok=True)
        return target / "tool_settings.json"
    except OSError:
        fallback = PROJECT_ROOT / "windows_tool"
        fallback.mkdir(parents=True, exist_ok=True)
        return fallback / "tool_settings.json"


SETTINGS_PATH = _settings_path()


@dataclass
class PipelineSettings:
    input_dir: str = ""
    template_docx: str = ""
    workspace_dir: str = ""


def load_settings() -> PipelineSettings:
    if SETTINGS_PATH.exists():
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        return PipelineSettings(**data)
    return PipelineSettings()


def save_settings(settings: PipelineSettings) -> None:
    SETTINGS_PATH.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")


def run_pipeline(settings: PipelineSettings, log: Callable[[str], None]) -> dict[str, Path]:
    input_dir = Path(settings.input_dir).expanduser().resolve()
    template_docx = Path(settings.template_docx).expanduser().resolve()
    workspace_dir = Path(settings.workspace_dir).expanduser().resolve()

    if not input_dir.exists():
        raise FileNotFoundError(f"数据文件夹不存在：{input_dir}")
    if not template_docx.exists():
        raise FileNotFoundError(f"模板报告不存在：{template_docx}")

    analysis_dir = workspace_dir / "analysis"
    outputs_dir = workspace_dir / "outputs"
    analysis_dir.mkdir(parents=True, exist_ok=True)
    outputs_dir.mkdir(parents=True, exist_ok=True)

    stats_xlsx = outputs_dir / "2025年度ESI高水平论文统计结果.xlsx"
    report_docx = outputs_dir / "我校2025年度ESI高水平论文的产出情况及影响力分析.docx"
    process_docx = outputs_dir / "2025年度ESI高水平论文统计口径与核验说明.docx"
    author_flow_docx = outputs_dir / "2025年度ESI高水平论文作者统计流程详解.docx"

    log("== 分析 6 期附表 ==")
    summary = run_analysis(input_dir, analysis_dir, template_docx)
    log(f"年度总量：高水平论文 {summary['totals']['high_level']} 篇，高被引 {summary['totals']['high_cited']} 篇，热点 {summary['totals']['hot']} 篇")
    log("")

    log("== 生成统计工作簿 ==")
    generate_stats_workbook(analysis_dir, stats_xlsx)
    log(str(stats_xlsx))
    log("")

    log("== 生成正式报告 ==")
    generate_report_doc(analysis_dir, template_docx, report_docx)
    log(str(report_docx))
    log("")

    log("== 生成统计口径说明 ==")
    generate_process_doc(analysis_dir, process_docx)
    log(str(process_docx))
    log("")

    log("== 生成作者统计流程文档 ==")
    generate_author_flow_doc(analysis_dir, author_flow_docx)
    log(str(author_flow_docx))
    log("")

    return {
        "analysis_dir": analysis_dir,
        "outputs_dir": outputs_dir,
        "stats_xlsx": stats_xlsx,
        "report_docx": report_docx,
        "process_docx": process_docx,
        "author_flow_docx": author_flow_docx,
    }
