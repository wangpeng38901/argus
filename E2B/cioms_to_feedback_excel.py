#!/usr/bin/env python3
"""
将 CIOMS PDF 内容映射到数据反馈 Excel 模板。

功能：
1. 从 URL 或本地路径加载 CIOMS PDF 和 Excel 模板；
2. 判断 PDF 是否为电子档（可提取文本），若疑似扫描件则报错；
3. 解析 CIOMS 常见字段；
4. 通过 JSON/YAML 字段映射配置，将结果写入 Excel 模板并保存输出文件。
"""

from __future__ import annotations

import argparse
import io
import json
import re
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
from pypdf import PdfReader

try:
    import requests  # type: ignore
except Exception:
    requests = None


DEFAULT_PDF_URL = "https://raw.githubusercontent.com/wangpeng38901/argus/main/CIOMS/CIOMS文件.pdf"
DEFAULT_TEMPLATE_URL = (
    "https://raw.githubusercontent.com/wangpeng38901/argus/main/CIOMS/"
    "数据反馈结果-14910028650182082562330.xlsx"
)
DEFAULT_MAPPING_FILE_NAME = "cioms_field_mapping.json"


@dataclass
class DrugUsage:
    seq: int
    generic_name: str = ""
    trade_name: str = ""
    dosage_form: str = ""
    dose: str = ""
    dose_unit: str = ""
    frequency: str = ""
    days: str = ""
    route: str = ""
    start_date: str = ""
    end_date: str = ""
    indication: str = ""
    kind: str = "怀疑"  # 怀疑 / 合并


@dataclass
class CiomsData:
    report_code: str = ""
    report_type: str = ""
    report_source: str = ""
    organization_type: str = ""
    gender: str = ""
    age: str = ""
    age_unit: str = "岁"
    ethnicity: str = ""
    weight: str = ""
    event_date: str = ""
    reaction_name: str = ""
    reaction_description: str = ""
    outcome: str = ""
    severe_criteria: List[str] = field(default_factory=list)
    dechallenge: str = ""
    rechallenge: str = ""
    reporter_assessment: str = ""
    company_assessment: str = ""
    report_date: str = ""
    info_source: str = ""
    notes: str = ""
    primary_disease: str = ""
    important_info: str = ""
    suspected_drugs: List[DrugUsage] = field(default_factory=list)
    concomitant_drugs: List[DrugUsage] = field(default_factory=list)


def download_binary(source: str) -> bytes:
    parts = urllib.parse.urlsplit(source)

    # 支持本地路径（例如 ./CIOMS文件.pdf）
    if parts.scheme in ("", "file"):
        if parts.scheme == "file":
            local_path = Path(urllib.request.url2pathname(parts.path))
        else:
            local_path = Path(source).expanduser()

        if not local_path.is_absolute():
            local_path = (Path.cwd() / local_path).resolve()
        if not local_path.exists():
            raise FileNotFoundError(f"本地文件不存在：{local_path}")
        return local_path.read_bytes()

    # URL 下载路径
    if requests is not None:
        response = requests.get(source, timeout=30)
        response.raise_for_status()
        return response.content

    safe_path = urllib.parse.quote(parts.path, safe="/")
    encoded_url = urllib.parse.urlunsplit((parts.scheme, parts.netloc, safe_path, parts.query, parts.fragment))
    with urllib.request.urlopen(encoded_url, timeout=30) as resp:
        return resp.read()


def normalize_spaces(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    return text


def to_iso_date(cn_date: str) -> str:
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", cn_date)
    if not m:
        return ""
    year, month, day = m.groups()
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def extract_between(text: str, start: str, end: str) -> str:
    start_idx = text.find(start)
    if start_idx < 0:
        return ""
    end_idx = text.find(end, start_idx + len(start))
    if end_idx < 0:
        end_idx = len(text)
    return text[start_idx + len(start) : end_idx].strip()


def detect_electronic_pdf(pdf_bytes: bytes) -> Tuple[bool, List[str]]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    page_texts: List[str] = []
    for page in reader.pages:
        txt = (page.extract_text() or "").strip()
        page_texts.append(txt)

    non_empty_pages = sum(1 for t in page_texts if len(t) >= 50)
    total_chars = sum(len(t) for t in page_texts)
    is_electronic = total_chars >= 500 and non_empty_pages >= max(1, len(page_texts) - 1)
    return is_electronic, page_texts


def parse_checkbox_answer(text: str, question: str, candidates: List[str]) -> str:
    idx = text.find(question)
    if idx < 0:
        return ""
    snippet = text[idx : idx + 1200]
    for item in candidates:
        if f"■{item}" in snippet or f"■ {item}" in snippet:
            return item
    return ""


def parse_cioms(text: str) -> CiomsData:
    data = CiomsData()
    text = normalize_spaces(text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # 公司编号
    m = re.search(r"24b\.\s*生产企业控制编号\.?\s*([A-Z]\d+)", text)
    if not m:
        m = re.search(r"公司编号[:：]\s*([A-Z]\d+)", text)
    if m:
        data.report_code = m.group(1).strip()
        data.notes = data.report_code

    # 报告类型（首次/跟踪）
    m = re.search(r"25a\.\s*报告类型\s*([■□])首次\s*([■□])随访", text)
    if m:
        data.report_type = "首次报告" if m.group(1) == "■" else "跟踪报告"

    # 信息来源
    source_items = []
    for label in ("试验", "文献", "医疗专业", "其他"):
        if re.search(rf"■\s*{label}", text):
            source_items.append(label)
    data.info_source = "；".join(source_items)
    if "医疗专业" in source_items:
        data.organization_type = "医疗机构"

    # 年龄/性别/体重
    m = re.search(r"2a\.\s*年龄\s*(\d+(?:\.\d+)?)\s*岁", text)
    if m:
        data.age = m.group(1)
    m = re.search(r"3\.\s*性别\s*([男女])", text)
    if m:
        data.gender = m.group(1)
    m = re.search(r"3a\.\s*体重\s*(\d+(?:\.\d+)?)\s*公斤", text)
    if m:
        data.weight = m.group(1)

    # 民族/人群（可选）
    m = re.search(r"\d+岁，([^，。\s]{1,8}人)，参加", text)
    if m:
        data.ethnicity = m.group(1).strip()

    # 反应名称
    m = re.search(r"事件报告术语.*?\n([^\n]+)", text, re.S)
    if m:
        line = m.group(1).strip()
        name_match = re.search(r"（([^）]+)）", line)
        data.reaction_name = name_match.group(1) if name_match else line

    # 严重性标准
    severe_options = [
        "患者死亡",
        "导致或延长住院时间",
        "永久或重大伤残/丧失工作能力",
        "危及生命",
        "先天畸形",
        "其他",
    ]
    for label in severe_options:
        if re.search(rf"■\s*{re.escape(label)}", text):
            data.severe_criteria.append(label)

    # 事件发生日期（优先正文日期）
    m = re.search(r"(\d{4}年\d{2}月\d{2}日)，可疑感染", text)
    if m:
        data.event_date = to_iso_date(m.group(1))
    if not data.event_date:
        m = re.search(r"年\s*(\d{4})", text)
        if m:
            # 4-6 日期结构被拆分时，尽力拼接
            n = re.search(r"日\s*(\d{1,2}).*?月\s*(\d{1,2}).*?年\s*(\d{4})", text, re.S)
            if n:
                y = int(n.group(3))
                mo = int(n.group(2))
                d = int(n.group(1))
                data.event_date = f"{y:04d}-{mo:02d}-{d:02d}"

    # 病例描述（含续页）
    desc1 = extract_between(text, "病例描述:", " (续附加信息页)")
    desc2 = extract_between(text, "7+13. 不良反应描述（续）", "13. 相关实验室检查")
    data.reaction_description = " ".join(x for x in [desc1, desc2] if x).strip()

    # 转归
    m = re.search(r"结局为[:：]\s*([^。；\n]+)", text)
    if m:
        data.outcome = m.group(1).strip()

    # 停药减轻/再挑战
    data.dechallenge = parse_checkbox_answer(text, "20. 停药后反应减轻了吗？", ["是", "否", "不明", "不适用"])
    data.rechallenge = parse_checkbox_answer(text, "21. 重新用药后是否再次出现", ["是", "否", "不明", "不适用"])

    # 报告日期
    m = re.search(r"本报告日期\s*(\d{4}年\d{2}月\d{2}日)", text)
    if m:
        data.report_date = to_iso_date(m.group(1))

    # 报告人/公司评价
    m = re.search(r"研究者因果关系判断[:：]\s*([\s\S]{0,300})", text)
    if m:
        judge = m.group(1)
        data.reporter_assessment = "可能" if "可能有关" in judge else "待评估"
    m = re.search(r"企业评估[:：]\s*([\s\S]{0,300})", text)
    if m:
        judge = m.group(1)
        data.company_assessment = "可能" if "可能有关" in judge or "预期不良事件" in judge else "待评估"

    # 原患疾病：优先 17. 适应症
    indication_block = extract_between(text, "17. 适应症", "18. 给药日期（从/到）")
    if indication_block:
        m = re.search(r"#\d+\)\s*([^\n\(]+)", indication_block)
        if m:
            data.primary_disease = m.group(1).strip()

    # 重要信息：实验室检查
    data.important_info = extract_between(text, "13. 相关实验室检查", "14-19. 怀疑药物（续）")

    # 怀疑用药（主页区块，仅 #1/#2 这组结构化字段最稳定）
    main_suspect_block = extract_between(text, "14. 怀疑药物（包括通用名称）", "20. 停药后反应减轻了吗？")

    drugs: Dict[int, DrugUsage] = {}
    suspect_ids = set()
    for line in [ln.strip() for ln in main_suspect_block.splitlines() if ln.strip()]:
        m = re.match(r"#(\d+)\)\s*([^(#]+?)\(([^)]+)\)\s*(.*)$", line)
        if not m:
            continue
        idx = int(m.group(1))
        suspect_ids.add(idx)
        name = m.group(2).strip()
        trade = m.group(3).strip()
        tail = m.group(4).strip()
        if "给药方案" in tail:
            continue
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.generic_name = item.generic_name or name
        item.trade_name = item.trade_name or trade
        if "注射剂" in tail:
            item.dosage_form = "注射剂"
        elif "冻干粉" in tail:
            item.dosage_form = "冻干粉"
        drugs[idx] = item

    # 剂量（从 15. 日剂量 区块读取）
    dose_block = extract_between(text, "15. 日剂量", "16. 给药途径")
    for m in re.finditer(r"#(\d+)\)\s*([\d.]+)\s*(mg|g|ml|ug|μg|毫克|克|毫升)\s*,\s*([^\n]+)", dose_block):
        idx = int(m.group(1))
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.dose = m.group(2)
        item.dose_unit = m.group(3).replace("μg", "ug")
        freq_raw = m.group(4).strip()
        freq_count = re.search(r"(\d+)\s*次", freq_raw)
        item.frequency = freq_count.group(1) if freq_count else freq_raw
        if not item.days:
            day_m = re.search(r"(\d+)\s*次", freq_raw)
            item.days = day_m.group(1) if day_m else ""
        drugs[idx] = item

    # 给药途径
    route_block = extract_between(text, "16. 给药途径", "17. 适应症")
    for m in re.finditer(r"#(\d+)\)\s*([^\n#]+)", route_block):
        idx = int(m.group(1))
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.route = m.group(2).strip()
        drugs[idx] = item

    # 适应症
    for m in re.finditer(r"#(\d+)\)\s*([^\n#\(]+)\s*(?:\(|$)", indication_block):
        idx = int(m.group(1))
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.indication = m.group(2).strip()
        drugs[idx] = item

    # 开始结束日期
    date_block = extract_between(text, "18. 给药日期（从/到）", "19. 给药持续时间")
    for m in re.finditer(r"#(\d+)\)\s*(\d{4}年\d{2}月\d{2}日)\s*/\s*(\d{4}年\d{2}月\d{2}日|继续|不明)", date_block):
        idx = int(m.group(1))
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.start_date = to_iso_date(m.group(2))
        end_raw = m.group(3)
        item.end_date = to_iso_date(end_raw) if "年" in end_raw else end_raw
        drugs[idx] = item

    # 持续时间
    duration_block = extract_between(text, "19. 给药持续时间", "□是□否□不明")
    for m in re.finditer(r"#(\d+)\)\s*([\d.]+)\s*day", duration_block):
        idx = int(m.group(1))
        item = drugs.get(idx, DrugUsage(seq=idx))
        item.days = m.group(2)
        drugs[idx] = item

    # 只保留较小序号（主页明确列出的怀疑药）
    for idx in sorted(suspect_ids):
        item = drugs[idx]
        if item.generic_name:
            item.kind = "怀疑"
            data.suspected_drugs.append(item)

    # 合并用药
    in_concomitant = False
    for line in lines:
        if "22. 合并药物和给药日期" in line:
            in_concomitant = True
            continue
        if in_concomitant and (line.startswith("23.") or line.startswith("IV.") or line.startswith("24a.")):
            in_concomitant = False
        if not in_concomitant:
            continue
        m = re.match(r"#(\d+)\)\s*([^(#]+)\(([^)]+)\)\s*;?\s*([^/]+)\s*/\s*(.+)$", line)
        if not m:
            continue
        idx = int(m.group(1))
        item = DrugUsage(
            seq=idx,
            generic_name=m.group(2).strip(),
            trade_name=m.group(3).strip(),
            start_date=to_iso_date(m.group(4).strip()) if "年" in m.group(4) else m.group(4).strip(),
            end_date=to_iso_date(m.group(5).strip()) if "年" in m.group(5) else m.group(5).strip(),
            kind="合并",
        )
        data.concomitant_drugs.append(item)

    if not data.outcome:
        data.outcome = "未好转" if "症状持续" in text else ""

    if not data.organization_type:
        data.organization_type = "医疗机构" if "医疗专业" in data.info_source else ""

    return data


def clear_sheet_from_row(sheet, row_start: int = 2) -> None:
    if sheet.max_row >= row_start:
        sheet.delete_rows(row_start, sheet.max_row - row_start + 1)


def load_mapping_config(config_path: Path) -> Dict[str, Any]:
    suffix = config_path.suffix.lower()
    raw = config_path.read_text(encoding="utf-8")
    if suffix == ".json":
        return json.loads(raw)
    if suffix in {".yaml", ".yml"}:
        try:
            import yaml  # type: ignore
        except ImportError as exc:
            raise RuntimeError("读取 YAML 映射配置需要安装 pyyaml：pip3 install pyyaml") from exc
        cfg = yaml.safe_load(raw)
        return cfg if isinstance(cfg, dict) else {}
    raise RuntimeError(f"不支持的映射配置格式：{config_path.suffix}（仅支持 .json/.yaml/.yml）")


def cioms_to_context(cioms: CiomsData) -> Dict[str, Any]:
    all_drugs = sorted(cioms.suspected_drugs, key=lambda x: x.seq) + sorted(
        cioms.concomitant_drugs, key=lambda x: x.seq
    )
    return {
        "feedback_code": "",
        "report_type": cioms.report_type or "首次报告",
        "report_severity": "严重" if cioms.severe_criteria else "一般",
        "severe_criteria_joined": "；".join(cioms.severe_criteria),
        "organization_type": cioms.organization_type,
        "gender": cioms.gender,
        "age": cioms.age,
        "age_unit": cioms.age_unit,
        "ethnicity": cioms.ethnicity,
        "weight": cioms.weight,
        "history_adr": "不详",
        "family_adr": "不详",
        "event_date": cioms.event_date,
        "reaction_description": cioms.reaction_description,
        "outcome": cioms.outcome,
        "dechallenge": cioms.dechallenge,
        "rechallenge": cioms.rechallenge,
        "reporter_assessment": cioms.reporter_assessment,
        "company_assessment": cioms.company_assessment,
        "report_date": cioms.report_date,
        "info_source": cioms.info_source,
        "notes": cioms.notes,
        "primary_disease": cioms.primary_disease,
        "reaction_name": cioms.reaction_name,
        "reaction_level": "严重" if cioms.severe_criteria else "一般",
        "important_info": cioms.important_info,
        "all_drugs": [d.__dict__ for d in all_drugs],
    }


def resolve_value(source: str, row_ctx: Dict[str, Any], global_ctx: Dict[str, Any], default: Any = "") -> Any:
    if source.startswith("global."):
        return global_ctx.get(source[7:], default)
    if source in row_ctx:
        return row_ctx.get(source, default)
    return global_ctx.get(source, default)


def header_to_col_map(sheet, header_row: int) -> Dict[str, int]:
    return {
        str(sheet.cell(header_row, col).value).strip(): col
        for col in range(1, sheet.max_column + 1)
        if sheet.cell(header_row, col).value is not None and str(sheet.cell(header_row, col).value).strip() != ""
    }


def apply_single_sheet_mapping(sheet, sheet_cfg: Dict[str, Any], global_ctx: Dict[str, Any]) -> None:
    clear_from = int(sheet_cfg.get("clear_from_row", 0) or 0)
    if clear_from > 0:
        clear_sheet_from_row(sheet, clear_from)

    target_row = int(sheet_cfg.get("target_row", 2))
    header_row = int(sheet_cfg.get("header_row", 1))
    header_map = header_to_col_map(sheet, header_row)

    for item in sheet_cfg.get("mappings", []):
        source = item.get("source", "")
        default = item.get("default", "")
        value = resolve_value(source, {}, global_ctx, default=default)
        col = item.get("column")
        if col is None and item.get("header"):
            col = header_map.get(str(item["header"]))
        if col is None:
            continue
        sheet.cell(target_row, int(col), value=value)


def apply_list_sheet_mapping(sheet, sheet_cfg: Dict[str, Any], global_ctx: Dict[str, Any]) -> None:
    clear_from = int(sheet_cfg.get("clear_from_row", 0) or 0)
    if clear_from > 0:
        clear_sheet_from_row(sheet, clear_from)

    source_key = sheet_cfg.get("source", "")
    rows = global_ctx.get(source_key, [])
    if not isinstance(rows, list):
        return

    start_row = int(sheet_cfg.get("start_row", 2))
    header_row = int(sheet_cfg.get("header_row", 1))
    header_map = header_to_col_map(sheet, header_row)

    for i, row_data in enumerate(rows):
        if not isinstance(row_data, dict):
            continue
        row_idx = start_row + i
        for item in sheet_cfg.get("mappings", []):
            source = item.get("source", "")
            default = item.get("default", "")
            value = resolve_value(source, row_data, global_ctx, default=default)
            col = item.get("column")
            if col is None and item.get("header"):
                col = header_map.get(str(item["header"]))
            if col is None:
                continue
            sheet.cell(row_idx, int(col), value=value)


def fill_feedback_code_from_template(wb, global_ctx: Dict[str, Any]) -> None:
    if global_ctx.get("feedback_code"):
        return
    if "药品不良反应报告表" not in wb.sheetnames:
        return
    sheet = wb["药品不良反应报告表"]
    headers = header_to_col_map(sheet, 1)
    col = headers.get("反馈码")
    if not col:
        return
    value = sheet.cell(2, col).value
    if value not in (None, ""):
        global_ctx["feedback_code"] = str(value)


def write_excel(cioms: CiomsData, template_bytes: bytes, output_path: Path, mapping_cfg: Dict[str, Any]) -> None:
    wb = load_workbook(io.BytesIO(template_bytes))
    context = cioms_to_context(cioms)
    fill_feedback_code_from_template(wb, context)

    sheets_cfg = mapping_cfg.get("sheets", [])
    for sheet_cfg in sheets_cfg:
        sheet_name = sheet_cfg.get("name")
        if not sheet_name or sheet_name not in wb.sheetnames:
            continue
        sheet = wb[sheet_name]
        mode = sheet_cfg.get("mode", "single")
        if mode == "list":
            apply_list_sheet_mapping(sheet, sheet_cfg, context)
        else:
            apply_single_sheet_mapping(sheet, sheet_cfg, context)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="CIOMS PDF -> 数据反馈Excel 映射工具")
    parser.add_argument(
        "--pdf-url",
        "--pdf-source",
        default=DEFAULT_PDF_URL,
        help="CIOMS PDF 来源（支持 URL 或本地路径）",
    )
    parser.add_argument(
        "--template-url",
        "--template-source",
        default=DEFAULT_TEMPLATE_URL,
        help="Excel 模板来源（支持 URL 或本地路径）",
    )
    parser.add_argument(
        "--mapping-config",
        default=str(Path(__file__).resolve().with_name(DEFAULT_MAPPING_FILE_NAME)),
        help="字段映射配置文件路径（支持 .json/.yaml/.yml）",
    )
    parser.add_argument(
        "--output",
        default="output_数据反馈结果-14910028650182082562330.xlsx",
        help="输出 Excel 路径",
    )
    args = parser.parse_args()

    pdf_bytes = download_binary(args.pdf_url)
    ok, page_texts = detect_electronic_pdf(pdf_bytes)
    if not ok:
        raise RuntimeError("检测失败：CIOMS PDF 疑似扫描件（可提取文本不足），请提供电子档 PDF。")

    full_text = "\n".join(page_texts)
    cioms = parse_cioms(full_text)

    template_bytes = download_binary(args.template_url)
    mapping_cfg = load_mapping_config(Path(args.mapping_config).resolve())
    output = Path(args.output).resolve()
    write_excel(cioms, template_bytes, output, mapping_cfg)

    print("处理完成：")
    print(f"- 电子档校验：通过（共 {len(page_texts)} 页）")
    print(f"- 映射配置：{Path(args.mapping_config).resolve()}")
    print(f"- 输出文件：{output}")


if __name__ == "__main__":
    main()
