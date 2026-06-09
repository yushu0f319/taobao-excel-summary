#!/usr/bin/env python3
"""Update a daily Taobao scale workbook from local Taobao export files.

The script intentionally edits the XLSX zip XML directly so existing workbook
formulas, styles, images, hidden sheets, and WPS-specific content are preserved.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import zipfile
from copy import deepcopy
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def col_idx(ref: str) -> int:
    letters = ""
    for char in ref:
        if char.isalpha():
            letters += char
        else:
            break
    result = 0
    for char in letters:
        result = result * 26 + ord(char.upper()) - 64
    return result - 1


def col_letter(index: int) -> str:
    index += 1
    result = ""
    while index:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result


def norm(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", "").strip()


def read_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings: list[str] = []
    for item in root.findall(q("si")):
        strings.append("".join(text.text or "" for text in item.findall(".//" + q("t"))))
    return strings


def cell_value(cell: ET.Element | None, shared_strings: list[str] | None = None) -> str:
    if cell is None:
        return ""
    if cell.attrib.get("t") == "s":
        value = cell.find(q("v"))
        if value is not None and shared_strings is not None:
            try:
                return shared_strings[int(value.text or "0")]
            except (IndexError, ValueError):
                return value.text or ""
    inline = cell.find(q("is"))
    if inline is not None:
        return "".join(text.text or "" for text in inline.findall(".//" + q("t")))
    value = cell.find(q("v"))
    return value.text if value is not None and value.text is not None else ""


def row_cells(row: ET.Element) -> dict[int, ET.Element]:
    return {col_idx(cell.attrib.get("r", "A1")): cell for cell in row.findall(q("c"))}


def renumber_row(row: ET.Element, row_number: int) -> ET.Element:
    row.attrib["r"] = str(row_number)
    for cell in row.findall(q("c")):
        column = col_idx(cell.attrib.get("r", "A1"))
        cell.attrib["r"] = f"{col_letter(column)}{row_number}"
    return row


def clear_cell(cell: ET.Element) -> None:
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)


def set_string(cell: ET.Element, value) -> None:
    clear_cell(cell)
    if value is None or value == "":
        return
    cell.attrib["t"] = "inlineStr"
    inline = ET.SubElement(cell, q("is"))
    text = ET.SubElement(inline, q("t"))
    text.text = str(value)


def decimal_text(value) -> str | None:
    text = norm(value).replace(",", "")
    if text == "":
        return ""
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1].strip()
    try:
        decimal = Decimal(text)
    except InvalidOperation:
        return None
    if is_percent:
        decimal = decimal / Decimal(100)
    output = format(decimal, "f")
    if "." in output:
        output = output.rstrip("0").rstrip(".")
    return output or "0"


def values_match(raw_value, written_value, is_text: bool) -> bool:
    if is_text:
        return norm(raw_value) == norm(written_value)
    raw_number = decimal_text(raw_value)
    written_number = decimal_text(written_value)
    if raw_number is None or written_number is None:
        return norm(raw_value) == norm(written_value)
    return Decimal(raw_number or "0") == Decimal(written_number or "0")


def set_number(cell: ET.Element, value) -> None:
    clear_cell(cell)
    output = decimal_text(value)
    if output is None:
        set_string(cell, value)
        return
    if output == "":
        return
    number = ET.SubElement(cell, q("v"))
    number.text = output


def template_cell_is_text(cell: ET.Element | None) -> bool:
    if cell is None:
        return False
    return cell.attrib.get("t") in {"inlineStr", "str", "s"}


def normalize_xl_target(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target


def first_worksheet_path(zip_file: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    first_sheet = workbook.find(q("sheets")).find(q("sheet"))
    return normalize_xl_target(rel_targets[first_sheet.attrib[f"{{{REL_NS}}}id"]])


def read_sheet_rows(path: Path) -> list[list[str]]:
    with zipfile.ZipFile(path) as zip_file:
        shared_strings = read_shared_strings(zip_file)
        sheet_path = first_worksheet_path(zip_file)
        root = ET.fromstring(zip_file.read(sheet_path))
    rows: list[list[str]] = []
    for row in root.findall(".//" + q("sheetData") + "/" + q("row")):
        values: dict[int, str] = {}
        max_column = -1
        for cell in row.findall(q("c")):
            index = col_idx(cell.attrib.get("r", "A1"))
            max_column = max(max_column, index)
            values[index] = cell_value(cell, shared_strings)
        rows.append([values.get(index, "") for index in range(max_column + 1)])
    return rows


def classify(headers: Iterable[str]) -> str | None:
    header_set = {norm(header) for header in headers}
    if {"总成交额（元）", "预计收入（元）", "有效订单", "顾客实付金额（元）"}.issubset(header_set):
        return "TB门店"
    if {"总曝光人数", "总进店人数", "总下单人数", "进店转化率", "下单转化率"}.issubset(header_set):
        return "TB流量"
    if {"履约渗透率", "轨迹渗透率", "及时送达率", "流失订单", "拣货及时订单率"}.issubset(header_set):
        return "TB履约"
    return None


@dataclass
class RawTable:
    kind: str
    path: Path
    headers: list[str]
    rows: list[list[str]]

    @property
    def dates(self) -> set[str]:
        return {norm(row[0]) for row in self.rows if row}


def load_raw_tables(raw_paths: Iterable[Path]) -> dict[str, RawTable]:
    tables: dict[str, RawTable] = {}
    for path in raw_paths:
        rows = read_sheet_rows(path)
        if not rows:
            raise ValueError(f"原始表没有可读取内容: {path}")
        headers = [norm(header) for header in rows[0]]
        kind = classify(headers)
        if kind is None:
            raise ValueError(f"无法识别原始表类型: {path.name}")
        if kind in tables:
            raise ValueError(f"重复提供 {kind} 原始表: {tables[kind].path.name}, {path.name}")
        tables[kind] = RawTable(kind=kind, path=path, headers=headers, rows=rows[1:])
    missing = {"TB门店", "TB流量", "TB履约"} - set(tables)
    if missing:
        raise ValueError(f"缺少原始表: {', '.join(sorted(missing))}")
    dates = set().union(*(table.dates for table in tables.values()))
    if len(dates) != 1:
        raise ValueError(f"三张原始表日期不一致: {', '.join(sorted(dates))}")
    for table in tables.values():
        if not table.rows:
            raise ValueError(f"{table.kind} 没有数据行: {table.path.name}")
    return tables


def workbook_sheet_paths(zip_file: zipfile.ZipFile) -> tuple[ET.Element, dict[str, str]]:
    workbook = ET.fromstring(zip_file.read("xl/workbook.xml"))
    rels = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    sheet_paths: dict[str, str] = {}
    for sheet in workbook.find(q("sheets")).findall(q("sheet")):
        sheet_paths[sheet.attrib["name"]] = normalize_xl_target(rel_targets[sheet.attrib[f"{{{REL_NS}}}id"]])
    return workbook, sheet_paths


def append_table_to_sheet(
    sheet_xml: bytes,
    shared_strings: list[str],
    raw_table: RawTable,
) -> tuple[bytes, dict]:
    root = ET.fromstring(sheet_xml)
    sheet_data = root.find(q("sheetData"))
    existing_rows = sheet_data.findall(q("row"))
    row_by_number = {int(row.attrib["r"]): row for row in existing_rows}
    max_row = max(row_by_number)

    header_rows = []
    for row_number, row in row_by_number.items():
        first_cell = row_cells(row).get(0)
        if norm(cell_value(first_cell, shared_strings)) == "日期":
            header_rows.append(row_number)
    if not header_rows:
        raise ValueError(f"{raw_table.kind} 找不到可复制的表头行")

    template_header_number = max(header_rows)
    template_blank_number = template_header_number - 1
    template_data_start = template_header_number + 1

    template_header_cells = row_cells(row_by_number[template_header_number])
    template_headers = [
        norm(cell_value(template_header_cells.get(column), shared_strings))
        for column in range(len(raw_table.headers))
    ]
    if template_headers != raw_table.headers:
        raise ValueError(
            f"{raw_table.kind} 表头与模板不匹配: "
            f"模板={template_headers}, 原表={raw_table.headers}"
        )

    append_blank_number = max_row + 1
    append_header_number = max_row + 2
    append_data_start = max_row + 3

    blank_template = row_by_number.get(template_blank_number)
    if blank_template is None:
        blank_row = ET.Element(q("row"), {"r": str(append_blank_number)})
    else:
        blank_row = renumber_row(deepcopy(blank_template), append_blank_number)
        for cell in blank_row.findall(q("c")):
            clear_cell(cell)
    sheet_data.append(blank_row)

    header_template = row_by_number[template_header_number]
    header_row = renumber_row(deepcopy(header_template), append_header_number)
    header_cells = row_cells(header_row)
    max_col = max(header_cells.keys()) if header_cells else len(raw_table.headers) - 1
    for column in range(max_col + 1):
        cell = header_cells.get(column)
        if cell is None:
            cell = ET.SubElement(header_row, q("c"), {"r": f"{col_letter(column)}{append_header_number}"})
        set_string(cell, raw_table.headers[column] if column < len(raw_table.headers) else "")
    sheet_data.append(header_row)

    value_mismatches = []
    for offset, raw_row in enumerate(raw_table.rows):
        template_number = template_data_start + min(offset, max_row - template_data_start)
        template_row = row_by_number[template_number]
        new_row_number = append_data_start + offset
        new_row = renumber_row(deepcopy(template_row), new_row_number)
        new_cells = row_cells(new_row)
        template_cells = row_cells(template_row)
        max_col = max(new_cells.keys()) if new_cells else len(raw_table.headers) - 1
        for column in range(max_col + 1):
            cell = new_cells.get(column)
            if cell is None:
                cell = ET.SubElement(new_row, q("c"), {"r": f"{col_letter(column)}{new_row_number}"})
            if column >= len(raw_table.headers):
                clear_cell(cell)
                continue
            raw_value = raw_row[column] if column < len(raw_row) else ""
            template_cell = template_cells.get(column)
            is_text = template_cell_is_text(template_cell)
            if is_text:
                set_string(cell, raw_value)
            else:
                set_number(cell, raw_value)
            if not values_match(raw_value, cell_value(cell), is_text):
                value_mismatches.append(
                    {
                        "row": new_row_number,
                        "column": col_letter(column),
                        "header": raw_table.headers[column],
                        "source": norm(raw_value),
                        "written": norm(cell_value(cell)),
                    }
                )
        sheet_data.append(new_row)

    if value_mismatches:
        raise ValueError(f"{raw_table.kind} 写入后校验失败: {value_mismatches[:5]}")

    dimension = root.find(q("dimension"))
    if dimension is not None:
        dimension.attrib["ref"] = f"A1:{col_letter(max_col)}{append_data_start + len(raw_table.rows) - 1}"

    report = {
        "kind": raw_table.kind,
        "source": str(raw_table.path),
        "rows": len(raw_table.rows),
        "dates": sorted(raw_table.dates),
        "blank_row": append_blank_number,
        "header_row": append_header_number,
        "data_start_row": append_data_start,
        "data_end_row": append_data_start + len(raw_table.rows) - 1,
        "value_mismatches": 0,
    }
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), report


def enable_recalculation(workbook_xml: ET.Element) -> bytes:
    calc_pr = workbook_xml.find(q("calcPr"))
    if calc_pr is None:
        calc_pr = ET.SubElement(workbook_xml, q("calcPr"))
    calc_pr.attrib["calcMode"] = "auto"
    calc_pr.attrib["fullCalcOnLoad"] = "1"
    calc_pr.attrib["forceFullCalc"] = "1"
    return ET.tostring(workbook_xml, encoding="utf-8", xml_declaration=True)


def update_workbook(base_path: Path, raw_paths: Iterable[Path], output_path: Path) -> dict:
    base_path = Path(base_path)
    raw_paths = [Path(path) for path in raw_paths]
    output_path = Path(output_path)
    raw_tables = load_raw_tables(raw_paths)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(base_path, "r") as input_zip:
        shared_strings = read_shared_strings(input_zip)
        workbook_xml, sheet_paths = workbook_sheet_paths(input_zip)
        modified: dict[str, bytes] = {}
        appended: list[dict] = []

        for kind in ["TB门店", "TB流量", "TB履约"]:
            if kind not in sheet_paths:
                raise ValueError(f"底稿缺少目标 sheet: {kind}")
            sheet_path = sheet_paths[kind]
            sheet_xml, report = append_table_to_sheet(
                input_zip.read(sheet_path),
                shared_strings,
                raw_tables[kind],
            )
            modified[sheet_path] = sheet_xml
            appended.append(report)

        modified["xl/workbook.xml"] = enable_recalculation(workbook_xml)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            for item in input_zip.infolist():
                output_zip.writestr(item, modified.get(item.filename, input_zip.read(item.filename)))

    return {
        "output": str(output_path),
        "date": sorted(set().union(*(table.dates for table in raw_tables.values())))[0],
        "appended": appended,
    }


def expand_inputs(inputs: Iterable[str]) -> list[Path]:
    paths: list[Path] = []
    for input_item in inputs:
        path = Path(input_item)
        if path.is_dir():
            paths.extend(sorted(path.glob("*.xlsx")))
        else:
            paths.append(path)
    return paths


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="更新淘宝日度规模数据 Excel。")
    parser.add_argument("--base", required=True, help="上一日成品表路径，例如 6月4日度-规模数据.xlsx")
    parser.add_argument("--input", action="append", required=True, help="原始表文件或目录，可重复传入")
    parser.add_argument("--output", required=True, help="输出 Excel 路径")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出处理报告")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        raw_paths = expand_inputs(args.input)
        report = update_workbook(Path(args.base), raw_paths, Path(args.output))
    except Exception as exc:
        print(f"处理失败: {exc}", file=sys.stderr)
        return 1

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"输出文件: {report['output']}")
        print(f"处理日期: {report['date']}")
        for item in report["appended"]:
            print(
                f"{item['kind']}: {Path(item['source']).name} -> "
                f"{item['data_start_row']}-{item['data_end_row']} 行, 共 {item['rows']} 行"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
