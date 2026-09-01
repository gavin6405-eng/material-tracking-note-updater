from __future__ import annotations

import io
import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import streamlit as st
except ModuleNotFoundError:  # 允許在未安裝 Streamlit 的環境執行核心比對測試
    st = None


APP_TITLE = "倉庫物管發料 × 物料追蹤彙整"
WEEKLY_SHEET_PATTERN = re.compile(r"^\d{4}~\d{4}$")


def normalize_header(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper()
    return re.sub(r"[\s\n\r\t_（）()／/\-]+", "", text)


def normalize_key(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", "", str(value).strip()).upper()


ALIASES = {
    "WO": ["製令", "專案代號", "製令單"],
    "CATEGORY": ["CATEGORY", "類別"],
    "FRAME": ["骨架/骨包", "FRAME/ FRAME SET", "FRAME SET", "FRAME"],
    "PU": ["PU"],
    "FACILITY": ["FACILITY"],
    "OTHER": ["其他託外模組", "其他託外", "託(其他)"],
    "RB": ["RB"],
    "LP": ["LP"],
    "AL": ["AL"],
    "FFU": ["FFU"],
    "X_TABLE": ["X-TABLE", "XTABLE", "X軸"],
    "MACHINED": ["加工件", "加工件(交期)", "加工件缺料(不含模組.市購件)"],
    "MARKET": ["市購件", "RZ市構件"],
    "CUSTOMER": ["客供", "客供料"],
    "SIGNAL": ["SIGNAL"],
    "INTERLOCK": ["INTERLOCK"],
}
NORMALIZED_ALIASES = {
    key: {normalize_header(alias) for alias in values} for key, values in ALIASES.items()
}


def clean_display_value(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).strftime("%Y/%m/%d")
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    text = re.sub(r"(\d{4})-(\d{2})-(\d{2}) 00:00:00", r"\1/\2/\3", text)
    return text


def is_meaningful(value: Any) -> bool:
    text = clean_display_value(value)
    return bool(text and text not in {"0", "0.0", "-"})


def find_header_row(df: pd.DataFrame) -> int:
    wo_aliases = NORMALIZED_ALIASES["WO"]
    for row_idx in range(min(15, len(df))):
        values = {normalize_header(v) for v in df.iloc[row_idx].tolist()}
        if values & wo_aliases:
            return row_idx
    raise ValueError("找不到『製令／專案代號』標題列。")


def map_columns(df: pd.DataFrame, header_row: int, include_second_header: bool) -> dict[str, int]:
    mapped: dict[str, int] = {}
    rows = [header_row]
    if include_second_header and header_row + 1 < len(df):
        rows.append(header_row + 1)
    for col_idx in range(df.shape[1]):
        labels = {normalize_header(df.iat[r, col_idx]) for r in rows}
        labels.discard("")
        for standard, aliases in NORMALIZED_ALIASES.items():
            if standard in mapped:
                continue
            if labels & aliases:
                mapped[standard] = col_idx
    return mapped


def source_header_depth(df: pd.DataFrame, header_row: int) -> int:
    if header_row + 1 >= len(df):
        return 1
    second = {normalize_header(v) for v in df.iloc[header_row + 1].tolist()}
    status_aliases = set().union(
        *(NORMALIZED_ALIASES[k] for k in ["FRAME", "PU", "FACILITY", "OTHER", "RB", "LP", "AL", "FFU", "X_TABLE", "MACHINED"])
    )
    return 2 if len(second & status_aliases) >= 2 else 1


def list_excel_sheets(file_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names


def choose_target_sheet(file_bytes: bytes) -> tuple[list[str], str]:
    sheets = list_excel_sheets(file_bytes)
    best_sheet = sheets[0]
    best_score = -1
    for sheet in sheets:
        try:
            preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, nrows=15, dtype=object)
            header_row = find_header_row(preview)
            mapping = map_columns(preview, header_row, include_second_header=False)
            score = len(mapping)
            if "WO" in mapping and score > best_score:
                best_score = score
                best_sheet = sheet
        except Exception:
            continue
    return sheets, best_sheet


def read_target(file_name: str, file_bytes: bytes, sheet_name: str | None = None) -> tuple[pd.DataFrame, int, str]:
    suffix = Path(file_name).suffix.lower()
    if suffix == ".csv":
        errors = []
        for encoding in ["utf-8-sig", "utf-8", "cp950", "big5"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=encoding, dtype=object)
                return df, 0, "CSV"
            except Exception as exc:
                errors.append(str(exc))
        raise ValueError("CSV 編碼無法辨識，請另存為 UTF-8 CSV 後再試。")
    if suffix not in {".xlsx", ".xlsm"}:
        raise ValueError("物料追蹤彙整只支援 CSV、XLSX 或 XLSM。")
    if not sheet_name:
        _, sheet_name = choose_target_sheet(file_bytes)
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, dtype=object)
    header_row = find_header_row(raw)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row, dtype=object)
    df.columns = [clean_display_value(c) or f"未命名欄位_{i + 1}" for i, c in enumerate(df.columns)]
    return df, header_row, sheet_name


def target_column_mapping(df: pd.DataFrame) -> dict[str, str]:
    mapped: dict[str, str] = {}
    for col in df.columns:
        label = normalize_header(col)
        for standard, aliases in NORMALIZED_ALIASES.items():
            if standard not in mapped and label in aliases:
                mapped[standard] = col
    return mapped


def build_source_records(
    file_bytes: bytes, selected_sheets: list[str]
) -> tuple[dict[tuple[str, str], dict[str, Any]], dict[str, list[dict[str, Any]]], pd.DataFrame]:
    workbook_sheets = list_excel_sheets(file_bytes)
    selected = [s for s in workbook_sheets if s in selected_sheets]
    if not selected:
        raise ValueError("沒有選到可讀取的倉庫週別工作表。")

    records: dict[tuple[str, str], dict[str, Any]] = {}
    source_log: list[dict[str, Any]] = []
    for sheet_order, sheet in enumerate(selected):
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, dtype=object)
        if raw.empty:
            continue
        try:
            header_row = find_header_row(raw)
        except ValueError:
            continue
        depth = source_header_depth(raw, header_row)
        columns = map_columns(raw, header_row, include_second_header=(depth == 2))
        if "WO" not in columns:
            continue
        for row_idx in range(header_row + depth, len(raw)):
            wo = normalize_key(raw.iat[row_idx, columns["WO"]])
            if not wo:
                continue
            category = normalize_key(raw.iat[row_idx, columns["CATEGORY"]]) if "CATEGORY" in columns else ""
            record: dict[str, Any] = {
                "WO": wo,
                "CATEGORY": category,
                "來源工作表": sheet,
                "來源列": row_idx + 1,
                "工作表順序": sheet_order,
            }
            for status_key in ["FRAME", "PU", "FACILITY", "OTHER", "RB", "LP", "AL", "FFU", "X_TABLE", "MACHINED", "MARKET", "CUSTOMER"]:
                record[status_key] = clean_display_value(raw.iat[row_idx, columns[status_key]]) if status_key in columns else ""
            records[(wo, category)] = record
            source_log.append(record.copy())

    by_wo: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for record in records.values():
        by_wo[record["WO"]].append(record)
    for wo, candidates in list(by_wo.items()):
        newest = max(r["工作表順序"] for r in candidates)
        by_wo[wo] = [r for r in candidates if r["工作表順序"] == newest]
    return records, dict(by_wo), pd.DataFrame(source_log)


def extract_module_status(other_status: str, keyword: str) -> str:
    if not other_status:
        return ""
    parts = re.split(r"[\n\r]+", other_status)
    matches = [part.strip() for part in parts if keyword.upper() in part.upper()]
    return "\n".join(matches)


def source_status_for_target(record: dict[str, Any], target_standard: str) -> str:
    if target_standard == "SIGNAL":
        return extract_module_status(record.get("OTHER", ""), "SIGNAL")
    if target_standard == "INTERLOCK":
        return extract_module_status(record.get("OTHER", ""), "INTERLOCK")
    return clean_display_value(record.get(target_standard, ""))


def dated_status(status: str, update_date: date) -> str:
    return f"【{update_date.strftime('%Y/%m/%d')} 倉庫更新】{status}"


def compare_and_update(
    target_df: pd.DataFrame,
    exact_records: dict[tuple[str, str], dict[str, Any]],
    records_by_wo: dict[str, list[dict[str, Any]]],
    update_date: date,
    write_mode: str = "append",
    only_existing_shortage: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, dict[str, int]]:
    result = target_df.copy()
    columns = target_column_mapping(result)
    if "WO" not in columns:
        raise ValueError("物料追蹤彙整找不到『專案代號／製令』欄位。")
    if "CATEGORY" not in columns:
        raise ValueError("物料追蹤彙整找不到『Category』欄位。")

    status_fields = ["FRAME", "PU", "FACILITY", "OTHER", "RB", "LP", "AL", "FFU", "X_TABLE", "MACHINED", "MARKET", "CUSTOMER", "SIGNAL", "INTERLOCK"]
    logs: list[dict[str, Any]] = []
    unmatched: list[dict[str, Any]] = []
    matched_rows: set[int] = set()
    fallback_rows = 0

    for idx, row in result.iterrows():
        wo = normalize_key(row.get(columns["WO"]))
        category = normalize_key(row.get(columns["CATEGORY"]))
        if not wo:
            continue
        record = exact_records.get((wo, category))
        match_type = "製令＋Category"
        if record is None:
            candidates = records_by_wo.get(wo, [])
            if len(candidates) == 1:
                record = candidates[0]
                match_type = "製令（唯一資料）"
                fallback_rows += 1

        shortage_fields = [
            field for field in status_fields if field in columns and is_meaningful(row.get(columns[field]))
        ]
        if record is None:
            if shortage_fields:
                unmatched.append({
                    "目標列": int(idx) + 2,
                    "製令": clean_display_value(row.get(columns["WO"])),
                    "Category": clean_display_value(row.get(columns["CATEGORY"])),
                    "原缺料欄位": "、".join(columns[field] for field in shortage_fields),
                    "原因": "倉庫週別工作表查無相同製令＋Category",
                })
            continue

        matched_rows.add(int(idx))
        for standard in status_fields:
            if standard not in columns:
                continue
            col_name = columns[standard]
            original = clean_display_value(row.get(col_name))
            if only_existing_shortage and not is_meaningful(original):
                continue
            warehouse_status = source_status_for_target(record, standard)
            if not is_meaningful(warehouse_status):
                continue
            update_text = dated_status(warehouse_status, update_date)
            if update_text in original:
                continue
            if write_mode == "overwrite":
                updated = update_text
            elif original:
                updated = f"{original}\n{update_text}"
            else:
                updated = update_text
            result.at[idx, col_name] = updated
            logs.append({
                "目標資料列": int(idx) + 2,
                "製令": clean_display_value(row.get(columns["WO"])),
                "Category": clean_display_value(row.get(columns["CATEGORY"])),
                "更新欄位": col_name,
                "原缺料": original,
                "倉庫最新狀態": warehouse_status,
                "更新後": updated,
                "更新日期": update_date.strftime("%Y/%m/%d"),
                "比對方式": match_type,
                "來源工作表": record["來源工作表"],
                "來源列": record["來源列"],
                "目標索引": int(idx),
            })

    stats = {
        "目標總列數": len(result),
        "成功比對列數": len(matched_rows),
        "更新儲存格數": len(logs),
        "未比對列數": len(unmatched),
        "製令唯一值備援比對列數": fallback_rows,
    }
    return result, pd.DataFrame(logs), pd.DataFrame(unmatched), stats


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
UPDATED_FILL = PatternFill("solid", fgColor="FFF2CC")


def style_data_sheet(ws, data_row_height: int = 45) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx in range(1, ws.max_column + 1):
        values = [clean_display_value(ws.cell(row=r, column=col_idx).value) for r in range(1, min(ws.max_row, 80) + 1)]
        width = min(max(max((len(v) for v in values), default=8) + 2, 10), 34)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    for row in ws.iter_rows(min_row=2):
        ws.row_dimensions[row[0].row].height = data_row_height
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def add_dataframe_sheet(wb, title: str, df: pd.DataFrame, drop_internal: bool = False) -> None:
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    data = df.copy()
    if drop_internal:
        data = data.drop(columns=[c for c in ["目標索引"] if c in data.columns], errors="ignore")
    if data.empty and len(data.columns) == 0:
        data = pd.DataFrame({"結果": ["無資料"]})
    for col_idx, col in enumerate(data.columns, start=1):
        ws.cell(row=1, column=col_idx, value=str(col))
    for row_idx, values in enumerate(data.itertuples(index=False, name=None), start=2):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row=row_idx, column=col_idx, value=clean_display_value(value))
    style_data_sheet(ws)


def build_output_workbook(
    target_name: str,
    target_bytes: bytes,
    target_sheet: str,
    target_header_row: int,
    updated_df: pd.DataFrame,
    update_log: pd.DataFrame,
    unmatched_df: pd.DataFrame,
) -> bytes:
    suffix = Path(target_name).suffix.lower()
    if suffix == ".csv":
        wb = Workbook()
        ws = wb.active
        ws.title = "物料追蹤彙整_更新後"
        for col_idx, col in enumerate(updated_df.columns, start=1):
            ws.cell(row=1, column=col_idx, value=str(col))
        for row_idx, values in enumerate(updated_df.itertuples(index=False, name=None), start=2):
            for col_idx, value in enumerate(values, start=1):
                ws.cell(row=row_idx, column=col_idx, value=clean_display_value(value))
        style_data_sheet(ws)
        for item in update_log.to_dict("records"):
            target_idx = int(item["目標索引"]) + 2
            target_col = list(updated_df.columns).index(item["更新欄位"]) + 1
            ws.cell(row=target_idx, column=target_col).fill = UPDATED_FILL
    else:
        keep_vba = suffix == ".xlsm"
        wb = load_workbook(io.BytesIO(target_bytes), keep_vba=keep_vba)
        ws = wb[target_sheet]
        header_excel_row = target_header_row + 1
        header_map = {
            clean_display_value(ws.cell(row=header_excel_row, column=c).value): c
            for c in range(1, ws.max_column + 1)
        }
        for item in update_log.to_dict("records"):
            col_name = item["更新欄位"]
            if col_name not in header_map:
                continue
            excel_row = header_excel_row + 1 + int(item["目標索引"])
            cell = ws.cell(row=excel_row, column=header_map[col_name])
            cell.value = item["更新後"]
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.fill = UPDATED_FILL

    add_dataframe_sheet(wb, "自動比對紀錄", update_log, drop_internal=True)
    add_dataframe_sheet(wb, "未比對清單", unmatched_df)
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def main() -> None:
    if st is None:
        raise RuntimeError("尚未安裝 Streamlit，請先執行 pip install -r requirements.txt。")
    st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")
    st.title("📦 倉庫物管發料 × 物料追蹤彙整")
    st.caption("依『製令＋Category』比對，保留原缺料內容並追加日期與倉庫最新狀態。")

    col1, col2 = st.columns(2)
    with col1:
        target_file = st.file_uploader("① 上傳物料追蹤彙整（CSV／XLSX）", type=["csv", "xlsx", "xlsm"])
    with col2:
        warehouse_file = st.file_uploader("② 上傳倉庫物管發料（XLSX）", type=["xlsx"])

    if not target_file or not warehouse_file:
        st.info("請先上傳兩份檔案。程式不會改動原始檔，完成後會產生新的 Excel。")
        return

    target_bytes = target_file.getvalue()
    warehouse_bytes = warehouse_file.getvalue()
    target_sheet = "CSV"
    if Path(target_file.name).suffix.lower() != ".csv":
        target_sheets, suggested = choose_target_sheet(target_bytes)
        target_sheet = st.selectbox("物料追蹤彙整工作表", target_sheets, index=target_sheets.index(suggested))

    all_source_sheets = list_excel_sheets(warehouse_bytes)
    weekly_sheets = [s for s in all_source_sheets if WEEKLY_SHEET_PATTERN.fullmatch(s)]
    scope = st.radio("倉庫資料範圍", ["自動抓全部週別（同製令採較後工作表）", "只抓最新週別", "手動選擇週別"], horizontal=True)
    if scope.startswith("自動"):
        selected_source_sheets = weekly_sheets
    elif scope.startswith("只抓"):
        selected_source_sheets = weekly_sheets[-1:] if weekly_sheets else []
    else:
        selected_source_sheets = st.multiselect("選擇週別工作表", weekly_sheets, default=weekly_sheets[-2:])

    option1, option2, option3 = st.columns(3)
    with option1:
        update_date = st.date_input("備註更新日期", value=date.today())
    with option2:
        write_label = st.selectbox("寫入方式", ["保留原缺料並追加", "直接覆蓋原缺料"])
    with option3:
        only_existing = st.checkbox("只更新原本有缺料內容的欄位", value=True)

    if st.button("開始比對並產生更新檔", type="primary", use_container_width=True):
        try:
            with st.spinner("正在讀取、比對並製作 Excel…"):
                target_df, target_header_row, actual_target_sheet = read_target(
                    target_file.name,
                    target_bytes,
                    None if target_sheet == "CSV" else target_sheet,
                )
                exact_records, records_by_wo, _ = build_source_records(warehouse_bytes, selected_source_sheets)
                updated_df, update_log, unmatched_df, stats = compare_and_update(
                    target_df,
                    exact_records,
                    records_by_wo,
                    update_date,
                    write_mode="overwrite" if write_label.startswith("直接") else "append",
                    only_existing_shortage=only_existing,
                )
                output_bytes = build_output_workbook(
                    target_file.name,
                    target_bytes,
                    actual_target_sheet,
                    target_header_row,
                    updated_df,
                    update_log,
                    unmatched_df,
                )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("目標資料", f"{stats['目標總列數']} 列")
            m2.metric("成功比對", f"{stats['成功比對列數']} 列")
            m3.metric("已更新", f"{stats['更新儲存格數']} 格")
            m4.metric("未比對", f"{stats['未比對列數']} 列")

            if not update_log.empty:
                st.subheader("更新預覽")
                st.dataframe(
                    update_log.drop(columns=["目標索引"], errors="ignore").head(200),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.warning("沒有可寫入的更新。請確認週別範圍、製令、Category，以及目標欄位原本是否有缺料內容。")

            output_name = f"物料追蹤彙整_已更新_{update_date.strftime('%Y%m%d')}.xlsx"
            st.download_button(
                "下載更新後 Excel",
                data=output_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
            st.caption("更新儲存格以淡黃色標示；另附『自動比對紀錄』與『未比對清單』工作表。")
        except Exception as exc:
            st.error(f"處理失敗：{exc}")


if __name__ == "__main__":
    main()
