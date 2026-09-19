"""
Parser for Metro M.D.F. daily sanding defect report (.xls, one sheet per day).

File layout per day-sheet (0-indexed rows/cols):
  Row 10 (col A) == "TH"  -> marks a sheet that HAS a data table (some days have
                              no production, e.g. "ไม่ได้ขัดไม้" -> skip entirely)
  Row 10-12               -> 3-row merged header
  Row 13 onward            -> one row per (wood spec / เครื่องขัด) production line
      col A (0) = TH (thickness)
      col E (4) = หัวหน้ากะ (shift leader name)
      col G (6) = วันผลิต (production date, dd/mm) -- NOT reliable for year/month
      col H (7) = จำนวนที่เข้าขัด (แผ่น)
      col I (8) = ยอดไม้เข้าขัดทั้งหมด (CU.)   <-- denominator for %
      col N (13) = ตำหนิ(คิว) ทั้งหมด            <-- numerator for %
      col Q (17) onward = per-defect-type columns (Pcs), header spans row 11+12
  Row where col G == "Total" -> end of data table

The sheet NAME itself (e.g. "1".."31") is the day-of-month; it is combined with
a user-supplied month/year to build the real date, since the in-sheet date text
is inconsistently placed/formatted across days.
"""

import xlrd
import pandas as pd
import re

TABLE_START_ROW = 13      # 0-indexed; first data row
HEADER_ROWS = (10, 11, 12)
COL_THICKNESS = 0
COL_LEADER = 4
COL_QTY_SHEETS = 7
COL_CU_IN = 8              # ยอดไม้เข้าขัดทั้งหมด (CU.)
COL_DEFECT_TOTAL = 13      # ตำหนิ(คิว) ทั้งหมด
COL_DEFECT_TYPES_START = 17  # right of this = per-defect-type Pcs columns


def _sheet_has_data_table(sheet) -> bool:
    if sheet.nrows <= 10:
        return False
    try:
        return str(sheet.cell_value(10, COL_THICKNESS)).strip() == "TH"
    except IndexError:
        return False


def _parse_defect_cell(v):
    """
    Parse a defect-type cell value into (qty_for_column_total, list_of_(technician, qty)).

    Handles:
      - plain number -> (number, [])
      - "name/qty" or "qty/name" e.g. "วัฒน์/2", "2/วัต" -> (qty, [(name, qty)])
      - multiple comma-separated entries e.g. "วัฒน์/7,เจี๊ยบ/10"
        -> (17, [("วัฒน์",7), ("เจี๊ยบ",10)])
      - "OK" / blank / anything unparseable -> (0, [])
    """
    if isinstance(v, (int, float)):
        return float(v), []
    if not isinstance(v, str):
        return 0.0, []
    s = v.strip()
    if s == "" or s.upper() == "OK":
        return 0.0, []

    total = 0.0
    techs = []
    for part in s.split(","):
        part = part.strip()
        if not part:
            continue
        if "/" not in part:
            continue  # unparseable fragment, ignore
        a, b = [x.strip() for x in part.split("/", 1)]

        def _is_num(x):
            try:
                float(x)
                return True
            except ValueError:
                return False

        if _is_num(a) and not _is_num(b):
            qty, name = float(a), b
        elif _is_num(b) and not _is_num(a):
            name, qty = a, float(b)
        else:
            continue  # can't determine which side is the number

        total += qty
        techs.append((name, qty))

    return total, techs


def _build_defect_col_names(sheet) -> dict:
    """Combine header rows 11 and 12 into a readable defect-type name per column,
    and forward-fill the row-10 category label (blank cells belong to the last
    non-blank category to their left) so each column also gets a `category`.

    Returns: {col_index: {"label": str, "category": str}}
    """
    names = {}
    current_category = ""
    for c in range(COL_DEFECT_TYPES_START, sheet.ncols):
        cat_cell = str(sheet.cell_value(10, c)).strip()
        if cat_cell:
            current_category = cat_cell
        part1 = str(sheet.cell_value(11, c)).strip()
        part2 = str(sheet.cell_value(12, c)).strip()
        label = (part1 + part2).strip()
        if label:
            names[c] = {"label": label, "category": current_category}
    return names


def _find_total_row(sheet) -> int:
    for r in range(TABLE_START_ROW, sheet.nrows):
        v = sheet.cell_value(r, 6)
        if str(v).strip() == "Total":
            return r
    return sheet.nrows  # no Total marker found; read until end


def _is_no_sanding_day(sheet) -> bool:
    """True if row 1 explicitly states no sanding happened that day (e.g. 'ไม่ได้ขัดไม้')."""
    if sheet.nrows <= 1:
        return False
    row1 = str(sheet.cell_value(1, 0))
    return "ไม่ได้ขัดไม้" in row1


def parse_workbook(file_bytes, year_be: int, month: int):
    """
    Parse the whole monthly .xls workbook.

    Returns (df, tech_df, recorded_days, defect_label_category):
      df      : tidy row-per-production-line DataFrame, one row per production
                line, for days with real production/defect data only (CU>0 or
                defect>0 somewhere that day)
      tech_df : tidy row-per-technician-defect-mention DataFrame with columns
                [date, day, defect_type, technician, qty] -- built from cells
                like "วัฒน์/2" or "เจี๊ยบ/7,พงษ์/4" found in the defect-type columns
      recorded_days : set of day-of-month ints that count as "มีชุดข้อมูล" --
                days with real production data, PLUS days explicitly marked
                "ไม่ได้ขัดไม้" (no sanding that day, but still a real logged
                day). Days whose sheet exists only as an empty/zero placeholder
                template (all rows CU=0 and defect=0) are excluded.
      defect_label_category : dict mapping each defect-type label (matching the
                "ตำหนิ::<label>" column suffixes in df) to its big-category
                label from the file's row-10 header, e.g.
                {"แตกเครื่องขัด": "ตำหนิเครื่องขัด", "คราบฝุ่นบน": "ตำหนิจาก Press", ...}

    year_be: Buddhist-era year, e.g. 2569
    month: 1-12
    """
    wb = xlrd.open_workbook(file_contents=file_bytes)
    all_rows = []
    tech_rows = []
    recorded_days = set()
    defect_label_category = {}  # e.g. {"แตกเครื่องขัด": "ตำหนิเครื่องขัด", ...}

    for sheet_name in wb.sheet_names():
        if not re.fullmatch(r"\d+", sheet_name.strip()):
            continue  # skip any non-numeric sheet name defensively
        day = int(sheet_name.strip())
        sheet = wb.sheet_by_name(sheet_name)

        if not _sheet_has_data_table(sheet):
            # No production table -- still counts as a recorded day if it
            # explicitly says sanding didn't happen that day.
            if _is_no_sanding_day(sheet):
                recorded_days.add(day)
            continue

        defect_cols = _build_defect_col_names(sheet)
        for info in defect_cols.values():
            defect_label_category[info["label"]] = info["category"]
        total_row = _find_total_row(sheet)

        try:
            date = pd.Timestamp(year=year_be - 543, month=month, day=day)
        except ValueError:
            continue  # e.g. day 31 in a 30-day month -> skip invalid sheet

        day_rows = []
        day_tech_rows = []
        day_cu_total = 0.0
        day_defect_total = 0.0

        for r in range(TABLE_START_ROW, total_row):
            leader = str(sheet.cell_value(r, COL_LEADER)).strip()
            cu_in = sheet.cell_value(r, COL_CU_IN)
            defect_total = sheet.cell_value(r, COL_DEFECT_TOTAL)
            qty_pcs_in = sheet.cell_value(r, COL_QTY_SHEETS)  # จำนวนที่เข้าขัด (แผ่น)

            # Skip fully blank / non-data rows (e.g. stray formatting rows)
            if leader == "" and (cu_in in ("", None)):
                continue
            if not isinstance(cu_in, (int, float)):
                continue

            cu_val = float(cu_in) if cu_in not in ("", None) else 0.0
            defect_val = float(defect_total) if isinstance(defect_total, (int, float)) else 0.0
            qty_pcs_val = float(qty_pcs_in) if isinstance(qty_pcs_in, (int, float)) else 0.0
            day_cu_total += cu_val
            day_defect_total += defect_val

            row = {
                "date": date,
                "day": day,
                "หัวหน้ากะ": leader if leader else "(ไม่ระบุ)",
                "ยอดไม้เข้าขัด(CU)": cu_val,
                "ตำหนิหลังขัด(คิว)": defect_val,
                "จำนวนที่เข้าขัด(Pcs)": qty_pcs_val,
            }
            row_tech_entries = []
            for c, info in defect_cols.items():
                label = info["label"]
                raw_v = sheet.cell_value(r, c)
                qty, techs = _parse_defect_cell(raw_v)
                row[f"ตำหนิ::{label}"] = qty
                for tech_name, tech_qty in techs:
                    row_tech_entries.append({
                        "date": date,
                        "day": day,
                        "defect_type": label,
                        "technician": tech_name,
                        "qty": tech_qty,
                    })

            day_rows.append(row)
            day_tech_rows.extend(row_tech_entries)

        # ถ้าทั้งวันนั้น ยอดเข้าขัดรวม = 0 และ ตำหนิรวม = 0 -> ถือว่าไม่มีชุดข้อมูลจริง ตัดทิ้งทั้งวัน
        if day_cu_total == 0 and day_defect_total == 0:
            continue

        recorded_days.add(day)
        all_rows.extend(day_rows)
        tech_rows.extend(day_tech_rows)

    df = pd.DataFrame(all_rows) if all_rows else pd.DataFrame()
    tech_df = pd.DataFrame(tech_rows) if tech_rows else pd.DataFrame(
        columns=["date", "day", "defect_type", "technician", "qty"]
    )
    return df, tech_df, recorded_days, defect_label_category


def buddhist_year_from_filename_hint(default_be: int = None) -> int:
    return default_be


def categorize_defect_label(label: str, category: str) -> str:
    """
    Map one defect-type label + its file-native row-10 category into one of
    4 report groups: "Press", "เครื่องขัด", "แตกรถยก", "อื่นๆ".

    Rule (per user's spec):
      - label == "แตกรถยก" -> always its own group, regardless of file category
      - category == "ตำหนิจาก Press" -> "Press"
      - category == "ตำหนิเครื่องขัด" -> "เครื่องขัด" (แตกรถยก already carved out above)
      - everything else (category "อื่นๆ", or the bare "ตำหนิ" category e.g.
        "รอยกดแตกร้าว") -> "อื่นๆ"
    """
    if label == "แตกรถยก":
        return "แตกรถยก"
    if category == "ตำหนิจาก Press":
        return "Press"
    if category == "ตำหนิเครื่องขัด":
        return "เครื่องขัด"
    return "อื่นๆ"
