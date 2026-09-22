import streamlit as st
import pandas as pd
import plotly.express as px
import tempfile
import os
import zipfile
from datetime import datetime
from parser import parse_workbook, categorize_defect_label
from report import generate_pdf_report, pdf_to_jpgs

st.set_page_config(page_title="Dashboard ตำหนิหลังขัด", layout="wide", page_icon="📊")

# ---------- Upload & parse ----------
with st.sidebar:
    st.header("📁 อัปโหลดไฟล์")
    uploaded = st.file_uploader("ไฟล์ Excel รายเดือน (.xls)", type=["xls"])
    production_line = st.selectbox("ไลน์ผลิต", ["MDF LINE 1", "MDF LINE 2", "MDF LINE 5"])
    col_m, col_y = st.columns(2)
    with col_m:
        month = st.number_input("เดือน", min_value=1, max_value=12, value=9)
    with col_y:
        year_be = st.number_input("ปี (พ.ศ.)", min_value=2560, max_value=2600, value=2569)

st.title(f"📊 รายงานสรุปตำหนิหลังขัด {production_line}")
st.caption("Metro M.D.F. — วิเคราะห์ตำหนิหลังขัดจากรายงานไม้เข้าขัดประจำวัน")

if uploaded is None:
    st.info("⬅️ อัปโหลดไฟล์ Excel รายเดือน (แยก Sheet ตามวันที่ 1-31) ในแถบด้านซ้ายเพื่อเริ่มต้น")
    st.stop()

with st.spinner("กำลังอ่านไฟล์..."):
    try:
        df, tech_df, recorded_days, defect_label_category = parse_workbook(uploaded.read(), year_be=int(year_be), month=int(month))
    except Exception as e:
        st.error(f"ไม่สามารถอ่านไฟล์ได้: {e}")
        st.stop()

if df.empty:
    st.warning("ไม่พบข้อมูลตำหนิในไฟล์ที่อัปโหลด กรุณาตรวจสอบไฟล์และเดือน/ปีที่เลือก")
    st.stop()

defect_cols = [c for c in df.columns if c.startswith("ตำหนิ::")]

# ---------- Date filter ----------
min_date, max_date = df["date"].min().date(), df["date"].max().date()
st.sidebar.header("📅 เลือกช่วงวันที่")
date_range = st.sidebar.date_input(
    "ช่วงวันที่ที่ต้องการดู",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date,
)
if isinstance(date_range, tuple) and len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date, max_date

mask = (df["date"].dt.date >= start_date) & (df["date"].dt.date <= end_date)
fdf = df[mask].copy()

if not tech_df.empty:
    tmask = (tech_df["date"].dt.date >= start_date) & (tech_df["date"].dt.date <= end_date)
    ftech_df = tech_df[tmask].copy()
else:
    ftech_df = tech_df

if fdf.empty:
    st.warning("ไม่มีข้อมูลในช่วงวันที่ที่เลือก")
    st.stop()

def be(d):
    return f"{d.day:02d}/{d.month:02d}/{(d.year + 543) % 100:02d}"

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม",
]

def _period_label(start_date, end_date):
    """
    คืนค่า (period_text, is_monthly)
    - ถ้าช่วงวันที่อยู่ในเดือนเดียวกัน และจำนวนวัน (start ถึง end) >= 28 วัน
      -> "ประจำเดือน <ชื่อเดือนไทย> <ปี พ.ศ.>" , is_monthly=True
    - ไม่เช่นนั้น -> ช่วงวันที่แบบเดิม "ช่วงวันที่: dd/mm/yy - dd/mm/yy", is_monthly=False
    """
    n_selected_days = (end_date - start_date).days + 1
    if start_date.month == end_date.month and n_selected_days >= 28:
        month_name = THAI_MONTHS[start_date.month]
        year_be_label = start_date.year + 543
        return f"ประจำเดือน {month_name} {year_be_label}", True
    else:
        text = f"ช่วงวันที่: {be(start_date)} - {be(end_date)}"
        return text, False

# นับ "วันที่มีข้อมูล" เฉพาะวันในช่วงที่กรอง (ใช้ recorded_days จาก parser ซึ่งตัด Sheet
# ที่เป็น template ว่างเปล่าออกแล้ว แต่ยังนับวันที่ระบุชัดว่า "ไม่ได้ขัดไม้" ด้วย)
days_in_range = set(pd.date_range(start_date, end_date).day) if start_date.month == end_date.month else None
if days_in_range is not None:
    n_days_with_data = len(recorded_days & days_in_range)
else:
    # ช่วงวันที่คร่อมเดือน (ไม่ควรเกิดในไฟล์นี้ที่มีเดือนเดียว) -> fallback ใช้ recorded_days ทั้งหมด
    n_days_with_data = len(recorded_days)

_norm_start = start_date if hasattr(start_date, 'day') else pd.Timestamp(start_date)
_norm_end = end_date if hasattr(end_date, 'day') else pd.Timestamp(end_date)
period_label_text, is_monthly_period = _period_label(_norm_start, _norm_end)

st.caption(f"{period_label_text} ({n_days_with_data} วันที่มีข้อมูล)")

# สีกราฟ: ม่วง เมื่อเป็นรายงานประจำเดือน (28-31 วันในเดือนเดียวกัน), เขียวเดิมเมื่อเป็นช่วงวันที่ทั่วไป
CHART_COLOR = "#09093C" if is_monthly_period else "#009B77"


def _section_header(text):
    """
    หัวข้อ section แบบ custom (แทน st.subheader) เพื่อคุมสีให้เปลี่ยนเป็นม่วงได้
    เมื่อเป็นรายงานประจำเดือน (is_monthly_period=True)
    """
    color = "#09093C" if is_monthly_period else "#009B77"
    st.markdown(
        f'<h3 style="color:{color}; font-size:1.5rem; font-weight:600; margin-top:0.5rem; margin-bottom:0.5rem;">{text}</h3>',
        unsafe_allow_html=True,
    )


def _table_header_style(styler):
    """
    บังคับสีพื้นหลัง/ตัวหนังสือหัวตารางให้เป็นฟ้าเมื่อเป็นรายงานประจำเดือน
    เขียวตามปกติเมื่อไม่ใช่ (ใช้ set_table_styles คุม <th> โดยตรง ไม่ผูกกับ config.toml ของ Streamlit)
    """
    header_bg = "#09093C" if is_monthly_period else "#009B77"
    header_text = "white"
    return styler.set_table_styles(
        [{"selector": "th", "props": [("background-color", header_bg), ("color", header_text)]}],
        overwrite=False,
    )

# ---------- KPI: % ตำหนิหลังขัด ----------
DEFECT_TARGET_PCT = 3.0

def _kpi_font_size(value_text, base_rem=2.1, min_rem=1.1):
    """
    คำนวณขนาดฟอนต์ (rem) ให้ตัวเลขพอดี 1 บรรทัดเสมอ โดยลดขนาดลงตามความยาวข้อความ
    ค่าเริ่มต้น (<=4 ตัวอักษร) ใช้ base_rem เต็ม แล้วลดลงทีละขั้นตามจำนวนตัวอักษรที่เกิน
    (ใช้คู่กับ white-space:nowrap เพื่อไม่ให้ตกบรรทัดแม้ประมาณขนาดคลาดเคลื่อนเล็กน้อย)
    """
    n = len(value_text)
    if n <= 4:
        size = base_rem
    else:
        size = base_rem - (n - 4) * 0.11
    return max(size, min_rem)


def _kpi_box_html(value_text, label_text, box_bg, box_text_color, box_subtext_color):
    font_size = _kpi_font_size(value_text)
    return (
        f'<div style="background:{box_bg}; border-radius:8px; padding:16px 20px; min-height:88px; '
        f'display:flex; flex-direction:column; justify-content:center;">'
        f'<div style="font-size:{font_size}rem; font-weight:700; color:{box_text_color}; '
        f'line-height:1.2; white-space:nowrap;">{value_text}</div>'
        f'<div style="font-size:0.875rem; color:{box_subtext_color}; margin-top:2px;">{label_text}</div>'
        f'</div>'
    )


total_cu = fdf["ยอดไม้เข้าขัด(CU)"].sum()
total_defect = fdf["ตำหนิหลังขัด(คิว)"].sum()
pct_defect = (total_defect / total_cu * 100) if total_cu > 0 else 0
is_over_target = pct_defect > DEFECT_TARGET_PCT

k1, k2, k3 = st.columns(3)

with k1:
    if is_over_target:
        box_bg = "linear-gradient(135deg, #FF6B6B, #FF3B7F)"
        box_text = "#FFFFFF"
        box_subtext = "#FFE8EE"
    else:
        box_bg = "#F0F2F6"
        box_text = "#31333F"
        box_subtext = "#5C5F6D"

    st.markdown(
        f'<div style="position:relative; margin-top:14px;">'
        f'<div style="position:absolute; top:-14px; right:10px; font-size:0.72rem; color:#5C5F6D; background:#FFFFFF; padding:0 4px;">Target &le; {DEFECT_TARGET_PCT:.0f}%</div>'
        f'{_kpi_box_html(f"{pct_defect:.2f} %", "% ตำหนิหลังขัด", box_bg, box_text, box_subtext)}'
        f'</div>',
        unsafe_allow_html=True,
    )

with k2:
    kpi2_bg = "#E2E2ED" if is_monthly_period else "#F0F2F6"
    kpi2_text = "#09093C" if is_monthly_period else "#31333F"
    kpi2_subtext = "#5C5F6D"
    st.markdown(
        f'<div style="margin-top:14px;">'
        f'{_kpi_box_html(f"{total_cu:,.2f}", "ยอดไม้เข้าขัดทั้งหมด (m³)", kpi2_bg, kpi2_text, kpi2_subtext)}'
        f'</div>',
        unsafe_allow_html=True,
    )

with k3:
    kpi3_bg = "#E2E2ED" if is_monthly_period else "#F0F2F6"
    kpi3_text = "#09093C" if is_monthly_period else "#31333F"
    kpi3_subtext = "#5C5F6D"
    st.markdown(
        f'<div style="margin-top:14px;">'
        f'{_kpi_box_html(f"{total_defect:,.2f}", "ตำหนิหลังขัดทั้งหมด (m³)", kpi3_bg, kpi3_text, kpi3_subtext)}'
        f'</div>',
        unsafe_allow_html=True,
    )

st.divider()

# ---------- Top 5 defects ----------
left, right = st.columns([1.1, 1])

with left:
    _section_header("🔝 Top 5 ตำหนิ (จำนวนแผ่นรวม)")
    defect_sums = fdf[defect_cols].sum().sort_values(ascending=False)
    defect_sums = defect_sums[defect_sums > 0]
    top5 = defect_sums.head(5)
    top5_display = top5.rename(index=lambda x: x.replace("ตำหนิ::", ""))

    if top5_display.empty:
        st.info("ไม่มีข้อมูลตำหนิรายชนิดในช่วงวันที่นี้")
    else:
        fig = px.bar(
            top5_display[::-1],
            orientation="h",
            labels={"value": "จำนวนแผ่น (Pcs)", "index": "ประเภทตำหนิ"},
            text=top5_display[::-1].values,
            color_discrete_sequence=[CHART_COLOR],
        )
        fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig.update_layout(showlegend=False, height=350, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

        top5_table = top5_display.reset_index()
        top5_table.columns = ["ประเภทตำหนิ", "จำนวนแผ่น (Pcs)"]
        top5_table.index = top5_table.index + 1
        st.dataframe(_table_header_style(top5_table.style), use_container_width=True)

# ---------- % ตำหนิแยกตามหัวหน้ากะ ----------
with right:
    _section_header("👷 % ตำหนิหลังขัด แยกตามหัวหน้ากะ")
    leader_grp = fdf.groupby("หัวหน้ากะ")[["ยอดไม้เข้าขัด(CU)", "ตำหนิหลังขัด(คิว)"]].sum()
    leader_grp = leader_grp[leader_grp["ยอดไม้เข้าขัด(CU)"] > 0]  # ตัดหัวหน้ากะที่ไม่มียอดเข้าขัด (หารไม่ได้)
    leader_grp["% ตำหนิ"] = (leader_grp["ตำหนิหลังขัด(คิว)"] / leader_grp["ยอดไม้เข้าขัด(CU)"] * 100).round(2)
    leader_grp = leader_grp.sort_values("% ตำหนิ", ascending=False)
    leader_grp = leader_grp.rename(columns={
        "ยอดไม้เข้าขัด(CU)": "ยอดไม้เข้าขัด (m³)",
        "ตำหนิหลังขัด(คิว)": "ตำหนิหลังขัด (m³)",
    })
    leader_grp = leader_grp.reset_index()
    leader_grp.index = leader_grp.index + 1

    def _highlight_over_target_row(row):
        if row["% ตำหนิ"] > DEFECT_TARGET_PCT:
            return ["color:#E0304F; font-weight:700;"] * len(row)
        return [""] * len(row)

    st.dataframe(
        _table_header_style(
            leader_grp.style.format({
                "ยอดไม้เข้าขัด (m³)": "{:,.2f}",
                "ตำหนิหลังขัด (m³)": "{:,.2f}",
                "% ตำหนิ": "{:.2f}%",
            }).apply(_highlight_over_target_row, axis=1)
        ),
        use_container_width=True,
    )

    fig2 = px.bar(
        leader_grp,
        x="หัวหน้ากะ",
        y="% ตำหนิ",
        text="% ตำหนิ",
        color_discrete_sequence=[CHART_COLOR],
    )
    fig2.update_traces(texttemplate="%{text:.2f}%", textposition="outside")
    fig2.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig2, use_container_width=True)

    st.caption("รายละเอียดตำหนิแยกตามหัวหน้ากะและประเภทตำหนิ")
    leader_defect_long = fdf.melt(
        id_vars=["หัวหน้ากะ"],
        value_vars=defect_cols,
        var_name="ประเภทตำหนิ",
        value_name="จำนวนแผ่น (Pcs)",
    )
    leader_defect_long["ประเภทตำหนิ"] = leader_defect_long["ประเภทตำหนิ"].str.replace("ตำหนิ::", "", regex=False)
    leader_defect_detail = (
        leader_defect_long.groupby(["หัวหน้ากะ", "ประเภทตำหนิ"])["จำนวนแผ่น (Pcs)"]
        .sum()
        .reset_index()
    )
    leader_defect_detail = leader_defect_detail[leader_defect_detail["จำนวนแผ่น (Pcs)"] > 0]
    leader_defect_detail = leader_defect_detail.sort_values(
        ["หัวหน้ากะ", "จำนวนแผ่น (Pcs)"], ascending=[True, False]
    )
    leader_defect_detail.index = range(1, len(leader_defect_detail) + 1)
    st.dataframe(_table_header_style(leader_defect_detail.style), use_container_width=True, height=350)

st.divider()

# ---------- ตำหนิจากเครื่องขัด ----------
_section_header("🔧 ตำหนิจากเครื่องขัด")
tech_sum = pd.DataFrame(columns=["Operator เครื่องขัด", "จำนวนแผ่น (Pcs)"])
defect_type_sum = pd.DataFrame(columns=["ประเภทตำหนิ", "จำนวนแผ่น (Pcs)"])

if ftech_df.empty:
    st.info("ไม่มีข้อมูลระบุชื่อ Operator เครื่องขัดในช่วงวันที่นี้")
else:
    tech_col, detail_col = st.columns([1, 1.3])

    with tech_col:
        tech_sum = (
            ftech_df.groupby("technician")["qty"]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        tech_sum.columns = ["Operator เครื่องขัด", "จำนวนแผ่น (Pcs)"]
        tech_sum.index = tech_sum.index + 1
        st.dataframe(_table_header_style(tech_sum.style), use_container_width=True)

        defect_type_sum = (
            ftech_df.groupby("defect_type")["qty"]
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        defect_type_sum.columns = ["ประเภทตำหนิ", "จำนวนแผ่น (Pcs)"]

        fig3 = px.bar(
            defect_type_sum,
            x="ประเภทตำหนิ",
            y="จำนวนแผ่น (Pcs)",
            text="จำนวนแผ่น (Pcs)",
            color_discrete_sequence=[CHART_COLOR],
        )
        fig3.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
        fig3.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fig3, use_container_width=True)

    with detail_col:
        st.caption("รายละเอียดตำหนิแยกตามช่างและประเภทตำหนิ")
        detail = (
            ftech_df.groupby(["technician", "defect_type"])["qty"]
            .sum()
            .reset_index()
            .sort_values(["technician", "qty"], ascending=[True, False])
        )
        detail.columns = ["Operator เครื่องขัด", "ประเภทตำหนิ", "จำนวนแผ่น (Pcs)"]
        detail.index = range(1, len(detail) + 1)
        st.dataframe(_table_header_style(detail.style), use_container_width=True, height=350)

st.divider()

# ---------- สรุปตำหนิแยกตามแหล่งที่มา (Press / เครื่องขัด / แตกรถยก / อื่นๆ) ----------
_section_header("📦 สรุปตำหนิแยกตามแหล่งที่มา")

GROUP_ORDER = ["Press", "เครื่องขัด", "แตกรถยก", "อื่นๆ"]

group_pcs = {g: 0.0 for g in GROUP_ORDER}
for col in defect_cols:
    label = col.replace("ตำหนิ::", "")
    category = defect_label_category.get(label, "")
    group = categorize_defect_label(label, category)
    if group in group_pcs:
        group_pcs[group] += fdf[col].sum()

total_qty_in_pcs = fdf["จำนวนที่เข้าขัด(Pcs)"].sum()
source_rows = []
for g in GROUP_ORDER:
    pcs = group_pcs[g]
    pct_of_qty_in = (pcs / total_qty_in_pcs * 100) if total_qty_in_pcs > 0 else 0
    source_rows.append({"แหล่งที่มาตำหนิ": g, "จำนวนแผ่น (Pcs)": pcs, "% ตำหนิรวม": pct_of_qty_in})

source_summary_df = pd.DataFrame(source_rows)
source_summary_df.index = source_summary_df.index + 1

def _highlight_over_target_source_row(row):
    if row["% ตำหนิรวม"] > DEFECT_TARGET_PCT:
        return ["color:#E0304F; font-weight:700;"] * len(row)
    return [""] * len(row)

st.dataframe(
    _table_header_style(
        source_summary_df.style.format({
            "จำนวนแผ่น (Pcs)": "{:,.0f}",
            "% ตำหนิรวม": "{:.2f}%",
        }).apply(_highlight_over_target_source_row, axis=1)
    ),
    use_container_width=True,
)

st.divider()
with st.expander("📄 ดูข้อมูลดิบ (หลังกรองตามวันที่)"):
    st.dataframe(fdf.drop(columns=[c for c in defect_cols]), use_container_width=True)

# ---------- ดาวน์โหลดรายงานเป็นรูปภาพ JPG ----------
st.divider()
st.subheader("📥 ดาวน์โหลดรายงานสำหรับผู้บริหาร")
st.caption("สร้างรายงานสรุปเป็นรูปภาพ JPG จัดรูปแบบสวยงาม ตามช่วงวันที่ที่เลือกอยู่ด้านบน")

period_text = period_label_text

if st.button("🖼️ สร้างรายงาน JPG", type="primary"):
    with st.spinner("กำลังสร้างรายงาน..."):
        leader_for_pdf = leader_grp.copy()  # already: หัวหน้ากะ, ยอดไม้เข้าขัด (Cu.), ตำหนิหลังขัด (คิว), % ตำหนิ
        tech_summary_for_pdf = tech_sum.copy()
        tech_defecttype_for_pdf = defect_type_sum.copy()
        source_summary_for_pdf = source_summary_df.copy()

        pdf_path = os.path.join(tempfile.gettempdir(), "sanding_defect_report.pdf")
        generate_pdf_report(
            out_path=pdf_path,
            period_text=period_text,
            pct_defect=pct_defect,
            total_cu=total_cu,
            total_defect=total_defect,
            n_days=n_days_with_data,
            defect_target_pct=DEFECT_TARGET_PCT,
            top5_series=top5_display,
            leader_df=leader_for_pdf,
            tech_summary_df=tech_summary_for_pdf,
            tech_defecttype_df=tech_defecttype_for_pdf,
            source_summary_df=source_summary_for_pdf,
            production_line=production_line,
            generated_at_text=datetime.now().strftime("%d/%m/%Y %H:%M"),
            is_monthly=is_monthly_period,
        )

        jpg_paths = pdf_to_jpgs(pdf_path, tempfile.gettempdir(), base_name="sanding_defect_report")

    st.success("สร้างรายงานสำเร็จ")

    if len(jpg_paths) == 1:
        with open(jpg_paths[0], "rb") as f:
            jpg_bytes = f.read()
        st.download_button(
            label="⬇️ ดาวน์โหลด JPG",
            data=jpg_bytes,
            file_name="รายงานสรุปตำหนิหลังขัด.jpg",
            mime="image/jpeg",
        )
    else:
        zip_path = os.path.join(tempfile.gettempdir(), "sanding_defect_report_jpg.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, p in enumerate(jpg_paths, start=1):
                zf.write(p, arcname=f"รายงานสรุปตำหนิหลังขัด_หน้า{i}.jpg")
        with open(zip_path, "rb") as f:
            zip_bytes = f.read()
        st.download_button(
            label="⬇️ ดาวน์โหลด JPG (ทั้งหมด, .zip)",
            data=zip_bytes,
            file_name="รายงานสรุปตำหนิหลังขัด_jpg.zip",
            mime="application/zip",
        )
        preview_cols = st.columns(len(jpg_paths))
        for col, p in zip(preview_cols, jpg_paths):
            with col:
                st.image(p, use_container_width=True)
