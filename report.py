"""
Generate a polished executive-summary PDF report of the sanding-defect dashboard
for the currently selected date range. Designed to be sent to management:
clean layout, veridian-green theme, Thai text via bundled Sarabun font.
"""

import io
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pymupdf

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, HRFlowable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------- Theme: Viridian green ----------
VIRIDIAN = colors.HexColor("#009B77")
VIRIDIAN_DARK = colors.HexColor("#00694A")
VIRIDIAN_LIGHT = colors.HexColor("#E6F5F0")
GREY_TEXT = colors.HexColor("#333333")
GREY_LINE = colors.HexColor("#CCCCCC")

_FONT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
_FONT_REGULAR_PATH = os.path.join(_FONT_DIR, "Sarabun-Regular.ttf")
_FONT_BOLD_PATH = os.path.join(_FONT_DIR, "Sarabun-Bold.ttf")

_fonts_registered = False


def _ensure_fonts():
    global _fonts_registered
    if _fonts_registered:
        return
    pdfmetrics.registerFont(TTFont("Sarabun", _FONT_REGULAR_PATH))
    pdfmetrics.registerFont(TTFont("Sarabun-Bold", _FONT_BOLD_PATH))
    # matplotlib
    fm.fontManager.addfont(_FONT_REGULAR_PATH)
    prop = fm.FontProperties(fname=_FONT_REGULAR_PATH)
    plt.rcParams["font.family"] = prop.get_name()
    plt.rcParams["axes.unicode_minus"] = False
    _fonts_registered = True


def _styles():
    ss = getSampleStyleSheet()
    styles = {
        "title": ParagraphStyle(
            "ReportTitle", fontName="Sarabun-Bold", fontSize=20,
            textColor=colors.white, alignment=TA_LEFT, leading=26,
        ),
        "subtitle": ParagraphStyle(
            "ReportSubtitle", fontName="Sarabun", fontSize=11,
            textColor=colors.HexColor("#D8F3EA"), alignment=TA_LEFT, leading=15,
        ),
        "section": ParagraphStyle(
            "Section", fontName="Sarabun-Bold", fontSize=13,
            textColor=VIRIDIAN_DARK, spaceBefore=14, spaceAfter=8, leading=17,
        ),
        "body": ParagraphStyle(
            "Body", fontName="Sarabun", fontSize=10, textColor=GREY_TEXT, leading=14,
        ),
        "kpi_value": ParagraphStyle(
            "KpiValue", fontName="Sarabun-Bold", fontSize=26,
            textColor=VIRIDIAN_DARK, alignment=TA_CENTER, leading=30,
        ),
        "kpi_label": ParagraphStyle(
            "KpiLabel", fontName="Sarabun", fontSize=9.5,
            textColor=GREY_TEXT, alignment=TA_CENTER, leading=12,
        ),
        "footer": ParagraphStyle(
            "Footer", fontName="Sarabun", fontSize=8,
            textColor=colors.HexColor("#888888"), alignment=TA_CENTER,
        ),
    }
    return styles


def _make_bar_chart_png(labels, values, color=VIRIDIAN.hexval()[2:] if hasattr(VIRIDIAN, "hexval") else "#009B77",
                          title="", horizontal=False, figsize=(6.6, 3.2)):
    _ensure_fonts()
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    bar_color = "#009B77"
    if horizontal:
        y_pos = range(len(labels))
        ax.barh(list(y_pos), values, color=bar_color)
        ax.set_yticks(list(y_pos))
        ax.set_yticklabels(labels, fontsize=9)
        ax.invert_yaxis()
        for i, v in enumerate(values):
            ax.text(v, i, f" {v:,.0f}", va="center", fontsize=8.5, color="#00694A")
        ax.spines[["top", "right"]].set_visible(False)
    else:
        x_pos = range(len(labels))
        ax.bar(list(x_pos), values, color=bar_color)
        ax.set_xticks(list(x_pos))
        ax.set_xticklabels(labels, fontsize=8, rotation=30, ha="right", rotation_mode="anchor")
        for i, v in enumerate(values):
            ax.text(i, v, f"{v:,.0f}", ha="center", va="bottom", fontsize=8.5, color="#00694A")
        ax.spines[["top", "right"]].set_visible(False)
        ax.margins(y=0.15)

    if title:
        ax.set_title(title, fontsize=11, color="#00694A", pad=10, fontweight="bold")
    ax.tick_params(colors="#444444")
    fig.tight_layout()
    if not horizontal:
        fig.subplots_adjust(bottom=0.32)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _df_to_table(df_display, col_widths, header_bg=VIRIDIAN, align_last_right=True, red_row_indices=None):
    data = [list(df_display.columns)] + df_display.values.tolist()
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style = [
        ("FONTNAME", (0, 0), (-1, 0), "Sarabun-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Sarabun"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("BACKGROUND", (0, 0), (-1, 0), header_bg),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT" if align_last_right else "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, GREY_LINE),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, VIRIDIAN_LIGHT]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]
    if red_row_indices:
        red = colors.HexColor("#E0304F")
        for idx in red_row_indices:
            row_num = idx + 1  # +1 to account for header row
            style.append(("TEXTCOLOR", (0, row_num), (-1, row_num), red))
            style.append(("FONTNAME", (0, row_num), (-1, row_num), "Sarabun-Bold"))
    t.setStyle(TableStyle(style))
    return t


def _header_footer(canvas, doc, title_text, period_text):
    canvas.saveState()
    page_w, page_h = A4

    # Top banner
    canvas.setFillColor(VIRIDIAN)
    canvas.rect(0, page_h - 32 * mm, page_w, 32 * mm, fill=1, stroke=0)
    canvas.setFillColor(VIRIDIAN_DARK)
    canvas.rect(0, page_h - 32 * mm, page_w, 3, fill=1, stroke=0)

    canvas.setFont("Sarabun-Bold", 18)
    canvas.setFillColor(colors.white)
    canvas.drawString(20 * mm, page_h - 15 * mm, title_text)

    canvas.setFont("Sarabun", 10.5)
    canvas.setFillColor(colors.HexColor("#D8F3EA"))
    canvas.drawString(20 * mm, page_h - 22 * mm, period_text)

    # Footer
    canvas.setFont("Sarabun", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawCentredString(page_w / 2, 10 * mm, f"หน้า {doc.page}")
    canvas.drawString(20 * mm, 10 * mm, "Metro M.D.F. — QC/QA Department")

    canvas.restoreState()


def generate_pdf_report(
    out_path,
    period_text,
    pct_defect,
    total_cu,
    total_defect,
    n_days,
    top5_series,          # pandas Series: index=defect name, values=qty (already top 5, descending)
    leader_df,             # DataFrame with columns: หัวหน้ากะ, ยอดไม้เข้าขัด (m³), ตำหนิหลังขัด (m³), % ตำหนิ
    tech_summary_df,       # DataFrame with columns: ช่างเครื่องขัด, จำนวนแผ่น (Pcs)  (may be empty)
    tech_defecttype_df,    # DataFrame with columns: ประเภทตำหนิ, จำนวนแผ่น (Pcs)   (may be empty)
    generated_at_text,
    defect_target_pct=3.0,
    source_summary_df=None,  # DataFrame: แหล่งที่มาตำหนิ, จำนวนแผ่น (Pcs), % ตำหนิรวม
    production_line=None,   # e.g. "MDF LINE 2" -- appended to the report title when given
):
    _ensure_fonts()
    styles = _styles()

    report_title = "รายงานสรุปตำหนิหลังขัด" + (f" {production_line}" if production_line else "")

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        topMargin=38 * mm, bottomMargin=18 * mm,
        leftMargin=18 * mm, rightMargin=18 * mm,
        title=report_title,
    )

    story = []

    # ---------- KPI row ----------
    is_over_target = pct_defect > defect_target_pct

    if is_over_target:
        pct_box_bg = colors.HexColor("#FF3B7F")
        pct_box_border = colors.HexColor("#C40052")
        pct_value_color = colors.white
        pct_label_color = colors.HexColor("#FFE3ED")
    else:
        pct_box_bg = VIRIDIAN_LIGHT
        pct_box_border = VIRIDIAN
        pct_value_color = VIRIDIAN_DARK
        pct_label_color = GREY_TEXT

    pct_value_style = ParagraphStyle(
        "PctKpiValue", parent=styles["kpi_value"], textColor=pct_value_color,
    )
    pct_label_style = ParagraphStyle(
        "PctKpiLabel", parent=styles["kpi_label"], textColor=pct_label_color,
    )
    pct_target_style = ParagraphStyle(
        "PctKpiTarget", fontName="Sarabun", fontSize=7.5,
        textColor=colors.HexColor("#5C5F6D"), alignment=TA_RIGHT,
    )

    # แถวป้าย "Target <= 3%" ลอยอยู่เหนือกล่องแรก (มุมบนขวาด้านนอกกล่อง)
    target_label_row = Table(
        [[Paragraph(f"Target &le; {defect_target_pct:.0f}%", pct_target_style), "", "", ""]],
        colWidths=[42 * mm] * 4,
    )
    target_label_row.setStyle(TableStyle([
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(target_label_row)

    pct_cell_data = [
        [Paragraph(f"{pct_defect:.2f}%", pct_value_style)],
        [Paragraph("% ตำหนิหลังขัด", pct_label_style)],
    ]
    pct_cell_table = Table(pct_cell_data, colWidths=[42 * mm])
    pct_cell_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pct_box_bg),
        ("BOX", (0, 0), (-1, -1), 1, pct_box_border),
        ("TOPPADDING", (0, 0), (-1, 0), 14),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
        ("TOPPADDING", (0, 1), (-1, 1), 2),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 12),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))

    kpi_data = [
        [
            pct_cell_table,
            Paragraph(f"{total_cu:,.2f}", styles["kpi_value"]),
            Paragraph(f"{total_defect:,.2f}", styles["kpi_value"]),
            Paragraph(f"{n_days}", styles["kpi_value"]),
        ],
        [
            "",
            Paragraph("ยอดไม้เข้าขัดทั้งหมด (m³)", styles["kpi_label"]),
            Paragraph("ตำหนิหลังขัดทั้งหมด (m³)", styles["kpi_label"]),
            Paragraph("จำนวนวันที่มีข้อมูล", styles["kpi_label"]),
        ],
    ]
    kpi_table = Table(kpi_data, colWidths=[42 * mm] * 4)
    kpi_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (0, 1)),
        ("BACKGROUND", (1, 0), (-1, -1), VIRIDIAN_LIGHT),
        ("BOX", (1, 0), (1, -1), 1, VIRIDIAN),
        ("BOX", (2, 0), (2, -1), 1, VIRIDIAN),
        ("BOX", (3, 0), (3, -1), 1, VIRIDIAN),
        ("TOPPADDING", (1, 0), (-1, 0), 14),
        ("BOTTOMPADDING", (1, 0), (-1, 0), 2),
        ("TOPPADDING", (1, 1), (-1, 1), 2),
        ("BOTTOMPADDING", (1, 1), (-1, 1), 12),
        ("VALIGN", (0, 0), (0, 1), "TOP"),
        ("LEFTPADDING", (0, 0), (0, 1), 0),
        ("RIGHTPADDING", (0, 0), (0, 1), 0),
        ("TOPPADDING", (0, 0), (0, 1), 0),
        ("BOTTOMPADDING", (0, 0), (0, 1), 0),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 14))

    # ---------- Top 5 defects ----------
    story.append(Paragraph("Top 5 ตำหนิ (จำนวนแผ่นรวม)", styles["section"]))
    story.append(HRFlowable(width="100%", thickness=1, color=VIRIDIAN, spaceAfter=8))

    if top5_series is not None and len(top5_series) > 0:
        labels = list(top5_series.index)
        values = list(top5_series.values)
        chart_buf = _make_bar_chart_png(labels, values, horizontal=True, figsize=(6.6, 2.6))
        story.append(Image(chart_buf, width=165 * mm, height=65 * mm))
        story.append(Spacer(1, 6))

        top5_df = top5_series.reset_index()
        top5_df.columns = ["ประเภทตำหนิ", "จำนวนแผ่น (Pcs)"]
        top5_df["จำนวนแผ่น (Pcs)"] = top5_df["จำนวนแผ่น (Pcs)"].map(lambda x: f"{x:,.0f}")
        story.append(_df_to_table(top5_df, col_widths=[120 * mm, 50 * mm]))
    else:
        story.append(Paragraph("ไม่มีข้อมูลตำหนิรายชนิดในช่วงวันที่นี้", styles["body"]))

    story.append(Spacer(1, 16))

    # ---------- Leader breakdown ----------
    story.append(Paragraph("% ตำหนิหลังขัด แยกตามหัวหน้ากะ", styles["section"]))
    story.append(HRFlowable(width="100%", thickness=1, color=VIRIDIAN, spaceAfter=8))

    if leader_df is not None and len(leader_df) > 0:
        ld = leader_df.reset_index() if leader_df.index.name == "หัวหน้ากะ" else leader_df.copy()
        if "หัวหน้ากะ" not in ld.columns:
            ld = ld.rename(columns={ld.columns[0]: "หัวหน้ากะ"})
        display_cols = ["หัวหน้ากะ", "ยอดไม้เข้าขัด (m³)", "ตำหนิหลังขัด (m³)", "% ตำหนิ"]
        ld = ld[[c for c in display_cols if c in ld.columns]]

        red_row_indices = set()
        if "% ตำหนิ" in ld.columns:
            for i, v in enumerate(ld["% ตำหนิ"]):
                if v > defect_target_pct:
                    red_row_indices.add(i)

        for c in ["ยอดไม้เข้าขัด (m³)", "ตำหนิหลังขัด (m³)"]:
            if c in ld.columns:
                ld[c] = ld[c].map(lambda x: f"{x:,.2f}")
        if "% ตำหนิ" in ld.columns:
            ld["% ตำหนิ"] = ld["% ตำหนิ"].map(lambda x: f"{x:.2f}%")
        story.append(_df_to_table(
            ld, col_widths=[45 * mm, 45 * mm, 45 * mm, 35 * mm],
            red_row_indices=red_row_indices,
        ))
    else:
        story.append(Paragraph("ไม่มีข้อมูลหัวหน้ากะในช่วงวันที่นี้", styles["body"]))

    story.append(Spacer(1, 16))

    # ---------- Machine operator (technician) breakdown ----------
    story.append(Paragraph("ตำหนิเครื่องขัด", styles["section"]))
    story.append(HRFlowable(width="100%", thickness=1, color=VIRIDIAN, spaceAfter=8))

    if tech_summary_df is not None and len(tech_summary_df) > 0:
        left_tbl_df = tech_summary_df.copy()
        left_tbl_df.columns = ["ช่างเครื่องขัด", "จำนวนแผ่น (Pcs)"]
        left_tbl_df["จำนวนแผ่น (Pcs)"] = left_tbl_df["จำนวนแผ่น (Pcs)"].map(lambda x: f"{x:,.0f}")
        story.append(_df_to_table(left_tbl_df, col_widths=[85 * mm, 85 * mm]))
        story.append(Spacer(1, 8))

        if tech_defecttype_df is not None and len(tech_defecttype_df) > 0:
            labels = list(tech_defecttype_df.iloc[:, 0])
            values = list(tech_defecttype_df.iloc[:, 1])
            chart_buf2 = _make_bar_chart_png(labels, values, horizontal=False, figsize=(6.6, 3.0))
            story.append(Image(chart_buf2, width=165 * mm, height=72 * mm))
    else:
        story.append(Paragraph("ไม่มีข้อมูลระบุชื่อช่างเครื่องขัดในช่วงวันที่นี้", styles["body"]))

    story.append(Spacer(1, 16))

    # ---------- สรุปตำหนิแยกตามแหล่งที่มา (Press / เครื่องขัด / แตกรถยก / อื่นๆ) ----------
    story.append(Paragraph("สรุปตำหนิแยกตามแหล่งที่มา", styles["section"]))
    story.append(HRFlowable(width="100%", thickness=1, color=VIRIDIAN, spaceAfter=8))

    if source_summary_df is not None and len(source_summary_df) > 0:
        sd = source_summary_df.copy()
        sd.columns = ["แหล่งที่มาตำหนิ", "จำนวนแผ่น (Pcs)", "% ตำหนิรวม"]

        source_red_rows = set()
        for i, v in enumerate(sd["% ตำหนิรวม"]):
            if v > defect_target_pct:
                source_red_rows.add(i)

        sd["จำนวนแผ่น (Pcs)"] = sd["จำนวนแผ่น (Pcs)"].map(lambda x: f"{x:,.0f}")
        sd["% ตำหนิรวม"] = sd["% ตำหนิรวม"].map(lambda x: f"{x:.2f}%")
        story.append(_df_to_table(
            sd, col_widths=[65 * mm, 55 * mm, 55 * mm],
            red_row_indices=source_red_rows,
        ))
    else:
        story.append(Paragraph("ไม่มีข้อมูลตำหนิรายชนิดในช่วงวันที่นี้", styles["body"]))

    story.append(Spacer(1, 18))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GREY_LINE, spaceAfter=6))
    story.append(Paragraph(f"สร้างรายงานเมื่อ: {generated_at_text}", styles["footer"]))

    def _on_page(canvas, doc_):
        _header_footer(canvas, doc_, report_title, period_text)

    doc.build(story, onFirstPage=_on_page, onLaterPages=_on_page)
    return out_path


def pdf_to_jpgs(pdf_path, out_dir, base_name="report", dpi=200, quality=90):
    """
    Render every page of the PDF at pdf_path into a separate JPG file in
    out_dir, named "<base_name>_page<N>.jpg" (1-indexed). Returns the list of
    JPG file paths in page order.
    """
    doc = pymupdf.open(pdf_path)
    jpg_paths = []
    try:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            out_path = os.path.join(out_dir, f"{base_name}_page{i}.jpg")
            pix.save(out_path, jpg_quality=quality)
            jpg_paths.append(out_path)
    finally:
        doc.close()
    return jpg_paths
