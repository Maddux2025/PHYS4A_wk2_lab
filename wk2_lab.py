# app.py
import io
import os
import re
import textwrap
from datetime import datetime

from flask import Flask, render_template, request, send_file, abort

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.lib.pdfencrypt import StandardEncryption
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image as RLImage,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# PDF append support
try:
    from pypdf import PdfReader, PdfWriter
except Exception:
    from PyPDF2 import PdfReader, PdfWriter


# =========================
# Paths / folders
# =========================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(APP_DIR, "templates")
STATIC_DIR = os.path.join(APP_DIR, "static")
UPLOADS_DIR = os.path.join(APP_DIR, "uploads")
OUT_DIR = os.path.join(APP_DIR, "generated_pdfs")

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__, template_folder=TEMPLATES_DIR)

# (Optional) prevent huge uploads
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB


# =========================
# Branding
# =========================
HEADER_LINE = "De Anza College — PHYS 4A Lab — Spring 2026"
DEFAULT_HEADER_RIGHT = "Professor Name"


# =========================
# Static assets (filenames you confirmed)
# =========================
LOGO_PATH = os.path.join(STATIC_DIR, "deanza_logo.png")
METRIC_RULER_IMG = os.path.join(STATIC_DIR, "metric_ruler.png")
VERNIER_IMG_PAGE4 = os.path.join(STATIC_DIR, "vernier_page4.png")
VERNIER_IMG_PAGE5 = os.path.join(STATIC_DIR, "vernier_page5.png")
MICROMETER_IMG_1 = os.path.join(STATIC_DIR, "micrometer_1.png")
MICROMETER_IMG_2 = os.path.join(STATIC_DIR, "micrometer_2.png")

# =========================
# Colors
# =========================
YELLOW = colors.Color(1.0, 0.95, 0.20)
HEADER_BAR = colors.Color(0.45, 0.0, 0.0)
LIGHT_GREY = colors.Color(0.95, 0.95, 0.95)

UPLOAD_IMG_WIDTH = 7.0 * inch
UPLOAD_IMG_HEIGHT = 3.25 * inch

APP1_IMG_W = 2.3 * inch
APP1_IMG_H = 1.6 * inch
APP1_MASS_IMG_H = 0.35 * inch   # <-- ADD THIS

# =========================
# Appendix I — image box sizes
# =========================
APP1_MASS_COL_W = 3.25 * inch
APP1_MASS_HEADER_H = 0.30 * inch
#APP1_MASS_IMG_H = 1.1 * inch   # <- adjust ONLY this if needed


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}
PDF_EXTS = {".pdf"}
ALLOWED_EXTS = IMAGE_EXTS | PDF_EXTS


def safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[^a-zA-Z0-9._-]+", "_", s)
    return s or "file"


def file_ext(path: str) -> str:
    return os.path.splitext((path or "").lower())[1]


def is_image(path: str) -> bool:
    return file_ext(path) in IMAGE_EXTS


def is_pdf(path: str) -> bool:
    return file_ext(path) in PDF_EXTS


def save_upload(fs, prefix: str) -> str:
    """
    Saves an uploaded file to UPLOADS_DIR.
    Allows: pdf, png, jpg, jpeg, webp.
    Returns saved path or "".
    """
    if fs is None or not getattr(fs, "filename", ""):
        return ""
    if not fs.filename.strip():
        return ""

    ext = os.path.splitext(fs.filename)[1].lower()
    if ext not in ALLOWED_EXTS:
        # Keep strict so students don't upload .docx etc.
        raise ValueError(f"Unsupported upload type: {ext}. Allowed: {', '.join(sorted(ALLOWED_EXTS))}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{prefix}_{ts}_{safe_filename(fs.filename)}"
    out_path = os.path.join(UPLOADS_DIR, out_name)
    fs.save(out_path)
    return out_path


def fit_image_fixed_height(path: str, target_width: float, target_height: float) -> RLImage:
    """
    Places an image inside a bounding box WITHOUT stretching.
    ReportLab keeps aspect ratio with _restrictSize.
    """
    img = RLImage(path)
    img._restrictSize(target_width, target_height)
    return img

def fit_image_contain_box(path: str, box_w: float, box_h: float) -> Table:
    img = RLImage(path)
    img._restrictSize(box_w, box_h)
    img.hAlign = "CENTER"

    t = Table([[img]], colWidths=[box_w], rowHeights=[box_h])
    t.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.9, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t

def appendix_image_cell(path: str):
    if path and os.path.exists(path):
        return fit_image_contain_box(path, APP1_IMG_W, APP1_IMG_H)
    return Spacer(1, APP1_IMG_H)

def appendix_image_cell(path: str):
    if path and os.path.exists(path):
        img = fit_image_fixed_height(path, APP1_IMG_W, APP1_IMG_H)
        img.hAlign = "CENTER"
        return img
    else:
        return Spacer(1, APP1_IMG_H)


def image_with_template_box(path: str) -> Table:
    """
    Put an uploaded image inside a fixed 7in × 3.5in box with a border.
    Aspect ratio preserved (no stretching).
    """
    img = fit_image_fixed_height(path, UPLOAD_IMG_WIDTH, UPLOAD_IMG_HEIGHT)
    img.hAlign = "CENTER"

    t = Table([[img]], colWidths=[UPLOAD_IMG_WIDTH], rowHeights=[UPLOAD_IMG_HEIGHT])
    t.setStyle([
        ("BOX", (0, 0), (-1, -1), 1.2, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ])
    return t





def yellow_table_style():
    return TableStyle(
        [
            ("BOX", (0, 0), (-1, -1), 0.9, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.6, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("FONTSIZE", (0, 0), (-1, -1), 10),
        ]
    )


def caption(text: str, styles):
    return Paragraph(f"<i>{text}</i>", styles["Caption"])


def add_uploaded_block(story, styles, label: str, path: str, max_w=6.6 * inch, max_h=5.5 * inch):
    """
    Embeds image uploads inline. For PDFs, adds a note (PDF pages appended later).
    """
    if not path or not os.path.exists(path):
        return

    fname = os.path.basename(path)
    story.append(Spacer(1, 0.08 * inch))
    story.append(Paragraph(f"<b>{label}</b>: {fname}", styles["Small"]))
    story.append(Spacer(1, 0.06 * inch))

    if is_image(path):
       story.append(image_with_template_box(path))
       story.append(Spacer(1, 0.10 * inch))

    elif is_pdf(path):
        story.append(Paragraph("<i>(PDF pages will be appended to the end of the report.)</i>", styles["Small"]))
        story.append(Spacer(1, 0.10 * inch))
    else:
        story.append(Paragraph("<i>(Unsupported file type for preview.)</i>", styles["Small"]))
        story.append(Spacer(1, 0.10 * inch))


def append_pdf_uploads(main_pdf_bytes: bytes, upload_paths: list[str]) -> bytes:
    """
    Appends any uploaded PDF files to the end of the generated report.
    """
    pdf_paths = [p for p in upload_paths if p and is_pdf(p) and os.path.exists(p)]
    if not pdf_paths:
        return main_pdf_bytes

    writer = PdfWriter()
    main_reader = PdfReader(io.BytesIO(main_pdf_bytes))
    for page in main_reader.pages:
        writer.add_page(page)

    for p in pdf_paths:
        r = PdfReader(p)
        for page in r.pages:
            writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def build_pdf_wk2_lab(payload: dict, include_appendix_ii: bool, instructor_password: str = "") -> bytes:
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CenterTitle", parent=styles["Title"], alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="H1", parent=styles["Heading1"], alignment=TA_LEFT, spaceAfter=6))
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontSize=10, leading=12))
    styles.add(
        ParagraphStyle(
            name="Caption",
            parent=styles["BodyText"],
            fontSize=10,
            leading=12,
            alignment=TA_CENTER,
            spaceBefore=6,
            spaceAfter=6,
        )
    )

    story = []

    def draw_header_footer(canvas, doc):
        canvas.saveState()

        # Top maroon bar
        canvas.setFillColor(HEADER_BAR)
        canvas.rect(0, letter[1] - 0.75 * inch, letter[0], 0.5 * inch, fill=1, stroke=0)

        # Header text
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(doc.leftMargin, letter[1] - 0.55 * inch, HEADER_LINE)
        name = (payload.get("professor_name") or "").strip()

        if name:
            if not name.lower().startswith("professor"):
                professor = f"Professor {name}"
            else:
                professor = name
        else:
            professor = "Professor Name"

        canvas.drawRightString(letter[0] - doc.rightMargin, letter[1] - 0.55 * inch, professor)

        # Page counter box (thinner, below header bar)
        box_w = 0.9 * inch
        box_h = 0.20 * inch

        box_x = letter[0] - box_w

        # Place box so its TOP just touches the bottom of the red bar
        red_bar_bottom = letter[1] - 0.75 * inch
        box_y = red_bar_bottom - box_h

        canvas.setFillColor(colors.white)
        canvas.setStrokeColor(colors.black)
        canvas.rect(
            box_x,
            box_y,
            box_w,
            box_h,
            fill=0,
            stroke=1,
        )

        canvas.setFillColor(colors.black)
        canvas.setFont("Helvetica", 8.5)
        canvas.drawCentredString(
            box_x + box_w / 2,
            box_y + box_h / 2 - 0.02 * inch,
            f"{doc.page} | Page",
        )

        # Footer (single logo + department lines)
        y0 = 0.35 * inch
        # Red line (match header color)
        line_y = y0 + 0.32 * inch
        canvas.setStrokeColor(HEADER_BAR)
        canvas.setLineWidth(2)
        canvas.line(doc.leftMargin, line_y, letter[0] - doc.rightMargin, line_y)

        if os.path.exists(LOGO_PATH):
            try:
                logo = ImageReader(LOGO_PATH)
                canvas.drawImage(
                    logo,
                    doc.leftMargin,
                    y0 - 0.05 * inch,
                    width=8.00 * inch,
                    height=0.90 * inch,  # 0.52 * inch
                    preserveAspectRatio=True,
                    mask="auto",
                )
            except Exception:
                pass

        canvas.restoreState()
   # def header(canvas, doc):
       # canvas.setFillColor(colors.white)
        #canvas.setFont("Helvetica-Bold", 10)
        #canvas.drawString(doc.leftMargin, letter[1] - 0.55 * inch, HEADER_LINE)

        #professor = (payload.get("professor_name") or "Professor Name").strip()
        #canvas.drawRightString(letter[0] - doc.rightMargin, letter[1] - 0.55 * inch, professor)


    # =========================
    # Page 1: Cover
    # =========================
    story.append(Spacer(1, 0.70 * inch))
    story.append(Paragraph("<b>LAB 1</b>", styles["CenterTitle"]))
    story.append(Spacer(1, 0.10 * inch))
    story.append(Paragraph("<b>MEASUREMENTS AND<br/>ERROR ANALYSIS</b>", styles["CenterTitle"]))
    story.append(Spacer(1, 1.30 * inch))

    cover_rows = [
        ["Name :", payload.get("member1", "")],
        ["Lab partner name 1 :", payload.get("member2", "")],
        ["Lab partner name 2 :", payload.get("member3", "")],
        ["Date :", payload.get("lab_date", "")],
    ]
    t = Table(cover_rows, colWidths=[2.2 * inch, 4.8 * inch])

    st = yellow_table_style()
    st.add("BACKGROUND", (1, 0), (1, -1), YELLOW)

    # 🔥 Larger font for BOTH columns
    st.add("FONTSIZE", (0, 0), (-1, -1), 13)
    st.add("TOPPADDING", (0, 0), (-1, -1), 8)
    st.add("BOTTOMPADDING", (0, 0), (-1, -1), 8)

    t.setStyle(st)
    story.append(t)

    story.append(PageBreak())

    # =========================
    # Page 2: Objective
    # =========================
    story.append(Paragraph("<b>OBJECTIVE</b>", styles["H1"]))
    story.append(Spacer(1, 0.10 * inch))

    objective_text = (
        "1. To learn how to use the following measuring devices and understand the "
        "uncertainties associated with them.<br/>"
        "a) meter stick<br/>"
        "b) metric ruler<br/>"
        "c) triple-beam balance<br/>"
        "d) digital balance<br/>"
        "e) vernier calipers<br/>"
        "f) micrometer<br/><br/>"
        "2. Use the following general error propagation equation to analyze the errors involved "
        "in making calculations involving measurements with their own uncertainty.<br/><br/>"
        "<b>&sigma;<sub>f</sub> = &radic;((&part;f/&part;x)<sup>2</sup>&sigma;<sub>x</sub><sup>2</sup> + "
        "(&part;f/&part;y)<sup>2</sup>&sigma;<sub>y</sub><sup>2</sup> + "
        "(&part;f/&part;z)<sup>2</sup>&sigma;<sub>z</sub><sup>2</sup>)</b>"
    )
    story.append(Paragraph(objective_text, styles["BodyText"]))
    story.append(PageBreak())

    # =========================
    # THEORY — fixed page-by-page layout
    # =========================

    # Page 3
    story.append(Paragraph("<b>THEORY</b>", styles["H1"]))
    story.append(Paragraph("Refer to lab handout on Error Propagation.", styles["BodyText"]))
    story.append(Spacer(1, 0.10 * inch))

    story.append(Paragraph("<b>Using the Metric Ruler</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.06 * inch))

    metric_text = (
        "Consider the following standard metric ruler.<br/><br/>"
        "The ruler is incremented in units of centimeters (cm). The smallest scale division is "
        "a tenth of a centimeter or 1 mm. Therefore, the uncertainty &Delta;x = smallest increment/2 "
        "= 1 mm/2 = 0.5 mm = 0.05 cm. Note that a measurement made with this ruler must be stated "
        "to a tenth of a centimeter, since the uncertainty is stated to a tenth of a centimeter. "
        "In the example above, the length of the object would be stated as "
        "<b>x = 2.77 cm &plusmn; 0.05 cm.</b>"
    )
    story.append(Paragraph(metric_text, styles["Small"]))
    story.append(Spacer(1, 0.12 * inch))

    if os.path.exists(METRIC_RULER_IMG):
        story.append(fit_image_fixed_height(METRIC_RULER_IMG, target_width=6.6 * inch, target_height=2.35 * inch))
        story.append(caption("Figure 1 — Standard metric ruler.", styles))
    story.append(PageBreak())

    # Page 4
    story.append(Paragraph("<b>Using the Vernier Calipers</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.08 * inch))

    vernier_text = (
        "The Vernier caliper is an instrument that allows you measure lengths much more accurate "
        "than the metric ruler. The smallest increment in the vernier caliper you will be using is "
        "(1/50)mm = 0.02mm = 0.002cm. Thus, the uncertainty is &Delta;x = (1/2)0.002 cm = 0.001 cm."
    )
    story.append(Paragraph(vernier_text, styles["Small"]))
    story.append(Spacer(1, 0.12 * inch))

    if os.path.exists(VERNIER_IMG_PAGE4):
        story.append(fit_image_fixed_height(VERNIER_IMG_PAGE4, target_width=6.6 * inch, target_height=4.65 * inch))
        story.append(caption("Figure 2 — Vernier calipers (jaws and main scale).", styles))
    story.append(PageBreak())

    # Page 5
    vernier_example_text = (
        "Note that the zero line on the vernier scale falls between the 4.4 cm and 4.5 cm mark on the "
        "main scale. Thus, the first significant digits are 4.4 cm. The remaining two digits are obtained "
        "by noting which line on the vernier scale (0,2,4,6,8) coincides best with a line on the main scale. "
        "Looking closely at the picture below indicates that the 46 line lines up the closest. Therefore, the "
        "reading is 4.446 cm. Or in standard form 4.446 cm &plusmn; 0.001 cm"
    )
    story.append(Paragraph(vernier_example_text, styles["Small"]))
    story.append(Spacer(1, 0.12 * inch))

    if os.path.exists(VERNIER_IMG_PAGE5):
        story.append(fit_image_fixed_height(VERNIER_IMG_PAGE5, target_width=6.6 * inch, target_height=5.30 * inch))
        story.append(caption("Figure 3 — Vernier caliper example reading (4.446 cm ± 0.001 cm).", styles))
    story.append(PageBreak())

    # Page 6
    story.append(Paragraph("<b>Using The Micrometer Caliper</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.08 * inch))

    mic_text = textwrap.dedent(
        """The micrometer caliper has a linear scale engraved on its sleeve and a circular scale engraved on the thimble. The linear (sleeve) scale is divided into 1 mm divisions and is 25 mm long. Half-millimeter (0.5 mm) marks are provided below the main scale.

         The circular (thimble) scale has 50 divisions. One complete revolution of the thimble advances it by 0.5 mm along the linear scale. Therefore, each division on the thimble corresponds to 0.01 mm.

         In Figure 5, the main scale is marked with 0 and 5, which indicate millimeters. The marks below the main scale indicate half-millimeter increments, since one full rotation of the
          thimble advances the spindle by 0.5 mm.

         In the example shown in Figure 5, the half-millimeter mark to the right of the sixth main-scale mark is visible. Thus, the reading is between 6.5 mm and 7.0 mm. The thimble scale aligns near the 41st division, corresponding to 0.41 mm. By estimating one additional digit, the reading can be refined by 0.002 mm.

         Therefore, the micrometer reading is 6.5 mm + 0.41 mm + 0.002 mm = 6.912 mm. Converting to meters, 6.912 mm = 6.912 × 10<sup>-3</sup> m. This demonstrates that micrometer measurements can be estimated to the nearest thousandth of a millimeter (0.001 mm)."""
    )
    story.append(Paragraph(mic_text, styles["Small"]))
    story.append(Spacer(1, 0.08 * inch))

    if os.path.exists(MICROMETER_IMG_1):
        story.append(fit_image_fixed_height(MICROMETER_IMG_1, target_width=6.6 * inch, target_height=2.55 * inch))
        story.append(caption("Figure 4 — Micrometer example reading (Example 1).", styles))
        story.append(Spacer(1, 0.08 * inch))

    if os.path.exists(MICROMETER_IMG_2):
        story.append(fit_image_fixed_height(MICROMETER_IMG_2, target_width=6.6 * inch, target_height=2.55 * inch))
        story.append(caption("Figure 5 — Micrometer example reading (Example 2).", styles))

    story.append(PageBreak())

    # =========================
    # Student report fields
    # =========================
    story.append(Paragraph("<b>EQUIPMENT</b>", styles["H1"]))
    equip = (
        "1. one aluminum block<br/>"
        "2. meter stick<br/>"
        "3. metric ruler<br/>"
        "4. triple-beam balance<br/>"
        "5. digital balance<br/>"
        "6. vernier calipers<br/>"
        "7. micrometer"
    )
    story.append(Paragraph(equip, styles["BodyText"]))
    story.append(Spacer(1, 0.12 * inch))

    story.append(Paragraph("<b>PROCEDURE</b>", styles["H1"]))
    story.append(
        Paragraph(
            "(for this lab any measurements and calculations should be stated in the standard form of:<br/>"
            "measurement = x<sub>best</sub> &plusmn; &Delta;x)",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 0.12 * inch))

    # Part I
    story.append(Paragraph("<b>Part I (Using Measuring Devices)</b>", styles["Heading2"]))
    story.append(
        Paragraph(
            "1. Learn to use all the measuring devices listed above.<br/>"
            "2. Calculate the uncertainties of all measuring devices you will be using.",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 0.08 * inch))
    story.append(Paragraph("<b>Results :</b>", styles["BodyText"]))

    part1_rows = [
        ["", "Measuring Device", "Uncertainty"],
        ["a", "meter stick", payload.get("unc_meterstick", "")],
        ["b", "metric ruler", payload.get("unc_ruler", "")],
        ["c", "triple-beam balance", payload.get("unc_triple", "")],
        ["d", "digital balance", payload.get("unc_digital", "")],
        ["e", "vernier calipers", payload.get("unc_vernier", "")],
        ["f", "micrometer", payload.get("unc_micrometer", "")],
    ]
    t = Table(part1_rows, colWidths=[0.35 * inch, 3.7 * inch, 2.95 * inch])
    st = yellow_table_style()
    st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
    st.add("BACKGROUND", (2, 1), (2, -1), YELLOW)
    t.setStyle(st)
    story.append(t)

    story.append(PageBreak())

    # Part II
    story.append(Paragraph("<b>Part II (Volume of Table-Top)</b>", styles["Heading2"]))
    story.append(
        Paragraph(
            "1. Measure the dimensions of your table-top with the meter stick.<br/>"
            "2. Using the error propagation equation derive an expression for the uncertainty σV for "
            "the volume of the table top.<br/>"
            "3. Calculate the volume of the table top.<br/>"
            "4. Calculate the uncertainty σV of the table top.",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 0.08 * inch))
    story.append(Paragraph("<b>Results :</b>", styles["BodyText"]))

    t = Table(
        [["Measuring tool", payload.get("p2_tool", "")], ["Object", payload.get("p2_object", "")]],
        colWidths=[2.5 * inch, 4.5 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 0), (1, 1), YELLOW)
    t.setStyle(st)
    story.append(t)
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [["Tabletop", "Length", "Width", "Thickness"], ["Dimensions (cm)", payload.get("p2_L", ""), payload.get("p2_W", ""), payload.get("p2_T", "")]],
        colWidths=[1.7 * inch, 1.75 * inch, 1.75 * inch, 1.75 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 1), (3, 1), YELLOW)
    t.setStyle(st)
    story.append(t)
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [["expression for the uncertainty  σV in\nthe volume of the table top", payload.get("p2_expr", "")]],
        colWidths=[3.2 * inch, 3.8 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
    t.setStyle(st)
    story.append(t)
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [["", "Volume", "Units"], ["Tabletop volume", payload.get("p2_volume", ""), payload.get("p2_vol_units", "")]],
        colWidths=[1.9 * inch, 3.0 * inch, 2.1 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 1), (2, 1), YELLOW)
    t.setStyle(st)
    story.append(t)
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [["Calculate the uncertainty  σV in the\nvolume of the table top", "Uploaded file: " + (os.path.basename(payload.get("p2_unc_upload", "")) if payload.get("p2_unc_upload") else "")]],
        colWidths=[3.2 * inch, 3.8 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
    t.setStyle(st)
    story.append(t)

    # Embed jpg/png inline + note for pdf
    add_uploaded_block(story, styles, "Tabletop uncertainty work", payload.get("p2_unc_upload", ""))

    t = Table(
        [["", "Volume", "", "Error", "Units"], ["Tabletop volume", payload.get("p2_vol_best", ""), "±", payload.get("p2_vol_err", ""), payload.get("p2_vol_err_units", "")]],
        colWidths=[1.9 * inch, 1.7 * inch, 0.35 * inch, 1.7 * inch, 1.35 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
    st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
    st.add("BACKGROUND", (3, 1), (4, 1), YELLOW)
    t.setStyle(st)
    story.append(t)

    def dims_block(title, tool_label, obj_key, L_key, W_key, T_key, expr_key, vol_key, volu_key, up_key, best_key, err_key, erru_key):
        story.append(PageBreak())
        story.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))
        story.append(Spacer(1, 0.06 * inch))

        t = Table([["Measuring tool", tool_label], ["Object", payload.get(obj_key, "")]], colWidths=[2.5 * inch, 4.5 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["Aluminum block", "Length", "Width", "Thickness"], ["Dimensions (cm)", payload.get(L_key, ""), payload.get(W_key, ""), payload.get(T_key, "")]],
                  colWidths=[1.8 * inch, 1.7 * inch, 1.7 * inch, 1.8 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 1), (3, 1), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["expression for the uncertainty  σV in\nthe volume of the aluminum block", payload.get(expr_key, "")]],
                  colWidths=[3.2 * inch, 3.8 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["", "Volume", "Units"], ["Aluminum block volume", payload.get(vol_key, ""), payload.get(volu_key, "")]],
                  colWidths=[1.9 * inch, 3.0 * inch, 2.1 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 1), (2, 1), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["Calculate the uncertainty  σV in the\nvolume of the aluminum block",
                    "Uploaded file: " + (os.path.basename(payload.get(up_key, "")) if payload.get(up_key) else "")]],
                  colWidths=[3.2 * inch, 3.8 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
        t.setStyle(st)
        story.append(t)

        add_uploaded_block(story, styles, "Uncertainty work (volume)", payload.get(up_key, ""))

        t = Table([["", "Volume", "", "Error", "Units"],
                   ["Aluminum block volume", payload.get(best_key, ""), "±", payload.get(err_key, ""), payload.get(erru_key, "")]],
                  colWidths=[1.9 * inch, 1.7 * inch, 0.35 * inch, 1.7 * inch, 1.35 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
        st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
        st.add("BACKGROUND", (3, 1), (4, 1), YELLOW)
        t.setStyle(st)
        story.append(t)

    dims_block(
        "Part III (Density of Aluminum Block) — Dimensions Using Metric Ruler",
        "Metric Ruler",
        "p3_obj_ruler",
        "p3_ruler_L",
        "p3_ruler_W",
        "p3_ruler_T",
        "p3_ruler_expr",
        "p3_ruler_vol",
        "p3_ruler_vol_units",
        "p3_ruler_unc_upload",
        "p3_ruler_best",
        "p3_ruler_err",
        "p3_ruler_err_units",
    )

    dims_block(
        "Part III (Density of Aluminum Block) — Dimensions Using Vernier Caliper",
        "Vernier Caliper",
        "p3_obj_vernier",
        "p3_vernier_L",
        "p3_vernier_W",
        "p3_vernier_T",
        "p3_vernier_expr",
        "p3_vernier_vol",
        "p3_vernier_vol_units",
        "p3_vernier_unc_upload",
        "p3_vernier_best",
        "p3_vernier_err",
        "p3_vernier_err_units",
    )

    dims_block(
        "Part III (Density of Aluminum Block) — Dimensions Using Micrometer",
        "Micrometer",
        "p3_obj_mic",
        "p3_mic_L",
        "p3_mic_W",
        "p3_mic_T",
        "p3_mic_expr",
        "p3_mic_vol",
        "p3_mic_vol_units",
        "p3_mic_unc_upload",
        "p3_mic_best",
        "p3_mic_err",
        "p3_mic_err_units",
    )

    # Mass page
    story.append(PageBreak())
    story.append(Paragraph("<b>Mass of Aluminum Block</b>", styles["Heading2"]))

    def mass_block(title, tool_label, obj_key, m_key, err_key, u_key):
        story.append(Spacer(1, 0.10 * inch))
        story.append(Paragraph(f"<b>{title}</b>", styles["BodyText"]))
        t = Table([["Measuring tool", tool_label], ["Object", payload.get(obj_key, "")]], colWidths=[2.5 * inch, 4.5 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["", "Mass", "", "Error", "Units"], ["Aluminum Block", payload.get(m_key, ""), "±", payload.get(err_key, ""), payload.get(u_key, "")]],
                  colWidths=[1.9 * inch, 1.7 * inch, 0.35 * inch, 1.7 * inch, 1.35 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
        st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
        st.add("BACKGROUND", (3, 1), (4, 1), YELLOW)
        t.setStyle(st)
        story.append(t)

    mass_block("Mass Using Triple-beam Balance", "Triple-beam balance", "m_tb_obj", "m_tb_mass", "m_tb_err", "m_tb_units")
    mass_block("Mass Using Digital Balance", "Digital balance", "m_dig_obj", "m_dig_mass", "m_dig_err", "m_dig_units")

    def density_block(title, expr_key, dens_key, densu_key, up_key, best_key, err_key, erru_key):
        story.append(PageBreak())
        story.append(Paragraph(f"<b>{title}</b>", styles["Heading2"]))

        t = Table([["expression for the uncertainty  σρ in\nthe density", payload.get(expr_key, "")]], colWidths=[3.2 * inch, 3.8 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["", "Density", "Units"], ["Aluminum Block Density", payload.get(dens_key, ""), payload.get(densu_key, "")]],
                  colWidths=[1.9 * inch, 3.0 * inch, 2.1 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 1), (2, 1), YELLOW)
        t.setStyle(st)
        story.append(t)
        story.append(Spacer(1, 0.08 * inch))

        t = Table([["Calculate the uncertainty  σρ in the\ndensity of the aluminum block",
                    "Uploaded file: " + (os.path.basename(payload.get(up_key, "")) if payload.get(up_key) else "")]],
                  colWidths=[3.2 * inch, 3.8 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
        t.setStyle(st)
        story.append(t)

        add_uploaded_block(story, styles, "Uncertainty work (density)", payload.get(up_key, ""))

        t = Table([["", "Density", "", "Error", "Units"],
                   ["Aluminum Block Density", payload.get(best_key, ""), "±", payload.get(err_key, ""), payload.get(erru_key, "")]],
                  colWidths=[1.9 * inch, 1.7 * inch, 0.35 * inch, 1.7 * inch, 1.35 * inch])
        st = yellow_table_style()
        st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
        st.add("BACKGROUND", (1, 1), (1, 1), YELLOW)
        st.add("BACKGROUND", (3, 1), (4, 1), YELLOW)
        t.setStyle(st)
        story.append(t)

    density_block("Density Estimate - Triple-beam balance and the Metric ruler.", "d1_expr", "d1_density", "d1_units", "d1_upload", "d1_best", "d1_err", "d1_err_units")
    density_block("Density Estimate - Digital balance and the Vernier calipers.", "d2_expr", "d2_density", "d2_units", "d2_upload", "d2_best", "d2_err", "d2_err_units")
    density_block("Density Estimate - Digital balance and the Micrometer.", "d3_expr", "d3_density", "d3_units", "d3_upload", "d3_best", "d3_err", "d3_err_units")

    # Sample calc upload
    story.append(Spacer(1, 0.12 * inch))
    t = Table([["Sample Calculation :", "Uploaded file: " + (os.path.basename(payload.get("sample_calc", "")) if payload.get("sample_calc") else "")]],
              colWidths=[2.2 * inch, 4.8 * inch])
    st = yellow_table_style()
    st.add("BACKGROUND", (0, 0), (1, 0), YELLOW)
    t.setStyle(st)
    story.append(t)
    add_uploaded_block(story, styles, "Sample calculation", payload.get("sample_calc", ""))

    # Data analysis
    story.append(PageBreak())
    story.append(Paragraph("<b>Data Analysis</b>", styles["H1"]))
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [["% error table", "Triple beam\nbalance and the\nMetric Ruler", "Digital balance\nand the Vernier\nCaliper", "Digital balance\nand the\nMicrometer"],
         ["% error", payload.get("perr_1", ""), payload.get("perr_2", ""), payload.get("perr_3", "")]],
        colWidths=[1.5 * inch, 1.85 * inch, 1.85 * inch, 1.8 * inch],
    )
    st = yellow_table_style()
    st.add("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREY)
    st.add("BACKGROUND", (1, 1), (3, 1), YELLOW)
    t.setStyle(st)
    story.append(t)

    story.append(Spacer(1, 0.08 * inch))
    t = Table([["Sample Calculation :", "Uploaded file: " + (os.path.basename(payload.get("perr_upload", "")) if payload.get("perr_upload") else "")]],
              colWidths=[2.2 * inch, 4.8 * inch])
    st = yellow_table_style()
    st.add("BACKGROUND", (0, 0), (1, 0), YELLOW)
    t.setStyle(st)
    story.append(t)
    add_uploaded_block(story, styles, "Percent error sample calculation", payload.get("perr_upload", ""))

    qa = [
        ("2. Which of the 3 densities gave the most accurate answer? Was this what you expected why or why not? Explain!", "qa2"),
        ("3. Was the propagating error involved in calculating the density significant with any combination of the measuring devices? Explain.", "qa3"),
        ("4. What were the random errors involved and how did they affect the density and uncertainty calculation?", "qa4"),
        ("5. What systematic errors were involved?", "qa5"),
        ("6. Comment on any other sources of error that could have been involved.", "qa6"),
    ]

    def p(text: str, style):
        return Paragraph((text or "").replace("\n", "<br/>"), style)

    for prompt, key in qa:
        story.append(Spacer(1, 0.10 * inch))
        story.append(Paragraph(prompt, styles["BodyText"]))

        t = Table(
            [[p("Answer:", styles["BodyText"]), p(payload.get(key, ""), styles["BodyText"])]],
            colWidths=[1.0 * inch, 6.0 * inch],
        )
        st = yellow_table_style()
        st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
        t.setStyle(st)
        story.append(t)


    # =========================
    # Appendix I – Images of Measurements
    # =========================
    story.append(PageBreak())
    story.append(Paragraph("<b>APPENDIX I – Images of Measurements</b>", styles["H1"]))
    story.append(Spacer(1, 0.12 * inch))

    # ---- Part II: Volume of Tabletop (Meter Ruler)
    story.append(Paragraph("<b>Part II – Volume of Tabletop – Meter Ruler</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.08 * inch))

    t = Table(
        [[
            appendix_image_cell(payload.get("app1_table_length")),
            appendix_image_cell(payload.get("app1_table_width")),
            appendix_image_cell(payload.get("app1_table_height")),
        ]],
        colWidths=[APP1_IMG_W] * 3,
        rowHeights=[APP1_IMG_H],
    )
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
    ]))
    story.append(t)

    story.append(Spacer(1, 0.25 * inch))

    # ---- Part III: Density of Aluminum Block
    story.append(Paragraph("<b>Part III – Density of Aluminum Block</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.08 * inch))

    header = ["Metric Ruler", "Vernier Caliper", "Micrometer"]
    rows = ["LENGTH", "WIDTH", "HEIGHT"]

    data = []
    data.append(header)

    for r in rows:
        data.append([
            appendix_image_cell(payload.get(f"app1_{r.lower()}_ruler")),
            appendix_image_cell(payload.get(f"app1_{r.lower()}_vernier")),
            appendix_image_cell(payload.get(f"app1_{r.lower()}_micrometer")),
        ])

    t = Table(
        data,
        colWidths=[APP1_IMG_W] * 3,
        rowHeights=[0.4 * inch] + [APP1_IMG_H] * 3,
    )
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1.2, colors.black),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(t)

    story.append(Spacer(1, 0.18 * inch))
    story.append(Paragraph("<b>Mass Measurements</b>", styles["BodyText"]))
    story.append(Spacer(1, 0.06 * inch))

    mass_tbl = Table(
        [[
            Paragraph("<b>DIGITAL BALANCE</b>", styles["BodyText"]),
            Paragraph("<b>TRIPLE BEAM BALANCE</b>", styles["BodyText"]),
        ],
            [
                appendix_image_cell(payload.get("app1_mass_digital", "")),
                appendix_image_cell(payload.get("app1_mass_triplebeam", "")),
            ]],
        colWidths=[3.5 * inch, 3.5 * inch],
        rowHeights=[APP1_MASS_IMG_H, APP1_IMG_H],
    )

    mass_tbl.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 1.0, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(mass_tbl)

    # Appendix II signed data
    story.append(PageBreak())
    story.append(Paragraph("<b>Appendix II</b>", styles["H1"]))
    story.append(Paragraph("Instructor signed data from the experiment", styles["BodyText"]))
    story.append(Spacer(1, 0.08 * inch))

    signed = payload.get("signed_data", "")
    t = Table([["Signed data from the lab", "Uploaded file: " + (os.path.basename(signed) if signed else "")]], colWidths=[3.2 * inch, 3.8 * inch])
    st = yellow_table_style()
    st.add("BACKGROUND", (1, 0), (1, 0), YELLOW)
    t.setStyle(st)
    story.append(t)

    add_uploaded_block(story, styles, "Signed data", signed, max_h=6.6 * inch)



    # Appendix III (optional)
    if include_appendix_ii:
        story.append(PageBreak())
        story.append(Paragraph("<b>Appendix III</b>", styles["H1"]))
        story.append(Paragraph("Instructor-only reference section (locked).", styles["Small"]))

    # Encryption / lock
    if include_appendix_ii and instructor_password.strip():
        encrypt = StandardEncryption(
            userPassword=instructor_password.strip(),
            ownerPassword=instructor_password.strip() + "_OWNER",
            canPrint=1,
            canModify=0,
            canCopy=0,
            canAnnotate=0,
        )
    else:
        encrypt = StandardEncryption(
            userPassword="",
            ownerPassword="OWNER_LOCK",
            canPrint=1,
            canModify=0,
            canCopy=0,
            canAnnotate=0,
        )

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=1.0 * inch,
        bottomMargin=1.35 * inch,
        encrypt=encrypt,
        title="PHYS 4A Lab 1 - Measurements and Error Analysis",
    )
    doc.build(story, onFirstPage=draw_header_footer, onLaterPages=draw_header_footer)
    return buf.getvalue()


@app.get("/")
def home():
    return render_template("wk2_lab_HTMLform.html")


@app.post("/generate")
def generate():
    try:
        payload = {
            "member1": (request.form.get("member1") or "").strip(),
            "member2": (request.form.get("member2") or "").strip(),
            "member3": (request.form.get("member3") or "").strip(),
            "lab_date": (request.form.get("lab_date") or "").strip(),
            "unc_meterstick": (request.form.get("unc_meterstick") or "").strip(),
            "unc_ruler": (request.form.get("unc_ruler") or "").strip(),
            "unc_triple": (request.form.get("unc_triple") or "").strip(),
            "unc_digital": (request.form.get("unc_digital") or "").strip(),
            "unc_vernier": (request.form.get("unc_vernier") or "").strip(),
            "unc_micrometer": (request.form.get("unc_micrometer") or "").strip(),
            "p2_tool": (request.form.get("p2_tool") or "").strip(),
            "p2_object": (request.form.get("p2_object") or "").strip(),
            "p2_L": (request.form.get("p2_L") or "").strip(),
            "p2_W": (request.form.get("p2_W") or "").strip(),
            "p2_T": (request.form.get("p2_T") or "").strip(),
            "p2_expr": (request.form.get("p2_expr") or "").strip().replace("\n", "<br/>"),
            "p2_volume": (request.form.get("p2_volume") or "").strip(),
            "p2_vol_units": (request.form.get("p2_vol_units") or "").strip(),
            "p2_vol_best": (request.form.get("p2_vol_best") or "").strip(),
            "p2_vol_err": (request.form.get("p2_vol_err") or "").strip(),
            "p2_vol_err_units": (request.form.get("p2_vol_err_units") or "").strip(),
            "p3_obj_ruler": (request.form.get("p3_obj_ruler") or "").strip(),
            "p3_ruler_L": (request.form.get("p3_ruler_L") or "").strip(),
            "p3_ruler_W": (request.form.get("p3_ruler_W") or "").strip(),
            "p3_ruler_T": (request.form.get("p3_ruler_T") or "").strip(),
            "p3_ruler_expr": (request.form.get("p3_ruler_expr") or "").strip().replace("\n", "<br/>"),
            "p3_ruler_vol": (request.form.get("p3_ruler_vol") or "").strip(),
            "p3_ruler_vol_units": (request.form.get("p3_ruler_vol_units") or "").strip(),
            "p3_ruler_best": (request.form.get("p3_ruler_best") or "").strip(),
            "p3_ruler_err": (request.form.get("p3_ruler_err") or "").strip(),
            "p3_ruler_err_units": (request.form.get("p3_ruler_err_units") or "").strip(),
            "p3_obj_vernier": (request.form.get("p3_obj_vernier") or "").strip(),
            "p3_vernier_L": (request.form.get("p3_vernier_L") or "").strip(),
            "p3_vernier_W": (request.form.get("p3_vernier_W") or "").strip(),
            "p3_vernier_T": (request.form.get("p3_vernier_T") or "").strip(),
            "p3_vernier_expr": (request.form.get("p3_vernier_expr") or "").strip().replace("\n", "<br/>"),
            "p3_vernier_vol": (request.form.get("p3_vernier_vol") or "").strip(),
            "p3_vernier_vol_units": (request.form.get("p3_vernier_vol_units") or "").strip(),
            "p3_vernier_best": (request.form.get("p3_vernier_best") or "").strip(),
            "p3_vernier_err": (request.form.get("p3_vernier_err") or "").strip(),
            "p3_vernier_err_units": (request.form.get("p3_vernier_err_units") or "").strip(),
            "p3_obj_mic": (request.form.get("p3_obj_mic") or "").strip(),
            "p3_mic_L": (request.form.get("p3_mic_L") or "").strip(),
            "p3_mic_W": (request.form.get("p3_mic_W") or "").strip(),
            "p3_mic_T": (request.form.get("p3_mic_T") or "").strip(),
            "p3_mic_expr": (request.form.get("p3_mic_expr") or "").strip().replace("\n", "<br/>"),
            "p3_mic_vol": (request.form.get("p3_mic_vol") or "").strip(),
            "p3_mic_vol_units": (request.form.get("p3_mic_vol_units") or "").strip(),
            "p3_mic_best": (request.form.get("p3_mic_best") or "").strip(),
            "p3_mic_err": (request.form.get("p3_mic_err") or "").strip(),
            "p3_mic_err_units": (request.form.get("p3_mic_err_units") or "").strip(),
            "m_tb_obj": (request.form.get("m_tb_obj") or "").strip(),
            "m_tb_mass": (request.form.get("m_tb_mass") or "").strip(),
            "m_tb_err": (request.form.get("m_tb_err") or "").strip(),
            "m_tb_units": (request.form.get("m_tb_units") or "").strip(),
            "m_dig_obj": (request.form.get("m_dig_obj") or "").strip(),
            "m_dig_mass": (request.form.get("m_dig_mass") or "").strip(),
            "m_dig_err": (request.form.get("m_dig_err") or "").strip(),
            "m_dig_units": (request.form.get("m_dig_units") or "").strip(),
            "d1_expr": (request.form.get("d1_expr") or "").strip().replace("\n", "<br/>"),
            "d1_density": (request.form.get("d1_density") or "").strip(),
            "d1_units": (request.form.get("d1_units") or "").strip(),
            "d1_best": (request.form.get("d1_best") or "").strip(),
            "d1_err": (request.form.get("d1_err") or "").strip(),
            "d1_err_units": (request.form.get("d1_err_units") or "").strip(),
            "d2_expr": (request.form.get("d2_expr") or "").strip().replace("\n", "<br/>"),
            "d2_density": (request.form.get("d2_density") or "").strip(),
            "d2_units": (request.form.get("d2_units") or "").strip(),
            "d2_best": (request.form.get("d2_best") or "").strip(),
            "d2_err": (request.form.get("d2_err") or "").strip(),
            "d2_err_units": (request.form.get("d2_err_units") or "").strip(),
            "d3_expr": (request.form.get("d3_expr") or "").strip().replace("\n", "<br/>"),
            "d3_density": (request.form.get("d3_density") or "").strip(),
            "d3_units": (request.form.get("d3_units") or "").strip(),
            "d3_best": (request.form.get("d3_best") or "").strip(),
            "d3_err": (request.form.get("d3_err") or "").strip(),
            "d3_err_units": (request.form.get("d3_err_units") or "").strip(),
            "perr_1": (request.form.get("perr_1") or "").strip(),
            "perr_2": (request.form.get("perr_2") or "").strip(),
            "perr_3": (request.form.get("perr_3") or "").strip(),
            "qa2": (request.form.get("qa2") or "").strip(),
            "qa3": (request.form.get("qa3") or "").strip(),
            "qa4": (request.form.get("qa4") or "").strip(),
            "qa5": (request.form.get("qa5") or "").strip(),
            "qa6": (request.form.get("qa6") or "").strip(),
            "professor_name": (request.form.get("professor_name") or "").strip(),

        }

        # Uploads
        payload["p2_unc_upload"] = save_upload(request.files.get("p2_unc_upload"), "p2_unc")
        payload["p3_ruler_unc_upload"] = save_upload(request.files.get("p3_ruler_unc_upload"), "p3_ruler_unc")
        payload["p3_vernier_unc_upload"] = save_upload(request.files.get("p3_vernier_unc_upload"), "p3_vernier_unc")
        payload["p3_mic_unc_upload"] = save_upload(request.files.get("p3_mic_unc_upload"), "p3_mic_unc")

        payload["d1_upload"] = save_upload(request.files.get("d1_upload"), "d1")
        payload["d2_upload"] = save_upload(request.files.get("d2_upload"), "d2")
        payload["d3_upload"] = save_upload(request.files.get("d3_upload"), "d3")

        payload["sample_calc"] = save_upload(request.files.get("sample_calc"), "sample_calc")
        payload["perr_upload"] = save_upload(request.files.get("perr_upload"), "perr")
        payload["signed_data"] = save_upload(request.files.get("signed_data"), "signed")

        # Appendix I image uploads (NEW)
        payload["app1_table_length"] = save_upload(request.files.get("app1_table_length"), "app1_table_length")
        payload["app1_table_width"] = save_upload(request.files.get("app1_table_width"), "app1_table_width")
        payload["app1_table_height"] = save_upload(request.files.get("app1_table_height"), "app1_table_height")

        payload["app1_length_ruler"] = save_upload(request.files.get("app1_length_ruler"), "app1_length_ruler")
        payload["app1_length_vernier"] = save_upload(request.files.get("app1_length_vernier"), "app1_length_vernier")
        payload["app1_length_micrometer"] = save_upload(request.files.get("app1_length_micrometer"),
                                                        "app1_length_micrometer")

        payload["app1_width_ruler"] = save_upload(request.files.get("app1_width_ruler"), "app1_width_ruler")
        payload["app1_width_vernier"] = save_upload(request.files.get("app1_width_vernier"), "app1_width_vernier")
        payload["app1_width_micrometer"] = save_upload(request.files.get("app1_width_micrometer"),
                                                       "app1_width_micrometer")

        payload["app1_height_ruler"] = save_upload(request.files.get("app1_height_ruler"), "app1_height_ruler")
        payload["app1_height_vernier"] = save_upload(request.files.get("app1_height_vernier"), "app1_height_vernier")
        payload["app1_height_micrometer"] = save_upload(request.files.get("app1_height_micrometer"),
                                                        "app1_height_micrometer")

        # Mass images (NEW)
        payload["app1_mass_digital"] = save_upload(request.files.get("app1_mass_digital"), "app1_mass_digital")
        payload["app1_mass_triplebeam"] = save_upload(request.files.get("app1_mass_triplebeam"), "app1_mass_triplebeam")

        instructor_pw = (request.form.get("instructor_password") or "").strip()
        include_appendix_ii = bool(request.form.get("include_appendix_ii")) and bool(instructor_pw)

        if not payload["member1"]:
            raise ValueError("Lab member name 1 is required.")
        if not payload["lab_date"]:
            raise ValueError("Date is required.")

        # Build student PDF
        student_pdf = build_pdf_wk2_lab(payload, include_appendix_ii=False)

        # Append any uploaded PDFs to the end
        all_uploads = [
            payload.get("p2_unc_upload", ""),
            payload.get("p3_ruler_unc_upload", ""),
            payload.get("p3_vernier_unc_upload", ""),
            payload.get("p3_mic_unc_upload", ""),
            payload.get("d1_upload", ""),
            payload.get("d2_upload", ""),
            payload.get("d3_upload", ""),
            payload.get("sample_calc", ""),
            payload.get("perr_upload", ""),
            payload.get("signed_data", ""),
        ]
        student_pdf = append_pdf_uploads(student_pdf, all_uploads)

        base = safe_filename(payload["member1"] + "_" + payload["lab_date"]).replace(".", "_")
        student_name = f"{base}_STUDENT.pdf"
        student_path = os.path.join(OUT_DIR, student_name)
        with open(student_path, "wb") as f:
            f.write(student_pdf)

        # Optional instructor copy (locked)
        if include_appendix_ii:
            instr_pdf = build_pdf(payload, include_appendix_ii=True, instructor_password=instructor_pw)
            instr_pdf = append_pdf_uploads(instr_pdf, all_uploads)
            instr_name = f"{base}_INSTRUCTOR.pdf"
            with open(os.path.join(OUT_DIR, instr_name), "wb") as f:
                f.write(instr_pdf)

        return send_file(
            io.BytesIO(student_pdf),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=student_name,
        )

    except Exception as e:
        abort(400, f"Error: {e}")

if __name__ == "__main__":
    app.run(debug=True)
