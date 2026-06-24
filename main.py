from __future__ import annotations

from io import BytesIO
from pathlib import Path
from typing import List

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pydantic import BaseModel, Field
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas


BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "assets" / "logo.png"

COMPANY_NAME = "SHREE GANESH SOLAR ENERGY"
OFFICE_ADDRESS = "P/N 1764/29 HANUMAN NAGAR BHANDEWADI PARDI NAGPUR 440008"
EMAIL = "rakeshbhude69@gmail.com"
GST_NO = "27DSQPB6025M1ZC"
GST_FOOTER = "27DSQPB6025M"
CONTACT_NO = "+91 7385284731"
BANK_NAME = "INDIAN BANK"
ACCOUNT_HOLDER = "SHREE GANESH SOLAR ENERGY"
CC_ACCOUNT_NO = "8275515270"
IFSC_CODE = "IDIB000R558"
BRANCH = "DESHPANDE LAYOUT, NAGPUR"

PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 36
GREEN = colors.HexColor("#176B3A")
DARK = colors.HexColor("#172033")
MUTED = colors.HexColor("#536170")
GOLD = colors.HexColor("#E0A623")
LINE = colors.HexColor("#C8D3DC")
SOFT_GREEN = colors.HexColor("#EAF4ED")
SOFT_GOLD = colors.HexColor("#FFF4D4")


class Specification(BaseModel):
    item: str = ""
    brand_parameters: str = ""
    detail: str = ""
    specification: str = ""
    quantity: str = ""


class PriceRow(BaseModel):
    sr_no: str = ""
    system_detail: str = ""
    rate: str = ""


class Quotation(BaseModel):
    id: str = ""
    quotation_number: str
    customer_name: str
    location: str
    capacity_kw: str
    quote_date: str
    final_price: str = ""
    panel_brand: str = ""
    inverter_brand: str = ""
    introduction: str = ""
    specifications: List[Specification] = Field(default_factory=list)
    price_header: str = ""
    price_rows: List[PriceRow] = Field(default_factory=list)
    price_note: str = ""
    payment_note: str = ""
    terms_conditions: str = ""
    remark: str = ""


app = FastAPI(title=f"{COMPANY_NAME} - Solar Quotation Generator")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

saved_quotations: list[Quotation] = []


def clean_text(value: object = "") -> str:
    text = str(value or "")
    replacements = {
        "₹": "Rs. ",
        "–": "-",
        "—": "-",
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "\u00a0": " ",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return text.encode("latin-1", "replace").decode("latin-1")


def capacity_kw(quotation: Quotation) -> str:
    value = str(quotation.capacity_kw or "").strip()
    if not value:
        return "0KW"
    return value.upper() if value.upper().endswith("KW") else f"{value}KW"


def rupees(value: str) -> str:
    raw = str(value or "0").strip().replace(",", "").replace("Rs.", "").replace("Rs", "")
    try:
        return f"Rs. {float(raw):,.0f}"
    except ValueError:
        return f"Rs. {value}" if value else "Rs. 0"


def final_price_for_row(value: str) -> str:
    return f"{rupees(value)} /-"


def export_name(quotation: Quotation, extension: str) -> str:
    return f"{quotation.quotation_number}.{extension}"


def quotation_dump(quotation: Quotation) -> dict:
    return quotation.model_dump() if hasattr(quotation, "model_dump") else quotation.dict()


def paragraph_blocks(text: str) -> list[str]:
    return [block.strip() for block in clean_text(text).splitlines() if block.strip()]


def wrap_line(text: str, width: float, font_name: str, font_size: float) -> list[str]:
    words = clean_text(text).split()
    if not words:
        return [""]
    lines: list[str] = []
    current = ""
    for word in words:
        trial = f"{current} {word}".strip()
        if pdfmetrics.stringWidth(trial, font_name, font_size) <= width or not current:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def draw_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    size: float = 10,
    font: str = "Helvetica",
    fill=colors.black,
) -> None:
    c.setFillColor(fill)
    c.setFont(font, size)
    c.drawString(x, y, clean_text(text))


def draw_centered(
    c: canvas.Canvas,
    text: str,
    y: float,
    size: float = 12,
    font: str = "Helvetica-Bold",
    fill=colors.black,
) -> None:
    c.setFillColor(fill)
    c.setFont(font, size)
    c.drawCentredString(PAGE_WIDTH / 2, y, clean_text(text))


def draw_right(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    size: float = 10,
    font: str = "Helvetica",
    fill=colors.black,
) -> None:
    c.setFillColor(fill)
    c.setFont(font, size)
    c.drawRightString(x, y, clean_text(text))


def draw_wrapped(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    width: float,
    size: float = 10,
    leading: float = 13,
    font: str = "Helvetica",
    fill=colors.black,
) -> float:
    for line in wrap_line(text, width, font, size):
        draw_text(c, line, x, y, size, font, fill)
        y -= leading
    return y


def draw_logo(c: canvas.Canvas, x: float, y: float, width: float, height: float) -> None:
    c.setStrokeColor(LINE)
    c.setFillColor(colors.white)
    c.roundRect(x, y, width, height, 5, stroke=1, fill=1)
    if LOGO_PATH.exists():
        c.drawImage(
            str(LOGO_PATH),
            x + 6,
            y + 6,
            width=width - 12,
            height=height - 12,
            preserveAspectRatio=True,
            mask="auto",
        )
    else:
        c.setFillColor(MUTED)
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(x + width / 2, y + height / 2 - 5, "LOGO")


def draw_footer(c: canvas.Canvas) -> None:
    y = 46
    c.setStrokeColor(LINE)
    c.line(MARGIN, y + 17, PAGE_WIDTH - MARGIN, y + 17)
    draw_centered(c, COMPANY_NAME, y + 2, 8, "Helvetica-Bold", DARK)
    draw_centered(c, f"OFFICE ADD. {OFFICE_ADDRESS}", y - 10, 7.2, "Helvetica", MUTED)
    draw_centered(c, f"Email: {EMAIL}   Gst No: {GST_FOOTER}", y - 22, 7.2, "Helvetica", MUTED)


def draw_page_title(c: canvas.Canvas, title: str) -> None:
    draw_logo(c, MARGIN, PAGE_HEIGHT - 94, 94, 58)
    draw_right(c, COMPANY_NAME, PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 58, 14, "Helvetica-Bold", GREEN)
    draw_right(c, title, PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 76, 12, "Helvetica-Bold", DARK)
    c.setStrokeColor(GREEN)
    c.setLineWidth(1.2)
    c.line(MARGIN, PAGE_HEIGHT - 108, PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 108)


def draw_table(
    c: canvas.Canvas,
    x: float,
    y: float,
    col_widths: list[float],
    rows: list[list[str]],
    header_fill=GREEN,
    font_size: float = 8,
    leading: float = 9.5,
    min_height: float = 24,
) -> float:
    for row_index, row in enumerate(rows):
        wrapped_cells = [
            wrap_line(str(row[col_index] if col_index < len(row) else ""), col_widths[col_index] - 8, "Helvetica", font_size)
            for col_index in range(len(col_widths))
        ]
        row_height = max(min_height, max(len(lines) for lines in wrapped_cells) * leading + 10)
        fill = header_fill if row_index == 0 else colors.white
        text_color = colors.white if row_index == 0 else DARK
        font = "Helvetica-Bold" if row_index == 0 else "Helvetica"

        c.setFillColor(fill)
        c.setStrokeColor(LINE)
        c.rect(x, y - row_height, sum(col_widths), row_height, stroke=1, fill=1)

        cursor_x = x
        for col_index, width in enumerate(col_widths):
            if col_index:
                c.line(cursor_x, y, cursor_x, y - row_height)
            text_y = y - 13
            for line in wrapped_cells[col_index]:
                draw_text(c, line, cursor_x + 4, text_y, font_size, font, text_color)
                text_y -= leading
            cursor_x += width
        y -= row_height
    return y


def draw_cover_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_logo(c, MARGIN, PAGE_HEIGHT - 128, 118, 82)

    gst_x = PAGE_WIDTH - MARGIN - 182
    gst_y = PAGE_HEIGHT - 78
    c.setFillColor(SOFT_GOLD)
    c.setStrokeColor(GOLD)
    c.roundRect(gst_x, gst_y, 182, 32, 4, stroke=1, fill=1)
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(gst_x + 91, gst_y + 10, f"GST No: {GST_NO}")
    draw_right(c, COMPANY_NAME, PAGE_WIDTH - MARGIN, PAGE_HEIGHT - 112, 17, "Helvetica-Bold", GREEN)

    draw_centered(c, "QUOTATION", PAGE_HEIGHT - 225, 28, "Helvetica-Bold", DARK)
    draw_centered(c, "Proposal for installation, Engineering and Procurement of", PAGE_HEIGHT - 270, 13, "Helvetica-Bold", DARK)
    draw_centered(c, "ON GRID Connected Roof-Top Solar PV System.", PAGE_HEIGHT - 290, 13, "Helvetica-Bold", DARK)

    box_w = PAGE_WIDTH - (MARGIN * 2) - 80
    x = MARGIN + 40
    y = PAGE_HEIGHT - 410
    c.setStrokeColor(LINE)
    c.setFillColor(colors.white)
    c.roundRect(x, y - 142, box_w, 142, 8, stroke=1, fill=1)
    details = [
        ("NAME -", quotation.customer_name),
        ("LOCATION -", quotation.location),
        ("CAPACITY -", capacity_kw(quotation)),
        ("DATE -", quotation.quote_date),
    ]
    row_y = y - 30
    for label, value in details:
        draw_text(c, label, x + 28, row_y, 12, "Helvetica-Bold", DARK)
        draw_text(c, value, x + 130, row_y, 12, "Helvetica", DARK)
        row_y -= 28

    draw_footer(c)


def draw_introduction_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_page_title(c, "INTRODUCTION")
    y = PAGE_HEIGHT - 142
    for block in paragraph_blocks(quotation.introduction):
        if block.lower().startswith("why to go solar"):
            y -= 8
            draw_text(c, block, MARGIN, y, 13, "Helvetica-Bold", GREEN)
            y -= 22
            continue
        y = draw_wrapped(c, block, MARGIN, y, PAGE_WIDTH - (MARGIN * 2), 9.6, 13, "Helvetica")
        y -= 7
    draw_footer(c)


def draw_specification_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_page_title(c, "SYSTEM SPECIFICATION")
    rows = [["ITEM", "BRAND/PARAMETERS", "DETAIL", "SPECIFICATION", "QUANTITY"]]
    rows.extend(
        [
            [spec.item, spec.brand_parameters, spec.detail, spec.specification, spec.quantity]
            for spec in quotation.specifications
        ]
    )
    draw_table(c, MARGIN, PAGE_HEIGHT - 142, [92, 126, 116, 108, 80], rows, font_size=7.2, leading=8.5, min_height=28)
    draw_footer(c)


def draw_price_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_page_title(c, "PRICE QUOTION")
    y = PAGE_HEIGHT - 150
    c.setFillColor(SOFT_GREEN)
    c.setStrokeColor(GREEN)
    c.roundRect(MARGIN, y - 40, PAGE_WIDTH - (MARGIN * 2), 40, 5, stroke=1, fill=1)
    draw_centered(c, quotation.price_header or f"{capacity_kw(quotation)} SANCTION AND {capacity_kw(quotation)} PV MODULE", y - 16, 12, "Helvetica-Bold", DARK)
    y -= 64

    rows = [["SR NO", "SYSTEM DETAIL", "RATE"]]
    for row in quotation.price_rows:
        rate = row.rate
        if row.system_detail.strip().upper() == "TOTAL":
            rate = final_price_for_row(quotation.final_price)
        rows.append([row.sr_no, row.system_detail, rate])
    y = draw_table(c, MARGIN + 34, y, [80, 260, 126], rows, font_size=9, leading=11, min_height=30)
    y -= 28
    draw_wrapped(c, f"* {quotation.price_note}", MARGIN + 20, y, PAGE_WIDTH - (MARGIN * 2) - 40, 10, 13, "Helvetica-Bold", DARK)
    y -= 72
    c.setFillColor(SOFT_GOLD)
    c.setStrokeColor(GOLD)
    c.roundRect(MARGIN, y - 50, PAGE_WIDTH - (MARGIN * 2), 50, 5, stroke=1, fill=1)
    draw_centered(c, quotation.payment_note, y - 21, 10, "Helvetica-Bold", DARK)
    draw_footer(c)


def draw_terms_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_page_title(c, "TERMS & CONDITIONS")
    columns = [
        (MARGIN, PAGE_HEIGHT - 136, (PAGE_WIDTH - (MARGIN * 2) - 18) / 2),
        (MARGIN + (PAGE_WIDTH - (MARGIN * 2) + 18) / 2, PAGE_HEIGHT - 136, (PAGE_WIDTH - (MARGIN * 2) - 18) / 2),
    ]
    column_index = 0
    x, y, width = columns[column_index]
    bottom = 78

    for block in paragraph_blocks(quotation.terms_conditions):
        font = "Helvetica-Bold" if block.upper().startswith("TERMS") else "Helvetica"
        size = 6.8 if font == "Helvetica" else 7.6
        leading = 8.2
        lines = wrap_line(block, width, font, size)
        needed = len(lines) * leading + 5
        if y - needed < bottom and column_index == 0:
            column_index = 1
            x, y, width = columns[column_index]
        for line in lines:
            if y < bottom:
                break
            draw_text(c, line, x, y, size, font, DARK)
            y -= leading
        y -= 5
    draw_footer(c)


def draw_bank_page(c: canvas.Canvas, quotation: Quotation) -> None:
    draw_page_title(c, "BANK DETAILS + SIGNATURE")
    y = PAGE_HEIGHT - 152
    c.setFillColor(colors.white)
    c.setStrokeColor(LINE)
    c.roundRect(MARGIN, y - 92, PAGE_WIDTH - (MARGIN * 2), 92, 6, stroke=1, fill=1)
    draw_text(c, "REMARK IF ANY", MARGIN + 18, y - 24, 11, "Helvetica-Bold", DARK)
    draw_wrapped(c, quotation.remark or "-", MARGIN + 18, y - 48, PAGE_WIDTH - (MARGIN * 2) - 36, 10, 13, "Helvetica")

    y -= 132
    bank_rows = [
        ["MAIL ID", EMAIL],
        ["CONTACT NO", CONTACT_NO],
        ["ACCOUNT HOLDER", ACCOUNT_HOLDER],
        ["BANK NAME", BANK_NAME],
        ["CC ACCOUNT NO", CC_ACCOUNT_NO],
        ["IFSC CODE", IFSC_CODE],
        ["GST", GST_NO],
        ["BRANCH", BRANCH],
    ]
    draw_table(c, MARGIN + 30, y, [170, 310], [["PARTICULAR", "DETAIL"], *bank_rows], font_size=9, leading=11, min_height=28)

    sig_y = 154
    c.setStrokeColor(DARK)
    c.line(PAGE_WIDTH - 255, sig_y, PAGE_WIDTH - MARGIN, sig_y)
    draw_right(c, "SIGNATURE & SEAL OF DIRECTOR", PAGE_WIDTH - MARGIN, sig_y - 18, 10, "Helvetica-Bold", DARK)
    draw_footer(c)


def create_pdf(quotation: Quotation) -> bytes:
    output = BytesIO()
    c = canvas.Canvas(output, pagesize=A4)

    draw_cover_page(c, quotation)
    c.showPage()
    draw_introduction_page(c, quotation)
    c.showPage()
    draw_specification_page(c, quotation)
    c.showPage()
    draw_price_page(c, quotation)
    c.showPage()
    draw_terms_page(c, quotation)
    c.showPage()
    draw_bank_page(c, quotation)
    c.save()

    return output.getvalue()


def set_doc_defaults(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)
    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)


def add_doc_title(document: Document, title: str) -> None:
    if LOGO_PATH.exists():
        document.add_picture(str(LOGO_PATH), width=Inches(1.35))
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT
    heading = document.add_paragraph()
    heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = heading.add_run(COMPANY_NAME)
    run.bold = True
    run.font.size = Pt(16)
    sub = document.add_paragraph(title)
    sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sub.runs[0].bold = True
    sub.runs[0].font.size = Pt(12)


def add_doc_table(document: Document, rows: list[list[str]]) -> None:
    table = document.add_table(rows=1, cols=len(rows[0]))
    table.style = "Table Grid"
    for index, header in enumerate(rows[0]):
        cell = table.rows[0].cells[index]
        cell.text = clean_text(header)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    for row in rows[1:]:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            cells[index].text = clean_text(value)


def add_bold_paragraph(document: Document, text: str) -> None:
    paragraph = document.add_paragraph()
    run = paragraph.add_run(clean_text(text))
    run.bold = True


def add_page_break(document: Document) -> None:
    document.add_page_break()


def create_word(quotation: Quotation) -> bytes:
    document = Document()
    set_doc_defaults(document)

    add_doc_title(document, "QUOTATION")
    document.add_paragraph(
        "Proposal for installation, Engineering and Procurement of ON GRID Connected Roof-Top Solar PV System."
    ).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_doc_table(
        document,
        [
            ["Field", "Details"],
            ["Quotation Number", quotation.quotation_number],
            ["NAME", quotation.customer_name],
            ["LOCATION", quotation.location],
            ["CAPACITY", capacity_kw(quotation)],
            ["DATE", quotation.quote_date],
            ["GST No", GST_NO],
        ],
    )

    add_page_break(document)
    add_doc_title(document, "INTRODUCTION")
    for block in paragraph_blocks(quotation.introduction):
        paragraph = document.add_paragraph(clean_text(block))
        if block.lower().startswith("why to go solar"):
            paragraph.runs[0].bold = True

    add_page_break(document)
    add_doc_title(document, "SYSTEM SPECIFICATION")
    add_doc_table(
        document,
        [
            ["ITEM", "BRAND/PARAMETERS", "DETAIL", "SPECIFICATION", "QUANTITY"],
            *[
                [spec.item, spec.brand_parameters, spec.detail, spec.specification, spec.quantity]
                for spec in quotation.specifications
            ],
        ],
    )

    add_page_break(document)
    add_doc_title(document, "PRICE QUOTION")
    add_bold_paragraph(document, quotation.price_header)
    add_doc_table(
        document,
        [
            ["SR NO", "SYSTEM DETAIL", "RATE"],
            *[
                [
                    row.sr_no,
                    row.system_detail,
                    final_price_for_row(quotation.final_price)
                    if row.system_detail.strip().upper() == "TOTAL"
                    else row.rate,
                ]
                for row in quotation.price_rows
            ],
        ],
    )
    document.add_paragraph(f"* {clean_text(quotation.price_note)}")
    add_bold_paragraph(document, quotation.payment_note)

    add_page_break(document)
    add_doc_title(document, "TERMS & CONDITIONS")
    for block in paragraph_blocks(quotation.terms_conditions):
        document.add_paragraph(clean_text(block))

    add_page_break(document)
    add_doc_title(document, "BANK DETAILS + SIGNATURE")
    add_bold_paragraph(document, "REMARK IF ANY")
    document.add_paragraph(clean_text(quotation.remark or "-"))
    add_doc_table(
        document,
        [
            ["PARTICULAR", "DETAIL"],
            ["MAIL ID", EMAIL],
            ["CONTACT NO", CONTACT_NO],
            ["ACCOUNT HOLDER", ACCOUNT_HOLDER],
            ["BANK NAME", BANK_NAME],
            ["CC ACCOUNT NO", CC_ACCOUNT_NO],
            ["IFSC CODE", IFSC_CODE],
            ["GST", GST_NO],
            ["BRANCH", BRANCH],
        ],
    )
    document.add_paragraph("\n\nSIGNATURE & SEAL OF DIRECTOR").alignment = WD_ALIGN_PARAGRAPH.RIGHT

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def write_sheet(ws, rows: list[list[object]]) -> None:
    header_fill = PatternFill("solid", fgColor="176B3A")
    for row in rows:
        ws.append([clean_text(value) for value in row])
    if ws.max_row:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    for column in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in column)
        ws.column_dimensions[column[0].column_letter].width = min(max_length + 4, 48)


def create_excel(quotation: Quotation) -> bytes:
    workbook = Workbook()
    cover = workbook.active
    cover.title = "01 Cover Page"
    write_sheet(
        cover,
        [
            ["Field", "Details"],
            ["Company", COMPANY_NAME],
            ["Quotation Number", quotation.quotation_number],
            ["Customer Name", quotation.customer_name],
            ["Location", quotation.location],
            ["Capacity", capacity_kw(quotation)],
            ["Date", quotation.quote_date],
            ["GST No", GST_NO],
            ["Office Address", OFFICE_ADDRESS],
            ["Email", EMAIL],
        ],
    )
    write_sheet(workbook.create_sheet("02 Introduction"), [["Introduction"], *[[block] for block in paragraph_blocks(quotation.introduction)]])
    write_sheet(
        workbook.create_sheet("03 Specification"),
        [
            ["ITEM", "BRAND/PARAMETERS", "DETAIL", "SPECIFICATION", "QUANTITY"],
            *[
                [spec.item, spec.brand_parameters, spec.detail, spec.specification, spec.quantity]
                for spec in quotation.specifications
            ],
        ],
    )
    write_sheet(
        workbook.create_sheet("04 Price"),
        [
            ["PRICE QUOTION", quotation.price_header],
            ["SR NO", "SYSTEM DETAIL", "RATE"],
            *[
                [
                    row.sr_no,
                    row.system_detail,
                    final_price_for_row(quotation.final_price)
                    if row.system_detail.strip().upper() == "TOTAL"
                    else row.rate,
                ]
                for row in quotation.price_rows
            ],
            ["NOTE", quotation.price_note],
            ["PAYMENT NOTE", quotation.payment_note],
        ],
    )
    write_sheet(workbook.create_sheet("05 Terms"), [["Terms & Conditions"], *[[block] for block in paragraph_blocks(quotation.terms_conditions)]])
    write_sheet(
        workbook.create_sheet("06 Bank Details"),
        [
            ["PARTICULAR", "DETAIL"],
            ["REMARK IF ANY", quotation.remark or "-"],
            ["SIGNATURE & SEAL", "SIGNATURE & SEAL OF DIRECTOR"],
            ["MAIL ID", EMAIL],
            ["CONTACT NO", CONTACT_NO],
            ["ACCOUNT HOLDER", ACCOUNT_HOLDER],
            ["BANK NAME", BANK_NAME],
            ["CC ACCOUNT NO", CC_ACCOUNT_NO],
            ["IFSC CODE", IFSC_CODE],
            ["GST", GST_NO],
            ["BRANCH", BRANCH],
        ],
    )

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


@app.post("/save-quotation")
def save_quotation(quotation: Quotation) -> dict:
    global saved_quotations
    saved_quotations = [item for item in saved_quotations if item.id != quotation.id]
    saved_quotations.insert(0, quotation)
    return {"success": True, "message": "Quotation accepted. Browser localStorage remains the storage for V1."}


@app.get("/history")
def history() -> dict:
    return {"quotations": [quotation_dump(item) for item in saved_quotations]}


@app.post("/generate-pdf")
def generate_pdf(quotation: Quotation) -> Response:
    content = create_pdf(quotation)
    return Response(
        content,
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{export_name(quotation, "pdf")}"'},
    )


@app.post("/generate-word")
def generate_word(quotation: Quotation) -> Response:
    content = create_word(quotation)
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{export_name(quotation, "docx")}"'},
    )


@app.post("/generate-excel")
def generate_excel(quotation: Quotation) -> Response:
    content = create_excel(quotation)
    return Response(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{export_name(quotation, "xlsx")}"'},
    )


app.mount("/", StaticFiles(directory=str(BASE_DIR), html=True), name="static")
