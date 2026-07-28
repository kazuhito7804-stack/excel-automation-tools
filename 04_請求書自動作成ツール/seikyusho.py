import unicodedata
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from collections import OrderedDict
from datetime import date

INPUT_FILE = "請求データ/seikyusaki.xlsx"
OUTPUT_FILE = "seikyusho_kekka.xlsx"

TITLE_FONT = Font(name="游ゴシック", size=16, bold=True)
HEADER_FONT = Font(name="游ゴシック", size=11, bold=True, color="FFFFFF")
NORMAL_FONT = Font(name="游ゴシック", size=11)
TOTAL_FONT = Font(name="游ゴシック", size=11, bold=True)

CURRENCY_FMT = "¥#,##0"

HEADER_FILL = PatternFill(fill_type="solid", start_color="4472C4", end_color="4472C4")
TOTAL_FILL = PatternFill(fill_type="solid", start_color="F2F2F2", end_color="F2F2F2")

THIN = Side(style="thin", color="808080")
MEDIUM = Side(style="medium", color="000000")
GRID_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

RIGHT_ALIGN = Alignment(horizontal="right", vertical="center")
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
LEFT_ALIGN = Alignment(horizontal="left", vertical="center")

# 列の最低幅(文字の全角/半角を考慮した実測幅に対する下限)
MIN_WIDTHS = {"A": 14, "B": 10, "C": 8, "D": 12}


def visual_width(text):
    """全角文字を2、半角文字を1として文字列の表示幅を概算する"""
    width = 0
    for ch in str(text):
        width += 2 if unicodedata.east_asian_width(ch) in ("W", "F", "A") else 1
    return width


def load_items(path):
    """入力Excel(会社名・品目・単価・数量)を読み込み、会社名ごとにまとめる"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    companies = OrderedDict()
    for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
        company, item, unit_price, qty = row
        if not company or not item:
            continue
        companies.setdefault(company, []).append((item, unit_price, qty))
    return companies


def build_invoice_sheet(ws, company, items):
    col_max_width = dict(MIN_WIDTHS)

    def track(col_letter, text):
        col_max_width[col_letter] = max(col_max_width[col_letter], visual_width(text))

    # タイトル
    ws["A1"] = "請求書"
    ws["A1"].font = TITLE_FONT

    # 宛名
    atena = f"{company}御中"
    ws["A3"] = atena
    ws["A3"].font = NORMAL_FONT
    track("A", atena)

    # 請求日
    seikyubi = f"請求日：{date.today().strftime('%Y年%m月%d日')}"
    ws["A4"] = seikyubi
    ws["A4"].font = NORMAL_FONT
    track("A", seikyubi)

    # 見出し行
    headers = ["品目", "単価", "数量", "金額"]
    header_row = 6
    for col_idx, text in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        cell = ws.cell(row=header_row, column=col_idx, value=text)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = GRID_BORDER
        track(col_letter, text)

    # 明細行
    first_item_row = header_row + 1
    for offset, (item, unit_price, qty) in enumerate(items):
        r = first_item_row + offset

        c_item = ws.cell(row=r, column=1, value=item)
        c_item.font = NORMAL_FONT
        c_item.alignment = LEFT_ALIGN
        c_item.border = GRID_BORDER
        track("A", item)

        c_price = ws.cell(row=r, column=2, value=unit_price)
        c_price.font = NORMAL_FONT
        c_price.number_format = CURRENCY_FMT
        c_price.alignment = RIGHT_ALIGN
        c_price.border = GRID_BORDER
        track("B", f"¥{unit_price:,}")

        c_qty = ws.cell(row=r, column=3, value=qty)
        c_qty.font = NORMAL_FONT
        c_qty.alignment = CENTER_ALIGN
        c_qty.border = GRID_BORDER
        track("C", qty)

        c_amount = ws.cell(row=r, column=4, value=f"=B{r}*C{r}")
        c_amount.font = NORMAL_FONT
        c_amount.number_format = CURRENCY_FMT
        c_amount.alignment = RIGHT_ALIGN
        c_amount.border = GRID_BORDER
        track("D", f"¥{unit_price * qty:,}")

    last_item_row = first_item_row + len(items) - 1

    # 小計・消費税・合計金額(明細の1行空けて配置)
    subtotal_row = last_item_row + 2
    tax_row = subtotal_row + 1
    total_row = tax_row + 1

    def set_summary_row(row, label, formula, font, fill=None):
        c_label = ws.cell(row=row, column=3, value=label)
        c_label.font = font
        c_label.alignment = RIGHT_ALIGN
        c_label.border = GRID_BORDER
        if fill:
            c_label.fill = fill
        track("C", label)

        c_value = ws.cell(row=row, column=4, value=formula)
        c_value.font = font
        c_value.number_format = CURRENCY_FMT
        c_value.alignment = RIGHT_ALIGN
        c_value.border = GRID_BORDER
        if fill:
            c_value.fill = fill

    set_summary_row(subtotal_row, "小計", f"=SUM(D{first_item_row}:D{last_item_row})", NORMAL_FONT)
    set_summary_row(tax_row, "消費税(10%)", f"=D{subtotal_row}*0.1", NORMAL_FONT)
    set_summary_row(total_row, "合計金額", f"=D{subtotal_row}+D{tax_row}", TOTAL_FONT, fill=TOTAL_FILL)

    # 合計金額行の上下は太めの罫線で強調
    for col in range(1, 5):
        cell = ws.cell(row=total_row, column=col)
        border = cell.border
        cell.border = Border(left=border.left, right=border.right, top=MEDIUM, bottom=MEDIUM)

    # 列幅(内容の実測幅+余白と最低幅の大きい方)
    for col, width in col_max_width.items():
        ws.column_dimensions[col].width = width + 3

    # 行の高さ(印刷時に詰まって見えないよう調整)
    ws.row_dimensions[1].height = 28  # タイトル
    ws.row_dimensions[3].height = 20  # 宛名
    ws.row_dimensions[4].height = 18  # 請求日
    ws.row_dimensions[header_row].height = 20
    for r in range(first_item_row, last_item_row + 1):
        ws.row_dimensions[r].height = 18
    ws.row_dimensions[subtotal_row].height = 18
    ws.row_dimensions[tax_row].height = 18
    ws.row_dimensions[total_row].height = 20

    # A4用紙1ページに収まるようページ設定
    ws.sheet_view.showGridLines = False
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(
        left=0.7, right=0.7, top=0.9, bottom=0.75, header=0.3, footer=0.3
    )
    ws.print_area = f"A1:D{total_row}"


def main():
    companies = load_items(INPUT_FILE)

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    for company, items in companies.items():
        # シート名に使えない文字を除去し31文字以内に収める
        safe_name = company[:31]
        ws = out_wb.create_sheet(title=safe_name)
        build_invoice_sheet(ws, company, items)

    out_wb.save(OUTPUT_FILE)
    print("請求書を作成しました")


if __name__ == "__main__":
    main()
