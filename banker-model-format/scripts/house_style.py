"""House style for banker Excel models (openpyxl), distilled from the Project Print model (Xerox RX, v18).

Import it in a build script; every color, number format and geometry value lives here so nothing is typed from
memory. The values are the evidence from one model: if the user hands over a workbook with its own established
format, read that workbook and override these constants (see SKILL.md, "Preserve before impose").

    from house_style import *
    wb = new_workbook()
    cover_page(wb, "Project Atlas", "Acquisition Model", "Rent roll as of 9/30/26", sections=[...])
    t = Tab(wb, "Operating Pro Forma", "Atlas Apartments", data_cols="E:N")
    t.header("Operating Pro Forma, Annual")
    ...
    finalize(wb)
    wb.save(path)

Run `python3 house_style.py --selftest out.xlsx` to write a one-tab specimen of every row type.
"""
from copy import copy

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

# ---------------------------------------------------------------- type
FONT = "Garamond"
SIZE = 10

# ---------------------------------------------------------------- font colors (the cover page legend)
INK = "000000"      # formulas and calculations on the same tab
BLUE = "0000FF"     # hard-coded inputs and historical data
GREEN = "008000"    # links to other tabs
RED = "FF0000"      # checks / highlight values
WHITE = "FFFFFF"
GREY_A1 = "A0A0A0"  # hidden A1 title echo and divider-tab labels
SUBTLE = "808080"   # cover subtitle
HLINK = "0563C1"

# ---------------------------------------------------------------- fills (theme dk2 = 0E2841 navy, lt2 = E8E8E8 grey)
NAVY = "002060"         # Heading 1: section band
GREY_H2 = "808080"      # Heading 2: secondary band (rare)
HDR = "E8E8E8"          # column header rows (lt2)
HDR_DK = "D1D1D1"       # dense headers, e.g. a comps table (lt2 -10%)
PROJ = "215F9A"         # "Projected" part of a band (dk2 +25%)
LTM = "4E95D9"          # LTM / stub column header (dk2 +50%)
ASIDE = "A6CAEC"        # side blocks: sensitivity headers, method sub-bands (dk2 +75%)
KEY_TOTAL = "DCEAF7"    # key totals: Total Debt, Net Debt, NOI, Levered CF (dk2 +90%)
INPUT = "FFF2CC"        # driver / assumption cells the user is meant to change

# tab colors by workbook section, darkest first (cover is grey)
TAB_COVER = "D1D1D1"
TAB_SECTIONS = ["0E2841", "215F9A", "4E95D9", "A6CAEC", "E8E8E8", "747474"]

# ---------------------------------------------------------------- number formats (copy exactly; never retype)
_Z = '_("–"_)_%;_(@_)_%'              # zero shows an en dash, text aligns with the numbers
NF = {
    "num": '_(#,##0_)_%;\\(#,##0\\)_%;' + _Z,                    # 1,234   (1,234)   –
    "num1": '_(#,##0.0_)_%;\\(#,##0.0\\)_%;' + _Z,               # prices, $ per SF with decimals
    "num2": '_(#,##0.00_)_%;\\(#,##0.00\\)_%;' + _Z,
    "dol": '_([$$]#,##0_)_%;\\([$$]#,##0\\)_%;' + _Z,            # first row of a block and totals
    "dol1": '_([$$]#,##0.0_)_%;\\([$$]#,##0.0\\)_%;' + _Z,
    "dol2": '_([$$]#,##0.00_)_%;\\([$$]#,##0.00\\)_%;' + _Z,     # per-share, per-SF rents
    "pct": '_(#,##0.0%_);\\(#,##0.0%\\);' + _Z,                  # growth, margins, recoveries
    "pct2": '_(#,##0.00%_);\\(#,##0.00%\\);' + _Z,               # cap rates, rates, input %
    "pct3": '_(#,##0.000%_);\\(#,##0.000%\\);' + _Z,             # coupons, SOFR
    "mult": '_(#,##0.0"x"_)_%;\\(#,##0.0"x"\\)_%;' + _Z,         # 5.3x  leverage, multiples, DSCR
    "mult2": '_(#,##0.00"x"_)_%;\\(#,##0.00"x"\\)_%;' + _Z,      # 1.25x DSCR, equity multiple
    "year_a": 'General"A"',                                       # 2025A
    "year_e": '0"E"',                                             # 2026E
    "year_n": '"Year "0',                                         # Year 3 (real estate hold years)
    "month": "mmm-yy",                                            # Mar-28
    "date": "m/d/yy;@",
    "spread": '"S + "#',                                          # S + 400
    "flag": "0_)_%",
    "onoff": '[=1]"On";[=0]"Off";General',
    "hidden": ";;;",
}

# ---------------------------------------------------------------- geometry
W_MARGIN = 1.73      # columns A, B and the right margin column
W_LABEL = 42         # label column C (40 to 60 depending on label length)
W_DATA = 13          # every data column
H_ROW = 12.75
H_TITLE = 14.25
H_SPACER = 3.0       # thin row under a column header

THIN = Side(style="thin")
HAIR = Side(style="hair")


# ======================================================================= low-level
def font(color=INK, b=False, i=False, u=None, size=SIZE):
    return Font(name=FONT, sz=size, b=b, i=i, u=u, color="FF" + color)


def fill(color):
    return PatternFill("solid", fgColor="FF" + color)


def note(cell, text, author="Note", width=320):
    """Sources, rationale and caveats go in a cell comment, never in a notes column."""
    cm = Comment(text, author)
    cm.width, cm.height = width, max(70, min(260, 18 * (len(text) // 48 + 2)))
    cell.comment = cm
    return cell


def cols_between(a, b):
    return [get_column_letter(k) for k in range(column_index_from_string(a), column_index_from_string(b) + 1)]


# role -> (font kwargs, fill, border sides, default number format, horizontal align, indent)
ROLES = {
    "label": (dict(), None, (), None, "left", 1),
    "label_head": (dict(b=True), None, (), None, "left", 0),           # first line of a block (Revenue)
    "label_sub": (dict(i=True), None, (), None, "left", 2),            # % growth, margin, memo lines
    "label_total": (dict(b=True), None, ("top",), None, "left", 0),
    "label_key": (dict(b=True), KEY_TOTAL, ("top", "bottom"), None, "left", 0),
    "label_memo": (dict(b=True), HDR, ("top_h", "bottom_h"), None, "left", 0),
    "label_check": (dict(i=True), None, (), None, "left", 2),
    "input": (dict(color=BLUE), INPUT, (), "num", "right", 0),         # a driver the user changes
    "hard": (dict(color=BLUE), None, (), "num", "right", 0),           # historical / sourced data
    "link": (dict(color=GREEN), None, (), "num", "right", 0),          # pulls from another tab
    "calc": (dict(), None, (), "num", "right", 0),
    "calc_sub": (dict(i=True), None, (), "pct", "right", 0),
    "total": (dict(b=True), None, ("top",), "dol", "right", 0),
    "key": (dict(b=True), KEY_TOTAL, ("top", "bottom"), "dol", "right", 0),
    "memo": (dict(b=True), HDR, ("top_h", "bottom_h"), "dol", "right", 0),
    "check": (dict(i=True), None, (), "num", "right", 0),
}


def style(cell, role, fmt=None, color=None, b=None, i=None, align=None):
    f_kw, fl, sides, nf, h, ind = ROLES[role]
    kw = dict(f_kw)
    if color:
        kw["color"] = color
    if b is not None:
        kw["b"] = b
    if i is not None:
        kw["i"] = i
    cell.font = font(**kw)
    cell.fill = fill(fl) if fl else PatternFill(fill_type=None)
    bd = {}
    for s in sides:
        bd[s.split("_")[0]] = HAIR if s.endswith("_h") else THIN
    cell.border = Border(**bd)
    if fmt or nf:
        cell.number_format = NF.get(fmt or nf, fmt or nf)
    cell.alignment = Alignment(horizontal=align or h, indent=ind)
    return cell


# ======================================================================= workbook
def new_workbook():
    """Empty workbook whose default (Normal) font is Garamond 10, so blank cells match too."""
    wb = Workbook()
    wb._fonts[0] = Font(name=FONT, sz=SIZE)   # font id 0 backs the Normal style
    wb.remove(wb.active)
    return wb


class Tab:
    """One model tab in house layout.

    Columns: A and B are 1.73 wide margin/marker columns, C holds labels, data_cols hold numbers, and one more
    1.73-wide column closes the right margin. Rows: 1 = A1 title echo + link to cover, 3 = title over a thin
    rule with the entity right-aligned, 4 = units line (italic), 6 = optional case row, 7/8 = first band.
    """

    def __init__(self, wb, name, entity, data_cols="D:J", label_width=W_LABEL, tab_color=None, index=None):
        self.ws = wb.create_sheet(name, index) if index is not None else wb.create_sheet(name)
        self.name, self.entity = name, entity
        a, b = data_cols.split(":")
        self.cols = cols_between(a, b)
        self.first, self.last = self.cols[0], self.cols[-1]
        self.right = get_column_letter(column_index_from_string(self.last) + 1)
        ws = self.ws
        ws.sheet_view.showGridLines = False
        for c, w in (("A", W_MARGIN), ("B", W_MARGIN), ("C", label_width), (self.right, W_MARGIN)):
            ws.column_dimensions[c].width = w
        for c in self.cols:
            ws.column_dimensions[c].width = W_DATA
        if tab_color:
            ws.sheet_properties.tabColor = tab_color
        self._marker_rows = []

    # ---------------------------------------------------------------- top of tab
    def header(self, title, units="All values in $ millions (USD)", cover="Cover Page"):
        ws = self.ws
        ws["A1"] = "=C3"
        ws["A1"].font, ws["A1"].number_format = font(GREY_A1), NF["hidden"]
        ws["B1"] = f"='{cover}'!A1"
        ws["B1"].font = Font(name=FONT, sz=11, u="single", color="FF" + HLINK)
        ws["B1"].hyperlink = f"#'{cover}'!A1"
        for c in ["C"] + self.cols:
            cell = ws[f"{c}3"]
            cell.font, cell.border = font(b=True), Border(bottom=THIN)
        ws["C3"] = title
        ws["C3"].alignment = Alignment(horizontal="left")
        ws[f"{self.last}3"] = self.entity
        ws[f"{self.last}3"].alignment = Alignment(horizontal="right")
        ws["C4"] = units
        ws["C4"].font = font(i=True)
        ws.row_dimensions[3].height = H_TITLE
        for r in (1, 2, 4):
            ws.row_dimensions[r].height = H_ROW

    def case_row(self, row=6, ref="'Cover Page'!$F$21", label="Active Case (select on Cover Page)"):
        ws = self.ws
        style(ws[f"C{row}"], "label_head").value = label
        c = ws[f"{self.first}{row}"]
        c.value = f"={ref}"
        c.font, c.fill = font(GREEN, b=True), fill(INPUT)
        c.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
        c.alignment = Alignment(horizontal="center")

    # ---------------------------------------------------------------- section band + column headers
    def band(self, row, title, texts=None, proj_from=None, aside_cols=(), markers=True):
        """Navy Heading-1 band. texts: {col: italic text}, centered across the following empty band cells.
        proj_from: first column painted in the lighter 'Projected' fill. aside_cols: columns in the ASIDE fill."""
        ws = self.ws
        if markers:
            for m in "AB":
                ws[f"{m}{row}"] = "x"
                ws[f"{m}{row}"].font = font()
            self._marker_rows.append(row)
        c = ws[f"C{row}"]
        c.value, c.font, c.fill = title, font(WHITE, b=True), fill(NAVY)
        c.border, c.alignment = Border(top=THIN), Alignment(horizontal="left")
        proj = False
        for col in self.cols:
            proj = proj or (proj_from == col)
            cell = ws[f"{col}{row}"]
            color = ASIDE if col in aside_cols else (PROJ if proj else NAVY)
            cell.fill, cell.border = fill(color), Border(top=THIN)
            cell.font = font(WHITE if color != ASIDE else INK, i=True, b=color == ASIDE)
            cell.alignment = Alignment(horizontal="centerContinuous")
            if texts and col in texts:
                cell.value = texts[col]
        ws.row_dimensions[row].height = H_ROW

    def col_header(self, row, labels, last=True, label_cell=None, fill_color=HDR, spacer=True):
        """Grey header row. labels: {col: text or number}; years can be numbers formatted with fmt_year()."""
        ws = self.ws
        for col in ["C"] + self.cols:
            cell = ws[f"{col}{row}"]
            cell.font, cell.fill = font(b=True), fill(fill_color)
            cell.border = Border(bottom=THIN) if last else Border()
            cell.alignment = Alignment(horizontal="left" if col == "C" else "center")
            if col in labels:
                cell.value = labels[col]
        if label_cell:
            ws[f"C{row}"] = label_cell
        if last and spacer:
            ws.row_dimensions[row + 1].height = H_SPACER

    def year_header(self, row, first_year, n_hist, ltm_col=None, last=True):
        """Years as numbers: first is typed, the rest =prev+1; actuals show 2025A, projections 2026E."""
        ws = self.ws
        self.col_header(row, {}, last=last)
        for k, col in enumerate(self.cols):
            cell = ws[f"{col}{row}"]
            cell.value = first_year if k == 0 else f"={self.cols[k - 1]}{row}+1"
            cell.number_format = NF["year_a"] if k < n_hist else NF["year_e"]
            if col == ltm_col:
                cell.value, cell.number_format, cell.fill = "LTM", "General", fill(LTM)
                cell.font = font(WHITE, b=True)

    # ---------------------------------------------------------------- rows
    def row(self, r, label, values=None, kind="calc", fmt=None, label_kind=None, cols=None):
        """Write one line. values: list (one per data col, None skips), dict {col: value}, or a callable
        col -> value. kind: input / hard / link / calc / calc_sub / total / key / memo / check."""
        ws = self.ws
        lk = label_kind or {"total": "label_total", "key": "label_key", "memo": "label_memo",
                            "check": "label_check", "calc_sub": "label_sub"}.get(kind, "label")
        style(ws[f"C{r}"], lk).value = label
        cols = cols or self.cols
        if isinstance(values, dict):
            items = values.items()
        elif callable(values):
            items = ((c, values(c)) for c in cols)
        else:
            items = zip(cols, values or [])
        if kind in ("total", "key", "memo"):        # rule / fill runs across the row's columns
            for c in cols:
                style(ws[f"{c}{r}"], kind, fmt)
        for c, v in items:
            if v is None:
                continue
            cell = ws[f"{c}{r}"]
            is_f = isinstance(v, str) and v.startswith("=")
            if kind in ("total", "key", "memo", "check", "calc_sub", "input"):
                hard = isinstance(v, (int, float)) and kind != "input"
                style(cell, kind, fmt, color=GREEN if (is_f and "!" in v) else (BLUE if hard else None))
            elif is_f:
                style(cell, "link" if "!" in v else "calc", fmt)   # cross-tab green, same-tab black
            elif isinstance(v, (int, float)):
                style(cell, "hard", fmt)               # a typed number is blue, never black
            else:
                style(cell, kind, fmt)
            cell.value = v
        ws.row_dimensions[r].height = H_ROW
        return r

    def check(self, r, label, formula_by_col, tol=0.5):
        """Italic check line that turns red when it is not zero."""
        self.row(r, label, formula_by_col, kind="check")
        rng = f"{self.first}{r}:{self.last}{r}"
        self.ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f"ABS({self.first}{r})>{tol}"], font=Font(color="FF" + RED)))

    def end(self, row):
        """'End' marker one row below the last block; x markers above it let Ctrl+arrow jump block to block."""
        ws = self.ws
        for m in "AB":
            ws[f"{m}{row - 1}"] = "x"
            ws[f"{m}{row - 1}"].font = font()
        ws[f"A{row}"] = "End"
        ws[f"A{row}"].font = font()


# ======================================================================= special tabs
def divider(wb, name, tab_color, cover="Cover Page"):
    """Section divider tab, e.g. 'Financial Statement Inputs >': one grey label and a link home."""
    ws = wb.create_sheet(name if name.endswith(">") else name + " >")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 60
    ws.sheet_properties.tabColor = tab_color
    ws.column_dimensions["A"].width = 24
    ws["A1"] = ws.title
    ws["A1"].font = font(GREY_A1, b=True)
    ws["B1"] = f"='{cover}'!A1"
    ws["B1"].font = font(HLINK)
    ws["B1"].hyperlink = f"#'{cover}'!A1"
    return ws


def cover_page(wb, entity, subtitle, data_as_of, sections=(), sources=(), case_label=None, cases=None,
               units="All $ in millions (USD), except per-unit values"):
    """Cover: entity + subtitle, Last Updated / Data As Of, color legend, sources, optional case selector,
    and a linked table of contents (each link shows the target tab's title via its hidden A1 echo).

    sections: [(heading, [tab names])]. cases: e.g. ["Bear", "Base", "Bull"] -> selector in F21 (input cell).
    """
    ws = wb.create_sheet("Cover Page", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = TAB_COVER
    for c, w in zip("ABCDEFGHIJ", (W_MARGIN, 13, 22, 15.5, 13, 13, 15.7, 15.5, 13, 13)):
        ws.column_dimensions[c].width = w
    ws["A1"] = "Return to Cover"
    ws["A1"].font = font(GREY_A1)
    span = list("CDEFGHIJ")
    for c in span:
        ws[f"{c}2"].font, ws[f"{c}2"].border = font(b=True), Border(top=THIN)
        ws[f"{c}2"].alignment = Alignment(horizontal="centerContinuous")
        ws[f"{c}3"].alignment = Alignment(horizontal="centerContinuous")
        ws[f"{c}8"].alignment = Alignment(horizontal="centerContinuous")
    ws["C2"] = entity
    ws["C3"], ws["C3"].font = units, font(i=True)
    ws["C8"], ws["C8"].font = subtitle, font(SUBTLE, i=True, size=12)
    ws.row_dimensions[2].height = 15

    def boxed(row, cols):
        for c in cols:
            ws[f"{c}{row}"].border = Border(top=THIN, bottom=THIN)
            ws[f"{c}{row}"].font = font(b=True)

    boxed(10, "CDEF")
    boxed(10, "HIJ")
    ws["C10"], ws["F10"] = "Last Updated", "=TODAY()"
    ws["F10"].number_format, ws["F10"].font = NF["date"], font()
    ws["F10"].alignment = Alignment(horizontal="center")
    ws["H10"], ws["J10"] = "Data As Of", data_as_of
    ws["J10"].font, ws["J10"].fill = font(BLUE, b=True), fill(INPUT)
    ws["J10"].alignment = Alignment(horizontal="centerContinuous")

    # color legend
    for c in "CDEF":
        ws[f"{c}12"].font, ws[f"{c}12"].border = font(b=True), Border(bottom=THIN)
    ws["C12"] = "Color Coding"
    legend = [("Formulas", font(INK), None), ("Hard-Code", font(BLUE), None),
              ("Link to Other Sheets", font(GREEN), None), ("Highlight Value (Check)", font(RED), None),
              ("Heading 1", font(WHITE, b=True), NAVY), ("Heading 2", font(WHITE, b=True), GREY_H2),
              ("Input / Driver", font(b=True), INPUT)]
    r = 13
    for text, f, fl in legend:
        ws[f"C{r}"], ws[f"C{r}"].font = text, f
        if fl:
            ws[f"C{r}"].fill = fill(fl)
            ws[f"C{r}"].alignment = Alignment(horizontal="center")
        r += 1

    # case selector
    if cases:
        boxed(21, "CDEF")
        ws["C21"] = case_label or "Active Case"
        ws["F21"] = cases[len(cases) // 2]
        ws["F21"].font, ws["F21"].fill = font(BLUE, b=True), fill(INPUT)
        ws["F21"].alignment = Alignment(horizontal="right")
        from openpyxl.worksheet.datavalidation import DataValidation
        dv = DataValidation(type="list", formula1='"' + ",".join(cases) + '"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add("F21")
        note(ws["F21"], "Select the active case; every tab reads it through its case row.")

    # sources
    r = 24
    for c in "CDEF":
        ws[f"C{r}"].font = font(b=True)
        ws[f"{c}{r}"].border = Border(bottom=THIN)
    ws[f"C{r}"] = "Sources"
    for s in sources:
        r += 1
        ws[f"C{r}"], ws[f"C{r}"].font = s, font()

    # table of contents (column H): headings bold-underlined, links show each tab's own title
    for c in "HIJ":
        ws[f"{c}12"].font, ws[f"{c}12"].border = font(b=True), Border(bottom=THIN)
        ws[f"{c}12"].alignment = Alignment(horizontal="centerContinuous")
    ws["H12"] = "Links (Ctrl + [ to Jump to Sheet)"
    r = 13
    for heading, tabs in sections:
        ws[f"H{r}"], ws[f"H{r}"].font = heading, font(b=True, u="single")
        r += 1
        for t in tabs:
            ws[f"H{r}"] = f"='{t}'!A1"
            ws[f"H{r}"].hyperlink = f"#'{t}'!A1"
            ws[f"H{r}"].font = font()
            ws[f"H{r}"].alignment = Alignment(horizontal="left", indent=1)
            r += 1
    return ws


def finalize(wb, zoom=85):
    """Delivery state: Garamond everywhere, gridlines off, every tab at A1 and the chosen zoom, cover active."""
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        if not ws.title.endswith(">"):
            ws.sheet_view.zoomScale = zoom
        ws.sheet_view.selection[0].activeCell = "A1"
        ws.sheet_view.selection[0].sqref = "A1"
        ws.sheet_view.topLeftCell = "A1"
        ws.sheet_view.tabSelected = False
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.active = 0
    wb.worksheets[0].sheet_view.tabSelected = True


# ======================================================================= specimen
def _selftest(path):
    wb = new_workbook()
    cover_page(wb, "Specimen Co.", "House Style Specimen", "6/30/26", sections=[("Specimen", ["Specimen"])],
               sources=["Company filings"], cases=["Bear", "Base", "Bull"])
    t = Tab(wb, "Specimen", "Specimen Co.", data_cols="D:I", tab_color=TAB_SECTIONS[0])
    t.header("Specimen, Every Row Type")
    t.case_row()
    t.band(8, "Income Statement", texts={"D": "Fiscal Year Ended December 31,", "G": "Projected"}, proj_from="G")
    t.year_header(9, 2023, 3)
    t.row(11, "Revenue", [100, 104, 108, "=F11*(1+G12)", "=G11*(1+H12)", "=H11*(1+I12)"], kind="hard",
          fmt="dol", label_kind="label_head")
    t.row(12, "% Growth", [None, "=E11/D11-1", "=F11/E11-1", 0.03, 0.03, 0.03], kind="calc_sub", fmt="pct")
    for c in "GHI":
        style(t.ws[f"{c}12"], "input", "pct")
    t.row(13, "Operating Expenses", lambda c: f"=-{c}11*{c}16", kind="calc")
    t.row(14, "EBITDA", lambda c: f"=SUM({c}11,{c}13)", kind="key")
    t.row(15, "% Margin", lambda c: f"={c}14/{c}11", kind="calc_sub", fmt="pct")
    t.row(16, "Opex % of revenue", [0.6] * 6, kind="input", fmt="pct")
    t.row(17, "Memo: Units", [200] * 6, kind="memo", fmt="num")
    t.check(18, "Check: EBITDA ties", lambda c: f"={c}14-({c}11+{c}13)")
    t.end(20)
    note(t.ws["C16"], "Illustrative driver: one input per cell, referenced by row 13.")
    finalize(wb)
    wb.save(path)
    return path


if __name__ == "__main__":
    import sys
    if len(sys.argv) == 3 and sys.argv[1] == "--selftest":
        print("wrote", _selftest(sys.argv[2]))
    else:
        print(__doc__)
