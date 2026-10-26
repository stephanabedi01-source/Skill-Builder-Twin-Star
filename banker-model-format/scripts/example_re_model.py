"""Worked example: a small multifamily acquisition model in house style, built only with house_style.py.

Usage: python3 example_re_model.py out.xlsx

Every number here is ILLUSTRATIVE (a 240-unit garden asset); it shows layout and color-coding, not a deal.
Tabs: Cover Page | Inputs > | Assumptions | Unit Mix | Analysis > | Operating Pro Forma | Debt | Returns.
Patterns shown: case selector -> INDEX/MATCH active column, one input per cell, blue inputs on the input fill,
green links across tabs, a key-total NOI line, a debt schedule in the house layout, an exit block, IRR and
multiple, a formula-built sensitivity (no Excel data table needed), checks that turn red, notes in comments.
"""
import sys

from house_style import (NF, TAB_SECTIONS, Tab, cover_page, divider, finalize, new_workbook, note, style)

OUT = sys.argv[1] if len(sys.argv) > 1 else "example_re_model.xlsx"
ENTITY = "Atlas Apartments (Illustrative)"
YEARS = 10
wb = new_workbook()

# ================================================================= Assumptions
a = Tab(wb, "Assumptions", ENTITY, data_cols="D:G", label_width=44, tab_color=TAB_SECTIONS[0])
a.header("Acquisition Assumptions", units="All values in $ thousands, except per-unit and per-SF values")
a.case_row()
ws = a.ws

a.band(8, "Property and Purchase", texts={"D": "Input"})
a.col_header(9, {"D": "Value"})
a.row(11, "Units", [240], kind="input", fmt="num")
a.row(12, "Net rentable SF", [218880], kind="input", fmt="num")
a.row(13, "Purchase price", [52800], kind="input", fmt="dol")
a.row(14, "Price per unit ($)", ["=D13*1000/D11"], kind="calc_sub", fmt="dol")
a.row(15, "Going-in cap rate (Year 1 NOI)", ["='Operating Pro Forma'!E23/D13"], kind="calc_sub", fmt="pct2")
a.row(16, "Closing costs (% of price)", [0.015], kind="input", fmt="pct2")
a.row(17, "Total acquisition cost", ["=D13*(1+D16)"], kind="total", fmt="dol")
note(ws["C13"], "Illustrative. Replace with the contract price; source the PSA in this comment.")

a.band(19, "Operating Assumptions by Case", texts={"D": "Case Inputs", "G": "Selected"}, aside_cols=("G",))
a.col_header(20, {"D": "Downside", "E": "Base", "F": "Upside", "G": "Active"})
rows = [("Market rent growth (annual)", (0.020, 0.030, 0.035), "pct"),
        ("Economic vacancy", (0.070, 0.055, 0.050), "pct"),
        ("Expense growth (annual)", (0.035, 0.030, 0.030), "pct"),
        ("Exit cap rate", (0.0600, 0.0550, 0.0525), "pct2")]
for k, (lbl, vals, f) in enumerate(rows):
    r = 22 + k
    a.row(r, lbl, list(vals), kind="input", fmt=f)
    ws[f"G{r}"] = f'=INDEX(D{r}:F{r},MATCH($D$6,$D$20:$F$20,0))'
    style(ws[f"G{r}"], "calc", f, b=True)
note(ws["G20"], "Active column follows the case selected on the Cover Page; every other tab reads only this column.")

a.band(27, "Other Operating Inputs", texts={"D": "Input"})
a.col_header(28, {"D": "Value"})
a.row(30, "Other income per unit per month ($)", [85], kind="input", fmt="dol")
a.row(31, "Operating expenses per unit, Year 1 ($)", [6200], kind="input", fmt="dol")
a.row(32, "Management fee (% of EGI)", [0.030], kind="input", fmt="pct2")
a.row(33, "Replacement reserve per unit ($)", [300], kind="input", fmt="dol")
a.row(34, "Hold period (years)", [5], kind="input", fmt="num")
a.row(35, "Cost of sale (% of gross sale price)", [0.020], kind="input", fmt="pct2")
a.row(36, "Real estate taxes, Year 1 (% of price)", [0.012], kind="input", fmt="pct2")
a.row(37, "Real estate tax growth (annual)", [0.020], kind="input", fmt="pct")
note(ws["C36"], "Illustrative. Replace with the reassessed tax bill after sale.")
note(ws["C34"], "Display only in this example: the Returns tab is laid out for a 5-year hold.")

a.band(39, "Financing", texts={"D": "Input"})
a.col_header(40, {"D": "Value"})
a.row(42, "Loan-to-value", [0.60], kind="input", fmt="pct")
a.row(43, "Loan amount", ["=D13*D42"], kind="total", fmt="dol")
a.row(44, "Interest rate", [0.0575], kind="input", fmt="pct2")
a.row(45, "Interest-only period (years)", [2], kind="input", fmt="num")
a.row(46, "Amortization (years)", [30], kind="input", fmt="num")
a.row(47, "Financing fees (% of loan)", [0.010], kind="input", fmt="pct2")
a.row(48, "Equity required", ["=D17+D43*D47-D43"], kind="key", fmt="dol")
a.end(50)

# ================================================================= Unit Mix
u = Tab(wb, "Unit Mix", ENTITY, data_cols="D:I", label_width=26, tab_color=TAB_SECTIONS[0])
u.header("Unit Mix and In-Place Rents", units="Rents in $ per unit per month; SF per unit")
ws = u.ws
u.band(8, "Unit Mix", texts={"D": "As of Rent Roll Date"})
u.col_header(9, {"D": "Units", "E": "SF / Unit", "F": "In-Place", "G": "Market", "H": "Loss to", "I": "Total SF"},
             last=False)
u.col_header(10, {"D": "", "E": "", "F": "Rent", "G": "Rent", "H": "Lease", "I": ""})
mix = [("1 Bed / 1 Bath", 96, 760, 1450, 1525), ("2 Bed / 2 Bath", 120, 980, 1760, 1850),
       ("3 Bed / 2 Bath", 24, 1180, 2050, 2150)]
for k, (nm, n, sf, ip, mk) in enumerate(mix):
    r = 12 + k
    u.row(r, nm, [n, sf, ip, mk, f"=G{r}/F{r}-1", f"=D{r}*E{r}"], kind="hard", fmt="num")
    for c, f in (("F", "dol"), ("G", "dol")):
        ws[f"{c}{r}"].number_format = NF[f]
    ws[f"H{r}"].number_format = NF["pct"]
r = 15
u.row(r, "Total / Weighted Average", ["=SUM(D12:D14)", "=I15/D15", "=SUMPRODUCT(D12:D14,F12:F14)/D15",
                                      "=SUMPRODUCT(D12:D14,G12:G14)/D15", "=G15/F15-1", "=SUM(I12:I14)"],
      kind="total", fmt="num")
for c, f in (("F", "dol"), ("G", "dol"), ("H", "pct")):
    ws[f"{c}15"].number_format = NF[f]
u.check(16, "Check: ties to Assumptions",
        {"D": "=D15-Assumptions!D11", "I": "=I15-Assumptions!D12"})
note(ws["C12"], "Illustrative rent roll. In a live model, paste the rent roll on its own tab and SUMIFS by unit type.")
u.end(18)

# ================================================================= Operating Pro Forma
cols = [chr(ord("E") + k) for k in range(YEARS)]          # E..N = Year 1..10
p = Tab(wb, "Operating Pro Forma", ENTITY, data_cols=f"D:{cols[-1]}", label_width=40, tab_color=TAB_SECTIONS[1])
p.header("Operating Pro Forma, Annual", units="All values in $ thousands")
p.case_row()
ws = p.ws
p.band(8, "Operating Cash Flow", texts={"E": "Hold Period Year"})
p.col_header(9, {"D": "Per Unit, Yr 1"} | {c: k + 1 for k, c in enumerate(cols)})
for k, c in enumerate(cols):
    ws[f"{c}9"].number_format = '"Year "0'
A = "Assumptions"
UNITS = f"{A}!$D$11"


def prev(c):
    return chr(ord(c) - 1)


p.row(11, "Gross potential rent", {c: (f"='Unit Mix'!$G$15*{UNITS}*12/1000" if k == 0 else
                                      f"={prev(c)}11*(1+{A}!$G$22)") for k, c in enumerate(cols)},
      kind="calc", fmt="dol", label_kind="label_head")
p.row(12, "% Growth", {c: f"={c}11/{prev(c)}11-1" for c in cols[1:]}, kind="calc_sub", fmt="pct")
p.row(13, "(-) Vacancy and credit loss", {c: f"=-{c}11*{A}!$G$23" for c in cols}, kind="calc")
p.row(14, "(+) Other income", {c: (f"={A}!$D$30*{UNITS}*12/1000" if k == 0 else
                                   f"={prev(c)}14*(1+{A}!$G$22)") for k, c in enumerate(cols)}, kind="calc")
p.row(15, "Effective Gross Income", {c: f"=SUM({c}11,{c}13:{c}14)" for c in cols}, kind="total", fmt="dol")
p.row(17, "(-) Operating expenses", {c: (f"=-{A}!$D$31*{UNITS}/1000" if k == 0 else
                                         f"={prev(c)}17*(1+{A}!$G$24)") for k, c in enumerate(cols)}, kind="calc")
p.row(18, "(-) Management fee", {c: f"=-{c}15*{A}!$D$32" for c in cols}, kind="calc")
p.row(19, "(-) Real estate taxes", {c: (f"=-{A}!$D$13*{A}!$D$36" if k == 0 else f"={prev(c)}19*(1+{A}!$D$37)")
                                    for k, c in enumerate(cols)}, kind="calc")
p.row(20, "Total Operating Expenses", {c: f"=SUM({c}17:{c}19)" for c in cols}, kind="total", fmt="dol")
p.row(21, "% of EGI", {c: f"=-{c}20/{c}15" for c in cols}, kind="calc_sub", fmt="pct")
p.row(23, "Net Operating Income", {c: f"={c}15+{c}20" for c in cols}, kind="key", fmt="dol")
p.row(24, "% Growth", {c: f"={c}23/{prev(c)}23-1" for c in cols[1:]}, kind="calc_sub", fmt="pct")
p.row(26, "(-) Replacement reserves", {c: f"=-{A}!$D$33*{UNITS}/1000" for c in cols}, kind="calc")
p.row(27, "Cash Flow Before Debt Service", {c: f"={c}23+{c}26" for c in cols}, kind="total", fmt="dol")
for r in (11, 15, 20, 23, 27):                               # Year 1 per unit, in whole dollars
    ws[f"D{r}"] = f"={cols[0]}{r}*1000/{UNITS}"
    style(ws[f"D{r}"], "link", "dol", b=r in (15, 20, 23, 27))
p.end(29)

# ================================================================= Debt
d = Tab(wb, "Debt", ENTITY, data_cols=f"D:{cols[-1]}", label_width=40, tab_color=TAB_SECTIONS[1])
d.header("Senior Loan Schedule", units="All values in $ thousands")
ws = d.ws
d.band(8, "Senior Mortgage", texts={"E": "Hold Period Year"})
d.col_header(9, {"D": "Closing"} | {c: k + 1 for k, c in enumerate(cols)})
for c in cols:
    ws[f"{c}9"].number_format = '"Year "0'
d.row(11, "Beginning balance", {c: f"={prev(c)}13" for c in cols}, kind="calc", fmt="dol")
d.row(12, "(-) Scheduled amortization",
      {c: f"=IF({c}$9>{A}!$D$45,PPMT({A}!$D$44,{c}$9-{A}!$D$45,{A}!$D$46,{A}!$D$43),0)" for c in cols},
      kind="calc")
d.row(13, "Ending balance", {"D": f"={A}!D43"} | {c: f"={c}11+{c}12" for c in cols}, kind="total", fmt="dol")
note(ws["C12"], "Annual-pay approximation: interest-only for the IO period, then level payments over the "
                "amortization term. Switch to a monthly schedule if the term sheet requires it.")
d.row(15, "Interest expense", {c: f"=-{c}11*{A}!$D$44" for c in cols}, kind="calc")
d.row(16, "(-) Principal", {c: f"={c}12" for c in cols}, kind="calc")
d.row(17, "Total Debt Service", {c: f"={c}15+{c}16" for c in cols}, kind="key", fmt="dol")
d.band(19, "Credit Metrics", texts={"E": "Hold Period Year"}, markers=True)
d.row(21, "DSCR (NOI / debt service)", {c: f"=-'Operating Pro Forma'!{c}23/{c}17" for c in cols},
      kind="calc", fmt="mult2")
d.row(22, "Debt yield (NOI / beginning balance)", {c: f"='Operating Pro Forma'!{c}23/{c}11" for c in cols},
      kind="calc", fmt="pct")
d.end(24)

# ================================================================= Returns
hold_cols = cols[:5]
rcols = ["D"] + hold_cols                                 # D = Year 0
t = Tab(wb, "Returns", ENTITY, data_cols="D:K", label_width=40, tab_color=TAB_SECTIONS[2])
t.header("Returns Summary", units="All values in $ thousands")
t.case_row()
ws = t.ws
t.band(8, "Levered Cash Flow", texts={"D": "Hold Period Year"})
t.col_header(9, {c: k for k, c in enumerate(rcols)})
for c in rcols:
    ws[f"{c}9"].number_format = '"Year "0'
EXIT = hold_cols[-1]
PF = "'Operating Pro Forma'"
t.row(11, "(-) Equity contribution", {"D": f"=-{A}!D48"}, kind="link", fmt="dol")
t.row(12, "Cash flow before debt service", {c: f"={PF}!{c}27" for c in hold_cols}, kind="link")
t.row(13, "(-) Debt service", {c: f"=Debt!{c}17" for c in hold_cols}, kind="link")
t.row(14, "Gross sale price (forward NOI / exit cap)", {EXIT: f"={PF}!{chr(ord(EXIT) + 1)}23/{A}!$G$25"},
      kind="calc")
t.row(15, "(-) Cost of sale", {EXIT: f"=-{EXIT}14*{A}!$D$35"}, kind="calc")
t.row(16, "(-) Loan repayment", {EXIT: f"=-Debt!{EXIT}13"}, kind="link")
t.row(17, "Levered Cash Flow", {c: f"=SUM({c}11:{c}16)" for c in rcols}, kind="key", fmt="dol", cols=rcols)

t.band(19, "Returns", texts={"D": "Value"})
t.row(21, "Levered IRR", {"D": f"=IRR(D17:{EXIT}17)"}, kind="calc", fmt="pct", label_kind="label_head")
style(t.ws["D21"], "calc", "pct", b=True)
t.row(22, "Equity multiple", {"D": f"=SUM(E17:{EXIT}17)/-D17"}, kind="calc", fmt="mult2")
t.row(23, "Exit cap rate (active case)", {"D": f"={A}!G25"}, kind="link", fmt="pct2")

t.band(25, "Exit Cap Sensitivity", texts={"D": "Levered Cash Flow by Year", "J": "Returns"},
       aside_cols=("J", "K"))
t.col_header(26, {c: k for k, c in enumerate(rcols)} | {"J": "IRR", "K": "Multiple"}, label_cell="Exit cap rate")
for c in rcols:
    ws[f"{c}26"].number_format = '"Year "0'
for k, cap in enumerate((0.0500, 0.0525, 0.0550, 0.0575, 0.0600)):
    r = 28 + k
    flows = {c: f"={c}$17" for c in rcols[:-1]}
    flows[EXIT] = f"={EXIT}$12+{EXIT}$13+{PF}!${chr(ord(EXIT) + 1)}$23/$C{r}*(1-{A}!$D$35)+{EXIT}$16"
    flows["J"] = f"=IRR(D{r}:{EXIT}{r})"
    flows["K"] = f"=SUM(E{r}:{EXIT}{r})/-D{r}"
    t.row(r, cap, flows, kind="calc")
    style(ws[f"C{r}"], "hard", "pct2", align="left")    # grid axis values: blue, no driver fill
    ws[f"J{r}"].number_format, ws[f"K{r}"].number_format = NF["pct"], NF["mult2"]
note(ws["C26"], "Each row recomputes only the exit-year flow at its cap rate; Years 0 to 4 equal row 17. "
                "No Excel data table, so it recalculates everywhere (LibreOffice and Google Sheets included).")
t.check(34, "Check: base-case row matches the IRR above",
        {"D": "=INDEX(J28:J32,MATCH(D23,C28:C32,0))-D21"}, tol=0.0001)
t.end(36)

# ================================================================= cover, dividers, order
cover_page(wb, ENTITY, "Acquisition Model (Illustrative)", "Rent roll 9/30/26",
           sections=[("Inputs", ["Assumptions", "Unit Mix"]),
                     ("Analysis", ["Operating Pro Forma", "Debt", "Returns"])],
           sources=["Rent roll and T-12 (seller)", "Broker OM", "Lender term sheet"],
           case_label="Active Case", cases=["Downside", "Base", "Upside"])
divider(wb, "Inputs >", TAB_SECTIONS[0])
divider(wb, "Analysis >", TAB_SECTIONS[1])
order = ["Cover Page", "Inputs >", "Assumptions", "Unit Mix", "Analysis >", "Operating Pro Forma", "Debt", "Returns"]
wb._sheets = [wb[n] for n in order]
finalize(wb)
wb.save(OUT)
print("saved", OUT)
