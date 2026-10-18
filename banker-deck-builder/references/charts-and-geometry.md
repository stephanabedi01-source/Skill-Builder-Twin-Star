# Native charts with full control (python-pptx + chart XML)

The approach that produced every chart in the Project Print deck: create a real, editable PowerPoint chart (data in
its embedded workbook), then replace its XML with a hand-built plot whose inner plot area fills the frame exactly
and whose axes are hidden. Because the mapping from value to inches is then linear and known, every label,
bracket, band, claim line and pill can be placed precisely with ordinary shapes.

## Contents
1. The pattern
2. Recipes (column, horizontal bar, floating/football field, waterfall, combo, percent-stacked, outline segments)
3. Placing labels and annotations
4. Gotchas

## 1. The pattern

```python
P = Plot(x, y, w, h, n=len(cats), vmin=0, vmax=10.2)             # columns
xs = [P.center(k) for k in range(n)]; yv = P.v                     # category centers, value -> y
chart(slide, P, cats, [("EV / LTM EBITDA", vals)],
      [bar_ser(0, "EV / LTM EBITDA", cats, vals, "grey_lt", dpt={9: "accent_fill"})], "A deals chart", gap=60)
line(slide, P.x, yv(0), P.x + P.w, yv(0), "grey_mid", 0.75)        # your own baseline
for k, v in enumerate(vals):                                        # your own labels
    textbox(slide, xs[k] - 0.4, yv(v) - 0.22, 0.8, 0.2, [[(f"{v:.1f}x", {"b": True})]], size=10.5,
            align=PP_ALIGN.CENTER)
```
- `chart_space(..., inner=(0,0,1,1))` + `cat_ax(delete=True)` + `val_ax(delete=True)` = plot fills the frame.
- For horizontal bars use `Plot(..., horizontal=True)`; `kind="bar"` reverses the category axis so the first
  category is on top. `P.v(value)` then returns x, `P.center(k)` returns y.
- Bar thickness = `P.bar_width(gap)` (slot / (1 + gap/100)); use it to place in-bar labels and break marks.
- For custom bodies (combos, percent-stacked) call `add_chart(slide, x, y, w, h, cats, series, body_fn, name)`
  with your own XML; the series list must match the XML series order (it fills the embedded workbook).

## 2. Recipes

- **Plain columns/bars**: `bar_chart_xml([bar_ser(...)], kind, "clustered", gap)`. Per-point colors with `dpt`.
- **Floating bars / football field**: grouping "stacked": `nofill_ser` for the low value (invisible base) +
  `bar_ser` for (high - low). Reference rows via `dpt` in grey_lt. Leave a blank category (None values) to create a
  gap row for a "Reference points" sub-header.
- **Waterfall / bridge**: stacked: invisible base + increases (grey_mid or ink) + decreases (accent or grey_mid) +
  totals (ink; first/last via `dpt`). Compute bases cumulatively; assert the run ends at the reported total.
- **Combo (bars + line)**: one `<c:barChart>` and one `<c:lineChart>` sharing the same axis ids; `line_ser` with
  white-filled markers. For a line on a different scale, add a second val/cat axis pair (hidden) and give the line
  chart those ids.
- **Outline segment** (e.g. "at face" extension, unsecured claims at risk): a second stacked series whose point has
  `dpt={k: (None, "accent", "dash")}`: no fill, dashed accent outline.
- **Claim vs. recovery overlay**: two clustered series with overlap 100: recovery filled, claim noFill with a
  grey_mid outline drawn on top.
- **Percent-stacked splits**: grouping "percentStacked", one category per row, one series per segment with `dpt`
  colors per row. White 0.75 pt segment borders separate the pieces cleanly.
- **Sensitivity grids**: shapes, not charts: a grid of `box(RECTANGLE)` cells with the base case dark and the rest
  `band`; header row and WACC column as textboxes. Use the model's own grid values so it ties out.
- **Single reference points** (consensus, a rating-agency figure): draw as shapes (small diamond/circle + label)
  rather than one-point series, which render unreliably.

## 3. Placing labels and annotations

- Value labels: above bars at `yv(v) - 0.22`; inside bar tops at `yv(v) + 0.05` (white on dark bars) when a line,
  band or median would otherwise cross them; right of horizontal bars at `P.v(v) + 0.06`; inside a horizontal bar
  at `P.v(lo) + 0.05` when a claim line sits just left of it.
- Category labels: textboxes centered on `P.center(k)`, two or three lines (name bold / detail grey / tertiary),
  slightly wider than the slot with wrap off if a name is long.
- Claim lines: dashed vertical lines from just above the plot to just below; label centered at the top above the
  plot; move the plot down if the label touches the panel header rule.
- Bands (ranges): rectangles behind the chart (add them BEFORE the chart so the bars draw on top), extending a
  little beyond the plot; label in the caption row or as a legend swatch.
- Medians and reference lines that must not cross labels: draw them before the chart (behind bars) and put labels
  inside bars.
- Brackets: `hbracket` under/over the plot spanning two values; `brace` beside a pair of bars spanning the gap.
- Pills tied to a point: place near the point; connect with a thin line if not adjacent.
- Timeline: equal spacing; solid line for the planned path, dashed accent segment to the deadline; dates above,
  two-line step text below.

## 4. Gotchas

- Escape `&`, `<`, `>` in any text that goes into chart XML (series names, categories): "PP&E" breaks the parse.
  deck_kit's `str_cache` does this; raw XML you write yourself must too.
- Every series needs an `<c:invertIfNegative val="0"/>` and data labels deleted unless used, or PowerPoint adds
  defaults.
- Charts inherit theme fonts; set `txPr` with the deck font (deck_kit does) or labels come out in Aptos/Calibri.
- LibreOffice renders charts close to PowerPoint but not identically; hollow markers and some outlines may not
  show in the render while present in the file.
- Do not use XL chart types other than the placeholder COLUMN_CLUSTERED in `add_chart`; the XML replacement
  defines the real type.
- After building, run the pptx skill's `validate.py --original <base>`; it checks chart parts and relationships.
