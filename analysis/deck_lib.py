"""Small SVG chart helpers for the HTML deck. No external libraries: the deck is one self-contained file.

Colours come from CSS variables so the charts follow the deck theme.
"""
from html import escape

GREEN = "var(--green)"
AMBER = "var(--amber)"
BRICK = "var(--brick)"
SLATE = "var(--slate)"
INK = "var(--ink)"
RULE = "var(--rule)"
MIST = "var(--mist)"


def esc(t):
    return escape(str(t), quote=True)


def svg(w, h, inner, cls="chart", label=""):
    return (f'<svg class="{cls}" viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" aria-label="{esc(label)}" '
            f'xmlns="http://www.w3.org/2000/svg">{inner}</svg>')


def text(x, y, s, size=12, anchor="start", weight=400, fill=INK, extra=""):
    return f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" text-anchor="{anchor}" font-weight="{weight}" fill="{fill}" {extra}>{esc(s)}</text>'


class Scale:
    def __init__(self, d0, d1, r0, r1):
        self.d0, self.d1, self.r0, self.r1 = d0, d1, r0, r1

    def __call__(self, v):
        return self.r0 + (v - self.d0) / (self.d1 - self.d0) * (self.r1 - self.r0)


def line_chart(w, h, xs, series, ymin, ymax, ystep, yfmt, margin=(52, 24, 28, 44), bands=(), vlines=(), fills=(), labels_end=True,
               ylabel="", label="", hlines=()):
    """series: list of dict(label, values, color, dash=None, width=2.5, end_label=True, dots=True)
    bands: list of (y0, y1, color, opacity, text)
    hlines: list of (y, color, text, dash)
    fills: list of (values_a, values_b, color, opacity) shade between a and b where a<b (deficit) .
    """
    ml, mr, mt, mb = margin
    pw, ph = w - ml - mr, h - mt - mb
    sx = Scale(0, len(xs) - 1, ml, ml + pw)
    sy = Scale(ymin, ymax, mt + ph, mt)
    out = []
    for (y0, y1, col, op, txt) in bands:
        out.append(f'<rect x="{ml}" y="{sy(y1):.1f}" width="{pw}" height="{sy(y0)-sy(y1):.1f}" fill="{col}" opacity="{op}"/>')
        if txt:
            out.append(text(ml + 6, sy(y0) - 8, txt, 11, fill=SLATE))
    v = ymin
    while v <= ymax + 1e-9:
        y = sy(v)
        out.append(f'<line x1="{ml}" x2="{ml+pw}" y1="{y:.1f}" y2="{y:.1f}" stroke="{RULE}" stroke-width="1"/>')
        out.append(text(ml - 8, y + 4, yfmt(v), 11, "end", fill=SLATE))
        v += ystep
    for i, xv in enumerate(xs):
        out.append(text(sx(i), mt + ph + 20, xv, 12, "middle", fill=SLATE))
    for (y, col, txt, dash) in hlines:
        out.append(f'<line x1="{ml}" x2="{ml+pw}" y1="{sy(y):.1f}" y2="{sy(y):.1f}" stroke="{col}" stroke-width="1.5" stroke-dasharray="{dash or "0"}"/>')
        if txt:
            out.append(text(ml + pw, sy(y) - 5, txt, 11, "end", fill=col, weight=600))
    for (a, b, col, op) in fills:
        # split into segments where a<b
        seg = []
        for i in range(len(a)):
            if a[i] < b[i]:
                seg.append(i)
            else:
                if seg:
                    out.append(_fill_poly(seg, a, b, sx, sy, col, op))
                    seg = []
        if seg:
            out.append(_fill_poly(seg, a, b, sx, sy, col, op))
    for xi, txt in vlines:
        out.append(f'<line x1="{sx(xi):.1f}" x2="{sx(xi):.1f}" y1="{mt}" y2="{mt+ph}" stroke="{SLATE}" stroke-width="1" stroke-dasharray="3 3"/>')
        if txt:
            out.append(text(sx(xi) + 5, mt + 12, txt, 11, fill=SLATE))
    for s in series:
        pts = " ".join(f"{sx(i):.1f},{sy(v):.1f}" for i, v in enumerate(s["values"]))
        out.append(f'<polyline points="{pts}" fill="none" stroke="{s["color"]}" stroke-width="{s.get("width", 2.5)}" '
                   f'stroke-dasharray="{s.get("dash") or "0"}" stroke-linejoin="round" stroke-linecap="round"/>')
        if s.get("dots", True):
            for i, v in enumerate(s["values"]):
                out.append(f'<circle cx="{sx(i):.1f}" cy="{sy(v):.1f}" r="3.2" fill="{s["color"]}"/>')
        if labels_end and s.get("end_label", True):
            out.append(text(sx(len(xs) - 1) + 8, sy(s["values"][-1]) + 4, s["label"], 12, weight=600, fill=s["color"]))
    if ylabel:
        out.append(text(ml, 12, ylabel, 11, fill=SLATE))
    return svg(w, h, "".join(out), label=label)


def _fill_poly(seg, a, b, sx, sy, col, op):
    top = [(sx(i), sy(b[i])) for i in seg]
    bot = [(sx(i), sy(a[i])) for i in reversed(seg)]
    pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in top + bot)
    return f'<polygon points="{pts}" fill="{col}" opacity="{op}"/>'


def bar_chart(w, h, cats, values, ymin, ymax, ystep, yfmt, colors, margin=(52, 20, 22, 46), val_fmt=None, band=None, label="", bar_w=0.58,
              hlines=(), cat_size=12, val_size=13):
    ml, mr, mt, mb = margin
    pw, ph = w - ml - mr, h - mt - mb
    sy = Scale(ymin, ymax, mt + ph, mt)
    n = len(cats)
    slot = pw / n
    out = []
    if band:
        y0, y1, col, op, txt = band
        out.append(f'<rect x="{ml}" y="{sy(y1):.1f}" width="{pw}" height="{sy(y0)-sy(y1):.1f}" fill="{col}" opacity="{op}"/>')
        if txt:
            out.append(text(ml + pw - 6, sy(y1) + 14, txt, 11, "end", fill=SLATE))
    v = ymin
    while v <= ymax + 1e-9:
        y = sy(v)
        out.append(f'<line x1="{ml}" x2="{ml+pw}" y1="{y:.1f}" y2="{y:.1f}" stroke="{RULE}" stroke-width="1"/>')
        out.append(text(ml - 8, y + 4, yfmt(v), 11, "end", fill=SLATE))
        v += ystep
    zero = sy(0) if ymin < 0 < ymax else sy(ymin)
    out.append(f'<line x1="{ml}" x2="{ml+pw}" y1="{zero:.1f}" y2="{zero:.1f}" stroke="{INK}" stroke-width="1.2"/>')
    for i, (c, val) in enumerate(zip(cats, values)):
        bw = slot * bar_w
        x = ml + slot * i + (slot - bw) / 2
        y = sy(val)
        top, hgt = (y, zero - y) if val >= 0 else (zero, y - zero)
        col = colors[i] if isinstance(colors, (list, tuple)) else colors
        out.append(f'<rect x="{x:.1f}" y="{top:.1f}" width="{bw:.1f}" height="{max(hgt,0.5):.1f}" fill="{col}"/>')
        lab = (val_fmt or yfmt)(val)
        ly = top - 6 if val >= 0 else top + hgt + 16
        out.append(text(x + bw / 2, ly, lab, val_size, "middle", 700, INK))
        for k, line in enumerate(str(c).split("\n")):
            out.append(text(x + bw / 2, mt + ph + 18 + 15 * k, line, cat_size, "middle", fill=SLATE))
    for (y, col, txt, dash) in hlines:
        out.append(f'<line x1="{ml}" x2="{ml+pw}" y1="{sy(y):.1f}" y2="{sy(y):.1f}" stroke="{col}" stroke-width="1.5" stroke-dasharray="{dash or "0"}"/>')
        if txt:
            out.append(text(ml + pw, sy(y) - 5, txt, 11, "end", col, 600))
    return svg(w, h, "".join(out), label=label)


def hbar_stacked(w, h, rows, parts, xmax, xstep, xfmt, margin=(150, 24, 12, 34), label="", markers=(), show_totals=True, total_fmt=None, bar_h=34):
    """rows: list of (row label, [values per part]); parts: list of (name, color)."""
    ml, mr, mt, mb = margin
    pw, ph = w - ml - mr, h - mt - mb
    sx = Scale(0, xmax, ml, ml + pw)
    n = len(rows)
    slot = ph / n
    out = []
    v = 0
    while v <= xmax + 1e-9:
        x = sx(v)
        out.append(f'<line x1="{x:.1f}" x2="{x:.1f}" y1="{mt}" y2="{mt+ph}" stroke="{RULE}" stroke-width="1"/>')
        out.append(text(x, mt + ph + 18, xfmt(v), 11, "middle", fill=SLATE))
        v += xstep
    for i, (name, vals) in enumerate(rows):
        yc = mt + slot * i + slot / 2
        out.append(text(ml - 10, yc + 4, name, 12, "end", 600, INK))
        acc = 0
        for (pn, col), val in zip(parts, vals):
            x0, x1 = sx(acc), sx(acc + val)
            out.append(f'<rect x="{x0:.1f}" y="{yc-bar_h/2:.1f}" width="{max(x1-x0,0):.1f}" height="{bar_h}" fill="{col}"/>')
            if x1 - x0 > 34:
                out.append(text((x0 + x1) / 2, yc + 4, f"{val:,.0f}", 11.5, "middle", 600, "#fff"))
            acc += val
        if show_totals:
            out.append(text(sx(acc) + 8, yc + 5, (total_fmt or (lambda t: f"{t:,.0f}"))(acc), 13, weight=700))
    for (xv, col, txt) in markers:
        x = sx(xv)
        out.append(f'<line x1="{x:.1f}" x2="{x:.1f}" y1="{mt-2}" y2="{mt+ph}" stroke="{col}" stroke-width="2" stroke-dasharray="5 3"/>')
        out.append(text(x, mt - 4 if mt > 14 else 10, txt, 11, "middle", 600, col))
    return svg(w, h, "".join(out), label=label)


def heat3(w, h, row_labels, col_labels, vals, fmt, color_fn, cell_h=64, label=""):
    """3x3 heat map. vals[r][c]."""
    ml, mt = 160, 40
    cw = (w - ml - 8) / len(col_labels)
    out = []
    for j, cl in enumerate(col_labels):
        out.append(text(ml + cw * j + cw / 2, 24, cl, 12.5, "middle", 600))
    for i, rl in enumerate(row_labels):
        y = mt + cell_h * i
        out.append(text(ml - 10, y + cell_h / 2 + 4, rl, 12, "end", 600))
        for j, _ in enumerate(col_labels):
            v = vals[i][j]
            bg, fg = color_fn(v)
            out.append(f'<rect x="{ml+cw*j+2:.1f}" y="{y+2}" width="{cw-4:.1f}" height="{cell_h-4}" fill="{bg}"/>')
            out.append(text(ml + cw * j + cw / 2, y + cell_h / 2 + 7, fmt(v), 20, "middle", 700, fg))
    return svg(w, h, "".join(out), label=label)


def pct(v, d=1):
    s = f"{v*100:.{d}f}%"
    return s.replace("-", "−")


def num(v, d=0):
    s = f"{v:,.{d}f}"
    return s.replace("-", "−")
