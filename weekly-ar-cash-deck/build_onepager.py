#!/usr/bin/env python3
"""One-page executive snapshot of the weekly AR, Cash & Payables deck (for the body of the cover email).

    python3 build_onepager.py --inputs inputs/weekly_inputs.xlsx --out output/JESA_AR_Cash_Snapshot_<Mon>_W<n>_<as-of>.pptx

Reads the same weekly_inputs.xlsx as build_deck.py and uses its drawing helpers. Amounts in MDH (= MMAD).
"""
import os, sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import argparse, datetime as dt
import build_deck as B
from build_deck import (rect, tb, card, bar_chart, doughnut, f0, fd, datef, delta_color, read_inputs, Deck,
                        NAVY, NAVY2, BLUE, LBLUE, GOLD, TEAL, RED, GREEN, DARK, WHITE, SLATE, ORANGE, SW,
                        RGBColor, MSO_SHAPE, MSO_ANCHOR, PP_ALIGN, XL_LEGEND_POSITION, XL_LABEL_POSITION, Inches, Pt)


def cash_waterfall(slide, x, y, w, h, steps, cat_size=8, label_size=9):
    """steps: (label, value, kind, colour); kind 'total' = bar from zero, 'delta' = move from the running balance.
    Bars may sit above, below or across zero. Value labels sit inside large bars and just outside small ones."""
    lo_hi = []; run = 0.0
    for _, v, kind, _ in steps:
        if kind == "total": a, b = 0.0, v; run = v
        else: a, b = run, run + v; run = run + v
        lo_hi.append((min(a, b), max(a, b)))
    span = (max(hh for _, hh in lo_hi) - min(ll for ll, _ in lo_hi)) or 1.0; car = span * 0.13
    ser = {k: [] for k in ("bp", "vp", "cp", "bn", "vn", "cn")}; place = []
    for (ll, hh), (_, v, kind, _) in zip(lo_hi, steps):
        pos = max(hh, 0) - max(ll, 0); neg = min(ll, 0) - min(hh, 0)          # visible parts above / below zero
        ser["bp"].append(max(ll, 0)); ser["vp"].append(pos); ser["bn"].append(min(hh, 0)); ser["vn"].append(neg)
        side = "p" if pos >= -neg else "n"; big = max(pos, -neg) >= span * 0.17
        ser["cp"].append(car if (side == "p" and not big) else 0.0); ser["cn"].append(-car if (side == "n" and not big) else 0.0)
        place.append((side, big))
    cd = B.CategoryChartData(); cd.categories = [st[0] for st in steps]
    for kname in ("bp", "vp", "cp", "bn", "vn", "cn"): cd.add_series(kname, ser[kname])
    gf = slide.shapes.add_chart(B.XL_CHART_TYPE.COLUMN_STACKED, Inches(x), Inches(y), Inches(w), Inches(h), cd); ch = gf.chart
    ch.font.name = B.FONT; ch.has_legend = False
    plot = ch.plots[0]; plot.gap_width = 35; plot.overlap = 100
    S = dict(zip(("bp", "vp", "cp", "bn", "vn", "cn"), plot.series))
    for kname in ("bp", "cp", "bn", "cn"): S[kname].format.fill.background(); S[kname].format.line.fill.background()
    for kname in ("vp", "vn"): S[kname].format.line.fill.background()
    for i, ((_, v, kind, col), (side, big)) in enumerate(zip(steps, place)):
        for kname in ("vp", "vn"):
            pt = S[kname].points[i]; pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
        tgt = S[("v" if big else "c") + side].points[i]
        tgt.data_label.has_text_frame = True; tgt.data_label.text_frame.text = (f0(v) if (kind == "total" and v >= 0) else fd(v))
        r = tgt.data_label.text_frame.paragraphs[0].runs[0]; r.font.size = Pt(label_size); r.font.bold = True; r.font.name = B.FONT
        r.font.color.rgb = WHITE if big else DARK
        tgt.data_label.position = XL_LABEL_POSITION.CENTER if big else XL_LABEL_POSITION.INSIDE_BASE
    B._style_axes(ch, cat_size=cat_size)
    ch.category_axis.tick_label_position = B.XL_TICK_LABEL_POSITION.LOW
    return ch


def build(data, meta, out):
    k = {r["metric"]: r for r in data["AR_KPI"]}
    tot, unb, unp = k["Total AR"], k["Unbilled"], k["Unpaid"]
    unrec = next(r for m, r in k.items() if m.startswith("Unrecognized")); gross = k["Total NEW AR"]
    def dw(r): return round(r["current"] or 0) - round(r["prev_week"] or 0)
    prev_s = datef(meta["ar_prev_date"], "%d-%b")
    month = datef(meta["coll_month"] + "-15", "%B") if meta.get("coll_month") else datef(meta["ar_current_date"], "%B")
    _cm = dt.date.fromisoformat((meta.get("coll_month") or str(meta["ar_current_date"])[:7]) + "-01"); prior_s = (_cm - dt.timedelta(days=1)).strftime("%b")
    dso = {r["bu"]: r for r in data.get("DSO", [])}.get("JESA Group", {})
    coll = data["COLLECTIONS_BU"]
    def cs(key): return sum((r.get(key) or 0) for r in coll)
    c_m, c_prior = cs("month_current"), cs("prior_month")
    m_psr, m_proc, p_psr, p_proc = cs("month_labor"), cs("month_non_labor"), cs("prior_labor"), cs("prior_non_labor")
    dep_m = float(meta.get("deposited_month") or 0); dep_p = float(meta.get("deposited_prior_labor") or 0) + float(meta.get("deposited_prior_non_labor") or 0)
    sg = [r for r in data.get("AR_SIGNED_BU", []) if "split" in r]
    def S(rp=None, c=None): return sum((r["current"] or 0) for r in sg if (rp is None or r["responsibility"] == rp) and (c is None or r["class"] == c))
    signed, ns_c, ns_j = S(c="Signed"), S("Client", "Not signed"), S("JESA", "Not signed"); scope = signed + ns_c + ns_j
    cb = data["CASH_BU"]
    def cbs(key): return sum((r.get(key) or 0) for r in cb)
    cash_in = cbs("cash_in"); paid = cbs("cash_out_suppliers") + cbs("cash_out_lc") + cbs("cash_out_customs")
    net_appr, net_all = cbs("net_after_approved"), cbs("net_all_ongoing"); appr_all = cbs("oct_approved"); pend = cbs("oct_pending")
    ong = {r["bucket"]: (r["amount"] or 0) for r in data.get("ONGOING", [])}
    lc_due = -(ong.get("Approved – October LC maturities") or 0); appr_sup = appr_all - lc_due
    ap_tot = sum(r["amount"] or 0 for r in data["AP_BUCKETS"]); ap_over = sum(r["amount"] or 0 for r in data["AP_BUCKETS"] if r["bucket"] == "Overdue")
    com = meta.get("cash_ongoing_month") or month[:3]

    D = Deck(meta)
    s = D.slide("AR, Cash & Payables – Weekly Snapshot", f"{meta.get('week_label', '')}   ·   AR position at {datef(meta['ar_current_date'])}   ·   amounts in MDH")
    D.entries[-1][2].text_frame.paragraphs[0].runs[0].text = ""   # single page: no page number

    # ---- KPI row: gross exposure, unrecognized, total AR, unbilled, unpaid, DSO, cash received (wider), invoices billed
    y0, ch_, gap, wn, wc_ = 1.25, 1.3, 0.1, 1.4, 2.2
    S8 = {"size": 8, "color": SLATE}; B10 = {"size": 10, "bold": True, "color": DARK}
    def wow(d): return [(f"vs {prev_s}  ", S8), (fd(d), {"size": 10, "bold": True, "color": delta_color(d)})]
    kp = [("Gross exposure", f0(gross["current"]), RED, wow(dw(gross)), wn, None),
          ("Unrecognized", f0(unrec["current"]), BLUE, wow(dw(unrec)), wn, None),
          ("Total AR", f0(tot["current"]), SLATE, wow(dw(tot)), wn, None),
          ("Unbilled", f0(unb["current"]), TEAL, [(f"{unb['current']/tot['current']*100:.0f}% of AR  ", S8), (fd(dw(unb)), {"size": 10, "bold": True, "color": delta_color(dw(unb))})], wn, None),
          ("Unpaid", f0(unp["current"]), GOLD, wow(dw(unp)), wn, None),
          ("DSO (days)", (f"{int(float(dso['dso_current']))}" if dso.get("dso_current") not in (None, "") else "n/a"), NAVY2,
           [((meta.get("dso_prior_label") or "prior") + "  ", S8), (f"{int(float(dso['dso_prior_month']))}" if dso.get("dso_prior_month") not in (None, "") else "–", B10)], wn, None),
          (f"Cash received – {month[:3]}", f0(c_m), GREEN, [(f"{prior_s}  ", S8), (f0(c_prior), B10), (f"   PSR {f0(p_psr)} · Proc {f0(p_proc)}", S8)], wc_, f"PSR {f0(m_psr)}\nProc {f0(m_proc)}"),
          (f"Billed – {month[:3]}", f0(dep_m), NAVY, [(f"{prior_s}  ", S8), (f0(dep_p), B10)], wn, None)]
    x0 = 0.3
    for lab, val, col, parts, w_, side in kp:
        card(s, x0, y0, w_, ch_, lab, val, fill=col, value_size=22, label_size=8)
        if side: tb(s, x0 + w_ - 1.05, y0 + 0.36, 0.95, 0.45, [[(t, {})] for t in side.split("\n")], size=9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, line_spacing=1.0)
        rect(s, x0 + 0.08, y0 + ch_ - 0.43, w_ - 0.16, 0.32, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
        tb(s, x0 + 0.08, y0 + ch_ - 0.43, w_ - 0.16, 0.32, [parts], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x0 += w_ + gap

    # ---- three panels
    py, ph = 2.75, 3.72
    def panel(x, w, title, sub=None):
        rect(s, x, py, w, ph, WHITE, line=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        tb(s, x + 0.15, py + 0.08, w - 0.3, 0.3, title, size=12, bold=True, color=NAVY)
        if sub: tb(s, x + 0.15, py + 0.38, w - 0.3, 0.25, sub, size=8.5, color=SLATE)
    # A. signature status of the unbilled roll-ups
    xa, wa = 0.3, 3.6
    panel(xa, wa, "Unbilled roll-ups – signature status", f"{f0(scope)} MDH of unbilled in scope")
    doughnut(s, xa + 0.02, py + 0.8, 1.95, 1.95, ["Signed", "Not signed – client", "Not signed – JESA"], [signed, ns_c, ns_j], [GREEN, GOLD, ORANGE], hole=58)
    tb(s, xa + 0.02, py + 1.47, 1.95, 0.4, f"{signed/(scope or 1)*100:.0f}%", size=18, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    tb(s, xa + 0.02, py + 1.82, 1.95, 0.25, "signed", size=9, color=SLATE, align=PP_ALIGN.CENTER)
    for i, (lab, v, col) in enumerate((("Signed by the client", signed, GREEN), ("Not signed – client", ns_c, GOLD), ("Not signed – JESA", ns_j, ORANGE))):
        yy = py + 0.92 + i * 0.64
        rect(s, xa + 2.03, yy + 0.05, 0.15, 0.15, col)
        tb(s, xa + 2.22, yy - 0.03, 1.35, 0.26, lab, size=8.5, bold=True, color=NAVY)
        tb(s, xa + 2.22, yy + 0.2, 1.35, 0.3, [[(f0(v), {"size": 13, "bold": True, "color": DARK}), (f"   {v/(scope or 1)*100:.0f}%", {"size": 9, "color": SLATE})]])
    tb(s, xa + 0.15, py + ph - 0.62, wa - 0.3, 0.55, [[("Not signed: ", {"bold": True, "color": NAVY}), (f"{f0(ns_c + ns_j)} MDH", {"bold": True, "color": RED}), (f" – {f0(ns_c)} external (client) and {f0(ns_j)} internal (JESA).", {})]], size=9, color=DARK, line_spacing=1.05)

    # B. unrecognized revenue: equipment vs services, signed vs not signed by the client
    xb, wb_ = xa + wa + 0.12, 3.9
    ut = data.get("UNRECOGNIZED", [])
    EQ = ("eq_pending_approval", "eq_signed_by_client", "eq_pending_submission", "eq_under_cn", "eq_at_risk")
    u_tot = sum((r.get("grand_total") or 0) for r in ut); eq_tot = sum((r.get(c) or 0) for r in ut for c in EQ); sv_tot = u_tot - eq_tot
    eq_sig = sum((r.get("eq_signed_by_client") or 0) for r in ut); sv_sig = sum((r.get("sub_signed_by_client") or 0) for r in ut)
    panel(xb, wb_, f"Unrecognized revenue – {f0(u_tot)} MDH", f"6% JESA cost + fee markup included; {fd(dw(unrec))} vs {prev_s}")
    hw = (wb_ - 0.3 - 0.12) / 2
    for i, (lab, v, col) in enumerate((("Equipment", eq_tot, NAVY2), ("Services", sv_tot, GOLD))):
        bx = xb + 0.15 + i * (hw + 0.12)
        rect(s, bx, py + 0.72, hw, 0.78, col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
        tb(s, bx + 0.1, py + 0.75, hw - 0.2, 0.24, lab.upper(), size=7.5, bold=True, color=WHITE)
        tb(s, bx + 0.1, py + 0.96, hw - 0.2, 0.45, [[(f0(v), {"size": 19, "bold": True, "color": WHITE}), (f"   {v/(u_tot or 1)*100:.0f}%", {"size": 9.5, "color": WHITE})]])
    cats = ["Equipment", "Services", "Total"]; sig = [eq_sig, sv_sig, eq_sig + sv_sig]; nsg = [eq_tot - eq_sig, sv_tot - sv_sig, u_tot - eq_sig - sv_sig]
    chb = bar_chart(s, xb + 0.05, py + 1.56, wb_ - 0.1, 1.5, cats, [("Signed by the client", sig), ("Not signed", nsg)], [GREEN, ORANGE], horizontal=True, stacked=True, gap=40,
                    cat_size=8.5, label_size=8.5, label_colors=[DARK, DARK], legend_pos=XL_LEGEND_POSITION.BOTTOM)
    B.hide_small_labels(chb, [sig, nsg], u_tot * 0.04)
    att = data.get("UR_ATTACHMENT_STATUS", [])
    can = sum((r.get("total") or 0) for r in att) if att else None   # attachments that can be signed, same basis as the deck's key message
    tb(s, xb + 0.15, py + 3.1, wb_ - 0.3, 0.26, [[("Signed by the client: ", {"bold": True, "color": NAVY}), (f"{f0(eq_sig + sv_sig)} MDH", {"bold": True, "color": GREEN}), (f" (equipment {f0(eq_sig)}, services {f0(sv_sig)}).", {})]], size=8.5, color=DARK)
    tb(s, xb + 0.15, py + 3.37, wb_ - 0.3, 0.26, "No revenue recognition booked on these balances to date.", size=8.5, bold=True, color=RED)

    # C. cash position: waterfall from cash IN to the balance after approved, then after not-approved payments
    xc = xb + wb_ + 0.12; wc = SW - 0.3 - xc
    panel(xc, wc, "Cash position (Procurement Back to Back)", f"Cumulative since July, {com} commitments included (MDH)")
    steps = [("Cash IN", cash_in, "total", GREEN), ("Paid", paid, "delta", NAVY2), (f"Approved\nsuppliers", appr_sup, "delta", GOLD), ("LC due", lc_due, "delta", GOLD),
             ("Balance after\napproved", net_appr, "total", RED if net_appr < 0 else GREEN), ("Not\napproved", pend, "delta", ORANGE), ("Balance after\nall", net_all, "total", RED if net_all < 0 else GREEN)]
    cash_waterfall(s, xc + 0.05, py + 0.62, wc - 0.1, 2.72, steps)
    tb(s, xc + 0.18, py + ph - 0.34, 2.2, 0.26, "Supplier payables", size=8.5, color=SLATE)
    tb(s, xc + 1.6, py + ph - 0.34, wc - 1.78, 0.26, f"{f0(ap_tot)}  ·  {f0(ap_over)} past due", size=9, bold=True, color=DARK, align=PP_ALIGN.RIGHT)

    # ---- bottom bands: context note and priority
    yb = py + ph + 0.12
    pc, pp = meta.get("posted_cur"), meta.get("posted_prev")
    pd_ = (round(float(pc)) - round(float(pp))) if (pc not in (None, "") and pp not in (None, "")) else None
    has_note = pd_ is not None and pd_ >= 50 and dw(tot) > 0
    wl = 6.3 if has_note else 0
    if has_note:
        rect(s, 0.3, yb, wl, 0.5, RGBColor(0xFD, 0xF3, 0xDC), line=GOLD, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
        tb(s, 0.42, yb, wl - 0.24, 0.5, [[("Note:  ", {"bold": True, "color": NAVY}), (f"AR {fd(dw(tot))} MDH vs {prev_s}, mainly the {meta.get('posted_period')} revenue posted ", {}), (f"({fd(pd_)} MDH)", {"bold": True, "color": NAVY}),
                                          (f"; excluding it, AR moves {fd(dw(tot) - pd_)} MDH.", {})]], size=9.5, color=DARK, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    xp = 0.3 + (wl + 0.12 if has_note else 0)
    rect(s, xp, yb, SW - 0.3 - xp, 0.5, NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
    tb(s, xp + 0.12, yb, SW - 0.3 - xp - 0.24, 0.5, [[("Priority:  ", {"bold": True, "color": GOLD}), ("secure the client's sign-off on completed work (roll-ups / attachments) and clear the internal backlog, to unlock cash for PSR and critical suppliers.", {"color": WHITE})]],
       size=9.5, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.0)
    D.save(out)
    print("saved", out)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("--inputs", default="inputs/weekly_inputs.xlsx"); ap.add_argument("--out", required=True)
    a = ap.parse_args()
    data, meta = read_inputs(a.inputs)
    build(data, meta, a.out)
