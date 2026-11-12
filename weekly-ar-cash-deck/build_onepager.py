#!/usr/bin/env python3
"""One-page executive snapshot of the weekly AR, Cash & Payables deck (for the body of the cover email).

    python3 build_onepager.py --inputs inputs/weekly_inputs.xlsx --out output/JESA_AR_Cash_Snapshot_<Mon>_W<n>_<as-of>.pptx

Reads the same weekly_inputs.xlsx as build_deck.py and uses its drawing helpers. Amounts in MDH (= MMAD).
"""
import os, sys; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import argparse
import build_deck as B
from build_deck import (rect, tb, card, bar_chart, doughnut, color_points, f0, fd, datef, short, delta_color, read_inputs, Deck,
                        NAVY, NAVY2, BLUE, LBLUE, GOLD, TEAL, RED, GREEN, GREY, DARK, WHITE, SLATE, ORANGE, SW, SH,
                        RGBColor, MSO_SHAPE, MSO_ANCHOR, PP_ALIGN, XL_LEGEND_POSITION, XL_LABEL_POSITION)


def build(data, meta, out):
    k = {r["metric"]: r for r in data["AR_KPI"]}
    tot, unb, unp, adv = k["Total AR"], k["Unbilled"], k["Unpaid"], k["Advance Payment"]
    unrec = next(r for m, r in k.items() if m.startswith("Unrecognized"))
    def dw(r): return round(r["current"] or 0) - round(r["prev_week"] or 0)
    cur_s, prev_s = datef(meta["ar_current_date"], "%d-%b"), datef(meta["ar_prev_date"], "%d-%b")
    month = datef(meta["coll_month"] + "-15", "%B") if meta.get("coll_month") else datef(meta["ar_current_date"], "%B")
    dso = {r["bu"]: r for r in data.get("DSO", [])}.get("JESA Group", {})
    coll = data["COLLECTIONS_BU"]; c_m = sum(r["month_current"] or 0 for r in coll); c_prior = sum(r["prior_month"] or 0 for r in coll)
    sg = [r for r in data.get("AR_SIGNED_BU", []) if "split" in r]
    def S(rp=None, c=None, bu=None): return sum((r["current"] or 0) for r in sg if (rp is None or r["responsibility"] == rp) and (c is None or r["class"] == c) and (bu is None or r["bu"] == bu))
    signed, ns_c, ns_j = S(c="Signed"), S("Client", "Not signed"), S("JESA", "Not signed"); scope = signed + ns_c + ns_j
    cb = data["CASH_BU"]; net_act = sum(r["net_actuals"] or 0 for r in cb); net_appr = sum(r["net_after_approved"] or 0 for r in cb); net_all = sum(r["net_all_ongoing"] or 0 for r in cb)
    cash_in = sum(r["cash_in"] or 0 for r in cb); appr = -sum(r["oct_approved"] or 0 for r in cb); pend = -sum(r["oct_pending"] or 0 for r in cb)
    ap_tot = sum(r["amount"] or 0 for r in data["AP_BUCKETS"]); ap_over = sum(r["amount"] or 0 for r in data["AP_BUCKETS"] if r["bucket"] == "Overdue")
    lc_open = k["LC"]["current"] if "LC" in k else 0
    com = meta.get("cash_ongoing_month") or month[:3]

    D = Deck(meta)
    s = D.slide("AR, Cash & Payables – Weekly Snapshot", f"{meta.get('week_label', '')}   ·   AR position at {datef(meta['ar_current_date'])}   ·   amounts in MDH")
    D.entries[-1][2].text_frame.paragraphs[0].runs[0].text = ""   # single page: no page number

    # ---- KPI row
    cw = (SW - 0.6 - 0.12 * 5) / 6; y0 = 1.25; ch_ = 1.3
    def pill(x0, parts):
        rect(s, x0 + 0.1, y0 + ch_ - 0.43, cw - 0.2, 0.32, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
        tb(s, x0 + 0.1, y0 + ch_ - 0.43, cw - 0.2, 0.32, [parts], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    def wow(d): return [(f"vs {prev_s}  ", {"size": 8.5, "color": SLATE}), (f"{fd(d)}", {"size": 10.5, "bold": True, "color": delta_color(d)})]
    def cs(key): return sum((r.get(key) or 0) for r in coll)
    m_psr, m_proc, p_psr, p_proc = cs("month_labor"), cs("month_non_labor"), cs("prior_labor"), cs("prior_non_labor")
    import datetime as _dt
    _cm = _dt.date.fromisoformat((meta.get("coll_month") or str(meta["ar_current_date"])[:7]) + "-01"); prior_s = (_cm - _dt.timedelta(days=1)).strftime("%b")
    dep_m = float(meta.get("deposited_month") or 0); dep_p = float(meta.get("deposited_prior_labor") or 0) + float(meta.get("deposited_prior_non_labor") or 0)
    S85 = {"size": 8.5, "color": SLATE}; B105 = {"size": 10.5, "bold": True, "color": DARK}
    kp = [("Total AR", f0(tot["current"]), SLATE, wow(dw(tot)), None),
          ("Unbilled", f0(unb["current"]), TEAL, [(f"{unb['current']/tot['current']*100:.0f}% of AR  ", S85), (fd(dw(unb)), {"size": 10.5, "bold": True, "color": delta_color(dw(unb))})], None),
          ("Unpaid", f0(unp["current"]), GOLD, wow(dw(unp)), None),
          ("DSO (days)", (f"{int(float(dso['dso_current']))}" if dso.get("dso_current") not in (None, "") else "n/a"), NAVY2,
           [((meta.get("dso_prior_label") or "prior") + "  ", S85), (f"{int(float(dso['dso_prior_month']))}" if dso.get("dso_prior_month") not in (None, "") else "–", B105)], None),
          (f"Cash received – {month[:3]}", f0(c_m), GREEN, [(f"{prior_s} ", S85), (f0(c_prior), B105), (f"  PSR {f0(p_psr)} · Proc {f0(p_proc)}", {"size": 8, "color": SLATE})], f"PSR {f0(m_psr)}\nProc {f0(m_proc)}"),
          (f"Invoices billed – {month[:3]}", f0(dep_m), BLUE, [(f"{prior_s}  ", S85), (f0(dep_p), B105)], "deposited\nto date")]
    for i, (lab, val, col, parts, side) in enumerate(kp):
        x0 = 0.3 + i * (cw + 0.12)
        card(s, x0, y0, cw, ch_, lab, val, fill=col, value_size=26, label_size=9)
        if side: tb(s, x0 + cw - 1.05, y0 + 0.36, 0.95, 0.45, [[(t, {})] for t in side.split("\n")], size=9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT, line_spacing=1.0)
        pill(x0, parts)

    # ---- three panels
    py, ph = 2.75, 3.72
    def panel(x, w, title, sub=None):
        rect(s, x, py, w, ph, WHITE, line=LBLUE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        tb(s, x + 0.15, py + 0.08, w - 0.3, 0.3, title, size=12, bold=True, color=NAVY)
        if sub: tb(s, x + 0.15, py + 0.38, w - 0.3, 0.25, sub, size=8.5, color=SLATE)
    # A. signature status of the unbilled roll-ups
    xa, wa = 0.3, 4.05
    panel(xa, wa, "Unbilled roll-ups – signature status", f"{f0(scope)} MDH of unbilled in scope")
    doughnut(s, xa + 0.05, py + 0.75, 2.2, 2.2, ["Signed", "Not signed – client", "Not signed – JESA"], [signed, ns_c, ns_j], [GREEN, GOLD, ORANGE], hole=58)
    tb(s, xa + 0.05, py + 1.52, 2.2, 0.4, f"{signed/(scope or 1)*100:.0f}%", size=20, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    tb(s, xa + 0.05, py + 1.9, 2.2, 0.25, "signed", size=9, color=SLATE, align=PP_ALIGN.CENTER)
    for i, (lab, v, col) in enumerate((("Signed by the client", signed, GREEN), ("Not signed – client", ns_c, GOLD), ("Not signed – JESA", ns_j, ORANGE))):
        yy = py + 0.95 + i * 0.66
        rect(s, xa + 2.32, yy + 0.05, 0.16, 0.16, col)
        tb(s, xa + 2.54, yy - 0.03, 1.5, 0.26, lab, size=8.5, bold=True, color=NAVY)
        tb(s, xa + 2.54, yy + 0.2, 1.5, 0.3, [[(f0(v), {"size": 13, "bold": True, "color": DARK}), (f"   {v/(scope or 1)*100:.0f}%", {"size": 9, "color": SLATE})]])
    tb(s, xa + 0.15, py + ph - 0.62, wa - 0.3, 0.55, [[("Not signed: ", {"bold": True, "color": NAVY}), (f"{f0(ns_c + ns_j)} MDH", {"bold": True, "color": RED}), (f" – {f0(ns_c)} external (client) and {f0(ns_j)} internal (JESA).", {})]], size=9, color=DARK, line_spacing=1.05)
    # B. unrecognized revenue by status
    xb, wb_ = 4.5, 4.45
    ut = data.get("UNRECOGNIZED", [])
    def U(key): return sum((r.get("eq_" + key) or 0) + (r.get("sub_" + key) or 0) for r in ut)
    u_tot = unrec["current"] or 0; u_sub, u_app, u_sig, u_cn, u_risk = U("pending_submission"), U("pending_approval"), U("signed_by_client"), U("under_cn"), U("at_risk")
    panel(xb, wb_, f"Unrecognized revenue – {f0(u_tot)} MDH", f"Executed scope not yet booked as AR, by status (MDH); {fd(dw(unrec))} vs {prev_s}")
    ucat = [("Pending JESA submission", u_sub, ORANGE), ("Pending client approval", u_app, GOLD), ("Signed by the client", u_sig, GREEN), ("Under change notice", u_cn, SLATE), ("At risk", u_risk, RED)]
    chb = bar_chart(s, xb + 0.05, py + 0.62, wb_ - 0.1, 2.1, [c[0] for c in ucat], [("MDH", [c[1] for c in ucat])], [NAVY], horizontal=True, legend=False, gap=45, cat_size=9, label_size=9.5)
    color_points(chb, 0, [c[2] for c in ucat])
    att = data.get("UR_ATTACHMENT_STATUS", [])
    can = sum((r.get("total") or 0) for r in att) if att else (u_app + u_sig)   # attachments that can be signed, same basis as the deck's key message
    top = sorted([r for r in ut if (r.get("grand_total") or 0) > 0.5], key=lambda r: -(r.get("grand_total") or 0))[:3]
    tb(s, xb + 0.18, py + 2.78, wb_ - 0.36, 0.85, [[("Signed by the client: ", {"bold": True, "color": NAVY}), (f"{f0(u_sig)} MDH", {"bold": True, "color": GREEN}), (f" of {f0(can)} MDH that can be signed ({u_sig/(can or 1)*100:.0f}%).", {})],
                                                    [("Largest: ", {"bold": True, "color": NAVY}), ("  ·  ".join(f"{short(r['bu'])} {f0(r['grand_total'])}" for r in top), {})]], size=9, color=DARK, line_spacing=1.25)
    # C. procurement cash position
    xc, wc = 9.1, 3.93
    panel(xc, wc, "Cash position (Procurement Back to Back)", f"Cumulative since July, after {com} commitments (MDH)")
    vals = [net_act, net_appr, net_all]
    chc = bar_chart(s, xc + 0.05, py + 0.6, wc - 0.1, 2.05, ["On actuals", f"After approved\n{com} payments", "After pending\npayments"], [("MDH", vals)], [NAVY], legend=False, gap=55, cat_size=8, label_size=10,
                    number_format='+#,##0;−#,##0;0', label_pos=XL_LABEL_POSITION.OUTSIDE_END)
    color_points(chc, 0, [GREEN if v >= 0 else RED for v in vals])
    chc.category_axis.tick_label_position = B.XL_TICK_LABEL_POSITION.LOW
    lo, hi = min(vals + [0]), max(vals + [0]); span = (hi - lo) or 1
    chc.value_axis.minimum_scale = lo - span * 0.28; chc.value_axis.maximum_scale = hi + span * 0.22   # room for the value labels
    rows = [("Cash IN", f0(cash_in)), (f"{com} approved payments", f0(appr)), (f"{com} pending payments", f0(pend)), ("Supplier payables", f"{f0(ap_tot)}  ·  {f0(ap_over)} past due")]
    for i, (lab, v) in enumerate(rows):
        yy = py + 2.72 + i * 0.235
        tb(s, xc + 0.18, yy, 1.9, 0.24, lab, size=8.5, color=SLATE)
        tb(s, xc + 1.7, yy, wc - 1.88, 0.24, v, size=9, bold=True, color=DARK, align=PP_ALIGN.RIGHT)

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
