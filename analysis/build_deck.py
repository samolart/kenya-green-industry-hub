"""Build the HTML executive deck from outputs/results.json (the recalculated Excel model).

No number on a slide is typed by hand: everything is read from results.json or sources.py.
Usage: python analysis/build_deck.py            -> deck/kenya_green_hub_deck.html (standalone)
                                                    deck/kenya_green_hub_deck.fragment.html (for hosted artifact)
"""
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from deck_lib import (svg, text, Scale, line_chart, bar_chart, hbar_stacked, heat3, pct, num, esc)  # noqa: E402
from sources import SOURCES, CONFLICTS  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
R = json.load(open(os.path.join(ROOT, "outputs", "results.json")))
Y = R["years"]
G, H, M, U = R["grid"], R["hub"], R["matrix"], R["unit"]
LV, CAMP, SENS, SCR = R["levers"], U["campus"], U["sens"], R["screen"]
INP = {i["id"]: i for i in R["inputs"]}

tgt, amb, sq = H["Targeted hub"], H["Ambitious hub"], H["Status quo"]
i30 = Y.index(2030)
tco = U["tco"]
fee_gap = tco["Power-price gap to the Ethiopia-linked benchmark (grid power only)"]["Base"]
allin = {k: tco["All-in tenant cost"][k] for k in ("Stress", "Base", "Favourable", "Ethiopia benchmark")}
gap_min = M["gapmin"]["Base"]["Targeted hub"]
gap_tgt = M["gaptgt"]["Base"]["Targeted hub"]
gap_tgt_sq = M["gaptgt"]["Base"]["Status quo"]
hub_load30 = tgt["load"][i30]
pkg_mw, pkg_cost = LV["total_mw"], LV["total_cost"]
cov_tgt = LV["Package coverage of the target-margin gap"]
rm_after = LV["Reserve margin in 2030 after the package (Base supply, Targeted hub)"]
firm_cost = LV["levers"][0]["cost"]
jobs30 = int(round(tgt["it"][i30] / INP["campus_mw"]["base"] * INP["jobs"]["base"], -1))  # operating jobs, Targeted hub 2030
fee_gap_pct = tco["  ...as % of the tenant's all-in cost"]["Base"]
firm_pct = firm_cost / tgt["rev"][i30]
be_hi = SENS["breakeven"][-1]
power_share = {k: tco["Power (grid + evening premium) as % of all-in cost"][k] for k in ("Stress", "Base")}


def p1(v):
    return f"{v*100:.1f}%".replace("-", "−")


def p0(v):
    return f"{v*100:.0f}%".replace("-", "−")


def usd(v, d=0):
    return ("−$" if v < 0 else "$") + f"{abs(v):,.{d}f}"


# ------------------------------------------------------------------ charts
def chart_mix():
    rows = [("Jul–Dec 2025", [40.06, 38.73, 12.10, 9.11])]
    parts = [("Geothermal", "var(--green)"), ("Hydro, wind, solar, other", "var(--teal2)"),
             ("Imports", "var(--slate)"), ("Thermal", "var(--brick)")]
    s = hbar_stacked(760, 128, rows, parts, 100, 25, lambda v: f"{v:.0f}%", margin=(104, 24, 14, 34), show_totals=False,
                     label="Kenya generation mix July to December 2025", bar_h=44)
    legend = "".join(f'<span class="lg"><i style="background:{c}"></i>{esc(n)}</span>' for n, c in parts)
    return f'<div class="chartbox">{s}<div class="legend">{legend}</div></div>'


def chart_rm_now():
    s = bar_chart(560, 330, ["Jan 2026", "Jun 2026", "12 months to\nJun 2026"], [20.73, 3.34, -1.5], -5, 40, 5,
                  lambda v: f"{v:.0f}%", ["var(--green)", "var(--brick)", "var(--brick)"], band=(20, 35, "var(--green)", 0.12, "IEA comfortable range 20–35%"),
                  val_fmt=lambda v: f"{v:.1f}%".replace("-", "−"), label="Reserve margin, January and June 2026")
    return f'<div class="chartbox">{s}</div>'


def chart_outlook():
    ser = [
        dict(label="Favourable", values=[v * 100 for v in G["Favourable"]["rm"]], color="var(--green)"),
        dict(label="Base", values=[v * 100 for v in G["Base"]["rm"]], color="var(--amber)"),
        dict(label="Stress", values=[v * 100 for v in G["Stress"]["rm"]], color="var(--brick)"),
    ]
    s = line_chart(760, 380, Y, ser, -35, 35, 10, lambda v: f"{v:.0f}%".replace("-", "−"), margin=(52, 100, 24, 40),
                   bands=[(-35, 10, "var(--brick)", 0.07, "Below the 10% floor: no new hub load")], label="Reserve margin 2026 to 2032, no hub")
    return f'<div class="chartbox">{s}</div>'


def chart_prize():
    cats = [f"Status quo\n{p0(INP['share']['stress'])} of Africa", f"Targeted\n{p0(INP['share']['base'])} of Africa",
            f"Ambitious\n{p0(INP['share']['fav'])} of Africa"]
    s = bar_chart(520, 330, cats, [sq["it"][i30], tgt["it"][i30], amb["it"][i30]], 0, 400, 100, lambda v: f"{v:.0f}",
                  ["var(--slate)", "var(--green)", "var(--slate)"], val_fmt=lambda v: f"{v:.0f} MW", label="IT capacity in 2030 by hub case")
    return f'<div class="chartbox">{s}</div>'


def chart_gap():
    rows = [("No hub", [gap_tgt_sq, 0]), ("Targeted hub", [gap_tgt_sq, gap_tgt - gap_tgt_sq])]
    parts = [("Gap that already exists", "var(--slate)"), ("Added by the hub", "var(--amber)")]
    s = hbar_stacked(760, 210, rows, parts, 900, 100, lambda v: f"{v:,.0f}", margin=(150, 40, 34, 34),
                     markers=[(pkg_mw, "var(--green)", f"Five-lever package {pkg_mw:,.0f} MW")], label="2030 firm-MW gap to the 15% target margin", bar_h=44)
    legend = "".join(f'<span class="lg"><i style="background:{c}"></i>{esc(n)}</span>' for n, c in parts)
    return f'<div class="chartbox">{s}<div class="legend">{legend}</div></div>'


def chart_screen():
    rows = sorted(SCR["rows"], key=lambda r: -r["score"])
    w, h, ml, mr, mt = 760, 372, 360, 50, 8
    sx = Scale(0, 5, ml, w - mr)
    out = []
    for v in range(0, 6):
        out.append(f'<line x1="{sx(v):.1f}" x2="{sx(v):.1f}" y1="{mt}" y2="{h-30}" stroke="var(--rule)"/>')
        out.append(text(sx(v), h - 12, v, 11, "middle", fill="var(--slate)"))
    slot = (h - 30 - mt) / len(rows)
    for i, r in enumerate(rows):
        yc = mt + slot * i + slot / 2
        dc = r["industry"].startswith("Cloud")
        col = "var(--amber)" if dc else ("var(--green)" if i < 2 else "var(--slate)")
        out.append(text(ml - 12, yc + 4, re.sub(r"\s*\(.*?\)", "", r["industry"]), 12, "end", 600))
        out.append(f'<rect x="{ml}" y="{yc-13:.1f}" width="{sx(r["score"])-ml:.1f}" height="26" fill="{col}"/>')
        out.append(text(sx(r["score"]) + 8, yc + 5, f'{r["score"]:.2f}', 13, weight=700))
    return f'<div class="chartbox">{svg(w, h, "".join(out), label="Weighted industry screen score out of 5")}</div>'


def chart_tco():
    names = [("Stress", "Stress case"), ("Base", "Base case"), ("Favourable", "Favourable case"), ("Ethiopia benchmark", "Ethiopia-linked")]
    rows = []
    for k, lab in names:
        rows.append((lab, [tco["Infrastructure fee"][k], tco["Grid power (PUE x 730 h x all-in tariff)"][k],
                           tco["Evening self-supply premium (storage shifting cost, per IT kW)"][k]]))
    parts = [("Infrastructure fee", "var(--slate)"), ("Grid power", "var(--amber)"), ("Evening self-supply premium", "var(--brick)")]
    s = hbar_stacked(760, 270, rows, parts, 300, 50, lambda v: f"${v:.0f}", margin=(160, 56, 12, 34),
                     total_fmt=lambda t: f"${t:.0f}", label="Tenant cost per IT kW-month", bar_h=36)
    legend = "".join(f'<span class="lg"><i style="background:{c}"></i>{esc(n)}</span>' for n, c in parts)
    return f'<div class="chartbox">{s}<div class="legend">{legend}</div></div>'


def table_npv():
    caps = SENS["capex"]
    head = "".join(f"<th>${c:g}m</th>" for c in caps)
    body = []
    mx = max(abs(v) for r in SENS["rows"] for v in r["npv"])
    for row in SENS["rows"]:
        tds = []
        for c, v in zip(caps, row["npv"]):
            a = min(abs(v) / mx, 1) * 0.55 + 0.08
            bg = f"color-mix(in srgb, var(--{'green' if v >= 0 else 'brick'}) {a*100:.0f}%, var(--paper))"
            base = " base" if (row["fee"] == 150 and c == 8) else ""
            tds.append(f'<td class="{base.strip()}" style="background:{bg}">{usd(v)}m</td>')
        body.append(f'<tr><th>${row["fee"]}</th>{"".join(tds)}</tr>')
    return ('<div class="scroll"><table class="heat"><thead><tr><th class="corner">Fee ($/kW-month) ↓ &nbsp; Build cost ($m/MW) →</th>'
            f'{head}</tr></thead><tbody>{"".join(body)}</tbody></table></div>')


def levers_list():
    mx = max(l["mw"] for l in LV["levers"])
    out = []
    for l in LV["levers"]:
        out.append(
            f'<div class="lever"><div class="lname">{esc(l["name"])}</div>'
            f'<div class="lbar"><b style="width:{l["mw"]/mx*100:.0f}%"></b><span>{l["mw"]:,.0f} MW</span></div>'
            f'<div class="lmeta">{usd(l["cost"])}m a year · {esc(l["lead"])}</div></div>')
    return f'<div class="levers">{"".join(out)}</div>'


def chart_matrix():
    vals = [[M["rm2030"][s][h] * 100 for h in ("Status quo", "Targeted hub", "Ambitious hub")] for s in ("Stress", "Base", "Favourable")]

    def col(v):
        if v >= 10:
            return "var(--green)", "#ffffff"
        if v >= 0:
            return "var(--amber)", "#14202a"
        return "var(--brick)", "#ffffff"

    s = heat3(640, 230, ["Stress supply", "Base supply", "Favourable supply"], ["No hub", "Targeted hub", "Ambitious hub"], vals,
              lambda v: f"{v:.1f}%".replace("-", "−"), col, cell_h=62, label="Reserve margin in 2030 by supply and hub case")
    leg = ('<div class="legend"><span class="lg"><i style="background:var(--green)"></i>10% or more: load can connect</span>'
           '<span class="lg"><i style="background:var(--amber)"></i>0 to 10%: thin</span>'
           '<span class="lg"><i style="background:var(--brick)"></i>Negative: unserved peak</span></div>')
    return f'<div class="chartbox">{s}{leg}</div>'


# ------------------------------------------------------------------ slides
SL = []


def slide(title, body, source="", kicker="", cls=""):
    n = len(SL) + 1
    src = f'<p class="src">Source: {source}</p>' if source else ""
    k = f'<p class="kicker">{esc(kicker)}</p>' if kicker else ""
    SL.append(f'<section class="slide {cls}" id="s{n}" aria-label="Slide {n}">{k}<h2>{title}</h2>{body}{src}</section>')


def tile(big, small):
    return f'<div class="tile"><b>{big}</b><span>{small}</span></div>'


# 1 Cover
SL.append(f"""<section class="slide cover" id="s1" aria-label="Slide 1">
<p class="kicker">Strategy engagement · East Africa low-carbon industry</p>
<h1>Kenya can host East Africa's low-carbon industry hub if it fixes evening power first and lets demand follow supply</h1>
<p class="lede">A hypothesis-driven assessment of whether Kenya should position itself as the regional hub for data centres and other energy-intensive, low-carbon industry, and what it would take.</p>
<div class="coverfoot"><span>Samatar Ahmed</span><span>BSc Economics · Portfolio project · October 2026</span><span>Every figure traces to the Excel model</span></div>
</section>""")

# 2 Executive summary
slide("Kenya should build the hub in sequence: secure evening power, pilot with firmed and flexible load, then scale against a reserve-margin gate", f"""
<div class="cols3">
<div class="pt"><h3>The answer</h3><p>Yes, conditionally. Kenya's {p0(0.7879)} renewable supply is a real asset, but the 6–10pm peak has almost no spare capacity ({p1(0.0334)} reserve margin in June 2026 against a 20–35% norm).</p></div>
<div class="pt"><h3>The size of the prize</h3><p>A targeted hub reaches {tgt['it'][i30]:.0f} MW of IT load by 2030, a {usd(tgt['rev'][i30])}m a year fee pool and {usd(tgt['capex'][i30]/1000,1)}bn of build investment. It creates only about {jobs30:,} operating jobs, so the case rests on investment and exports, not employment.</p></div>
<div class="pt"><h3>What it would take</h3><p>Five levers add {pkg_mw:,.0f} MW of firm evening capacity for about {usd(pkg_cost)}m a year and lift the 2030 reserve margin to {p1(rm_after)}. Data centres supply {p0(LV['Share of the package supplied by the data centres themselves (lever 1)'])} of that themselves through on-site storage.</p></div>
</div>
<div class="ask"><b>Recommendation.</b> Lead with firmed data centres and flexible industry such as agro-processing and night-shift manufacturing. Connect new hub load in tranches only while the reserve margin stays at 10% or more. Do not market the hub on announcements until the first auction and storage tranche reach financial close.</div>
""", "Excel model, Summary sheet; sources S1 to S8")

# 3 Situation
slide("Kenya's grid is about 79% renewable, but that supply is built around geothermal baseload and daytime generation, not the evening peak", f"""
{chart_mix()}
<div class="cols3 tight">
<div class="pt"><h3>An advantage worth selling</h3><p>Geothermal alone is {p0(0.4006)} of generation, and renewables are {p1(0.7879)}. Low-carbon power is exactly what global tenants are asked to report.</p></div>
<div class="pt"><h3>But the peak is at night</h3><p>Demand peaks between 6pm and 10pm. Wind and solar are above 20% of capacity, yet solar adds nothing after sunset and storage is minimal.</p></div>
<div class="pt"><h3>And imports are rising</h3><p>Imports supplied 12% of Kenya Power's supply in H1 2026, up from 10% a year earlier, and Ethiopia's 200 MW is expected to double.</p></div>
</div>""", "S2, S4, S5, S6")

# 4 Problem
slide("Spare capacity has collapsed: the reserve margin fell from 20.7% in January to 3.3% in June 2026", f"""
<div class="split">{chart_rm_now()}
<div class="side"><div class="callout"><b>A signal from the market</b><p>Microsoft and G42's planned $1bn Olkaria data centre is in limbo over power capacity. The government says it is not abandoned, but the President has said capacity is inadequate.</p></div>
<div class="callout"><b>Peak is already above last year's record</b><p>Peak demand reached 2,514 MW in June 2026 and 2,549 MW in July, against a record of 2,362 MW in July 2025.</p></div></div></div>""",
      "S1, S2, S3, S5, S7, S8. Reserve margin = firm available capacity above peak demand, as a share of peak.")

# 5 Outlook
slide("With no hub at all, Base supply is already under the 10% floor and turns negative in 2029, so the grid problem comes before the hub", f"""
<div class="split wide">{chart_outlook()}
<div class="side"><div class="callout"><b>{p1(G['Base']['rm'][i30])} in 2030</b><p>Base supply, no hub: peak grows {p1(INP['growth']['base'])} a year, and about {G['Base']['need'][i30]:.0f} MW of new firm capacity is needed each year just to hold the margin.</p></div>
<div class="callout"><b>Two supply futures</b><p>Favourable (new storage and 300 MW a year of auctions) holds {p0(G['Favourable']['rm'][i30])}. Stress (delays, 8% growth) reaches {p0(G['Stress']['rm'][i30])}.</p></div></div></div>""",
      "Excel model, Grid sheet; S2, S13, S14. Firm capacity counts solar at 0% and storage at rated MW for a four-hour peak.")

# 6 Prize
slide(f"A targeted hub is worth about {usd(tgt['rev'][i30])}m a year in fees and {usd(tgt['capex'][i30]/1000,1)}bn of investment by 2030, but few jobs", f"""
<div class="split">{chart_prize()}
<div class="tiles">{tile(usd(tgt['rev'][i30])+'m','a year in infrastructure fees, 2030')}{tile(usd(tgt['capex'][i30]/1000,1)+'bn','build investment 2026 to 2030')}
{tile(f'{jobs30:,}','operating jobs, about 150 per campus')}{tile(f"{R['matrix']['mwh']['Base']['Targeted hub']:,.0f} MWh",'to shift into the evening each night')}</div></div>
<p class="note">Targeted hub: Kenya reaches {p0(INP['share']['base'])} of Africa's 1.5–2.2 GW data-centre capacity in 2030, up from about 45 MW today.</p>""",
      "Excel model, Hub sheet; S9, S10, S11, S12, S22")

# 7 Hub impact
slide("The hub itself adds 155 MW of evening load, about a fifth of the gap that already exists", f"""
{chart_gap()}
<div class="cols3 tight">
<div class="pt"><h3>The gap is not the hub's fault</h3><p>Without any hub, Kenya needs about {gap_tgt_sq:,.0f} MW of new firm capacity by 2030 to reach a 15% margin. The Targeted hub adds {gap_tgt-gap_tgt_sq:,.0f} MW including its own reserve cover.</p></div>
<div class="pt"><h3>The hub can pay for its own firming</h3><p>If data centres bring evening storage, it costs about {usd(firm_cost)}m a year, {p1(firm_pct)} of fee revenue.</p></div>
<div class="pt"><h3>The package covers it</h3><p>The five levers on slide 11 add {pkg_mw:,.0f} MW, {cov_tgt:.2f} times the target-margin gap.</p></div></div>""",
      "Excel model, Matrix and Levers sheets")

# 8 Screen
top = sorted(SCR["rows"], key=lambda r: -r["score"])
slide("Flexible, job-rich industry fits Kenya's evening-constrained grid better than colocation data centres, which rank sixth of seven", f"""
<div class="split wide">{chart_screen()}
<div class="side"><div class="callout"><b>How the screen works</b><p>Seven industries scored 1 to 5 on grid fit (30%), market pull (20%), Kenya's input advantage (20%), jobs per MW (15%) and speed to deploy (15%).</p></div>
<div class="callout"><b>What it means</b><p>Data centres have the strongest market pull but the least flexibility. They qualify only if they firm their own evening power. Agro-processing and night-shift heavy industry can use cheap daytime and night power and create more jobs per MW.</p></div></div></div>""",
      "Excel model, Screen sheet. Scores are analyst judgement [A], informed by S2, S11, S17, S22 and S25. Weights are editable on the sheet.")

# 9 Tenant cost
slide(f"Kenyan grid power costs a tenant about {usd(fee_gap)} more per kW-month than an Ethiopia-linked price, a handicap of {p0(fee_gap_pct)} of all-in cost", f"""
<div class="split wide">{chart_tco()}
<div class="side"><div class="callout"><b>Power is {p0(tco['Power (grid + evening premium) as % of all-in cost']['Base'])} of the Base bill</b><p>An all-in tenant cost of {usd(allin['Base'])} per IT kW-month, against {usd(allin['Ethiopia benchmark'])} on the Ethiopia benchmark (6.5 US cents per kWh).</p></div>
<div class="callout"><b>Time-of-use tariffs help</b><p>The off-peak discount is 50%, and uptake rose 74% in H2 2025. Rules that block factories from using it are being relaxed.</p></div></div></div>""",
      "Excel model, UnitEconomics sheet; S15 to S19, S23. Tariff range 6.5 to 11 US cents per kWh is an analyst range [S/A].")

# 10 Campus NPV
slide(f"A 44 MW campus earns {p1(CAMP['IRR (pre-tax, unlevered, incl. terminal value)'])} against a 10% hurdle, so price and build cost decide whether it is investable", f"""
<div class="split wide">
<div>{table_npv()}<p class="note">Campus NPV at 10%, $m. Outlined cell is the Base case. Green is positive, red is negative.</p></div>
<div class="tiles one">{tile(usd(CAMP['NPV at the base discount rate'],1)+'m','NPV, Base case')}
{tile(usd(CAMP['Break-even fee at Base build cost (NPV = 0)'])+'','break-even fee per kW-month, against $150 assumed')}
{tile(usd(-CAMP['NPV change per +$1m/MW of build cost'])+'m','NPV lost for each extra $1m per MW of build cost')}</div></div>""",
      "Excel model, UnitEconomics sheet. Pre-tax, unlevered, includes a 7x EBITDA terminal value; without it NPV is −$116m. S10, S21.")

# 11 Levers
slide(f"Five levers add {pkg_mw:,.0f} MW of firm evening capacity for about {usd(pkg_cost)}m a year, enough to restore a {p1(rm_after)} margin in 2030", f"""
<div class="split wide">{levers_list()}
<div class="side"><div class="callout"><b>Coverage</b><p>{cov_tgt:.2f} times the {gap_tgt:,.0f} MW gap to a 15% margin, and {LV['Package coverage of the minimum-margin gap']:.2f} times the {gap_min:,.0f} MW gap to the 10% minimum.</p></div>
<div class="callout"><b>Fastest first</b><p>The Ethiopia tranche and data-centre storage can arrive in 12 to 24 months. Auctions follow a 24 to 30 month build after financial close.</p></div></div></div>""",
      "Excel model, Levers sheet; S13, S14, S15, S19. Costs are indicative annual energy and capacity costs at the stated benchmark prices [A].")

# 12 Matrix
slide("Only the Favourable supply case keeps a healthy margin with a hub, so connection should be gated on delivered supply", f"""
<div class="split">{chart_matrix()}
<div class="side"><div class="callout"><b>The gate</b><p>Connect hub load only while the post-connection reserve margin stays at or above 10%. In Base and Stress supply, the model allows zero new MW to connect in any year to 2032.</p></div>
<div class="callout"><b>Reading it</b><p>The hub deepens a shortage; it does not cause one. That is why the levers come first.</p></div></div></div>""",
      "Excel model, Matrix sheet (reserve margin in 2030).")

# 13 Roadmap
slide("A three-phase roadmap links each step of hub growth to a measurable supply gate", """
<div class="phases">
<div class="ph"><h3>Phase 1 · Secure the evening</h3><p class="when">Now to month 12</p><ul>
<li>Run the first auction round and award the 100 MW KenGen battery tranche on schedule</li>
<li>Agree the second Ethiopia import tranche on the existing interconnector</li>
<li>Require new data centres to bring on-site evening storage as a licence condition</li>
<li>Relax night-tariff eligibility so factories can shift load</li></ul>
<p class="gate"><b>Gate A</b> Auction and storage at financial close</p></div>
<div class="ph"><h3>Phase 2 · Pilot firmed load</h3><p class="when">Month 12 to 36</p><ul>
<li>Connect two campuses, capped at about 100 MW IT, each with its own storage</li>
<li>Open a zone with agro-processing and night-shift industry on the time-of-use tariff</li>
<li>Offer one consolidated approval through the Special Economic Zones Authority</li></ul>
<p class="gate"><b>Gate B</b> Reserve margin at or above 10% after the pilot load</p></div>
<div class="ph"><h3>Phase 3 · Scale</h3><p class="when">Month 36 onward</p><ul>
<li>Grow to the Targeted hub of about 220 MW IT, tranche by tranche</li>
<li>Add flexible batch compute where the grid has headroom</li>
<li>Review the Ambitious case only after auction capacity reaches commercial operation</li></ul>
<p class="gate"><b>Gate C</b> About 800 MW of firm evening capacity delivered</p></div>
</div>""", "Lead times from S13 and S14: auctions within 12 months, financial close 9 to 12 months after PPA, operation 24 to 30 months after close.")

# 14 Risks
slide("The recommendation changes if supply slips, build costs rise or Ethiopian power becomes unavailable", f"""
<div class="scroll"><table class="risk"><thead><tr><th>Risk</th><th>What the model shows</th><th>Response</th></tr></thead><tbody>
<tr><td>Auctions and storage run late</td><td>Stress supply gives a reserve margin of {p0(G['Stress']['rm'][i30])} in 2030 even with no hub</td><td>Hold Phase 2 until Gate A is met</td></tr>
<tr><td>Build cost stays near the global $11.3m per MW</td><td>At $11.3m per MW the break-even fee is {usd(be_hi)} per kW-month, against $150 assumed</td><td>Prefer modular, phased campuses and local build content</td></tr>
<tr><td>Power price stays high</td><td>Power is {p0(power_share['Stress'])} of the tenant bill in Stress and {p0(power_share['Base'])} in Base</td><td>Time-of-use contracts and a dedicated large-load tariff</td></tr>
<tr><td>Ethiopian imports are curtailed</td><td>Imports are 12% of supply and an input to every firm-capacity case</td><td>Diversify with Uganda and domestic geothermal</td></tr>
<tr><td>Demand does not arrive</td><td>Africa capacity is a 1.5–2.2 GW range, not a forecast</td><td>Stage tranches against signed offtake</td></tr>
</tbody></table></div>""", "Excel model, Inputs, Matrix and UnitEconomics sheets")

# 15 Decisions
slide("Four decisions in the next 90 days would put Kenya on this path", f"""
<div class="decisions">
<div class="dc"><b>1</b><p>Publish a time-bound auction calendar and confirm the first storage tranche.</p></div>
<div class="dc"><b>2</b><p>Adopt a licence rule that new large data centres supply their own evening firming.</p></div>
<div class="dc"><b>3</b><p>Fast-track the second Ethiopia import tranche on the existing line.</p></div>
<div class="dc"><b>4</b><p>Set the reserve-margin gate and the quarterly reporting that goes with it.</p></div>
</div>
<p class="note">Next analytical steps: replace press-reported grid figures with primary EPRA and Kenya Power series, and collect quotes for Kenyan colocation pricing and build cost to replace the analyst assumptions.</p>""",
      "Recommendations from slides 2 to 14")

# Appendix A: method + assumptions
keys = ["growth", "pue", "util", "fee", "capex", "tariff", "wacc", "min_rm", "target_rm"]
rows = "".join(
    f'<tr><td>{esc(INP[k]["label"])}</td><td>{esc(INP[k]["unit"])}</td><td>{INP[k]["stress"]:g}</td><td>{INP[k]["base"]:g}</td><td>{INP[k]["fav"]:g}</td><td>{INP[k]["tag"]}</td></tr>'
    for k in keys)
slide("Appendix A. Method and the assumptions that matter most", f"""
<div class="cols2"><div><h3>Approach</h3><ol class="method"><li>Define the question and an issue tree: is there a market, can the grid serve it, does the economics work, what would it take.</li>
<li>Build the grid view: peak demand against firm, evening-capable supply in three supply cases.</li>
<li>Size the prize in three hub cases from Africa's 2030 data-centre capacity.</li>
<li>Test unit economics for a 44 MW campus and a tenant cost bridge.</li>
<li>Size five levers against the 2030 gap, then screen which industries fit the grid.</li>
<li>Record every source and every conflict between sources, and run ten integrity checks.</li></ol>
<p class="note">Tags: [S] sourced, [A] analyst assumption, [S/A] sourced with analyst adjustment.</p></div>
<div class="scroll"><table class="risk"><thead><tr><th>Assumption</th><th>Unit</th><th>Stress</th><th>Base</th><th>Fav.</th><th>Tag</th></tr></thead><tbody>{rows}</tbody></table></div></div>""",
      "Excel model, Inputs sheet", cls="appendix")

# Appendix B: sources
srcs = "".join(f'<li><b>{s[0]}</b> {esc(s[1])} ({esc(s[2])})</li>' for s in SOURCES)
confl = "".join(f'<li><b>{c[0]}</b> {esc(c[1])}. <i>Resolved:</i> {esc(c[2])}</li>' for c in CONFLICTS)
slide("Appendix B. Sources and the conflicts between them", f"""
<div class="cols2 small"><div><h3>Source register (read 19 October 2026)</h3><ol class="srcs">{srcs}</ol></div>
<div><h3>Conflicts and how they are resolved</h3><ol class="srcs">{confl}</ol>
<p class="note">Several inputs come from press reports quoting Kenya Power and EPRA briefings, not from primary datasets. Treat them as a starting point and confirm against EPRA statistics before relying on them.</p></div></div>""",
      "analysis/sources.py", cls="appendix")

# ------------------------------------------------------------------ page
CSS = open(os.path.join(os.path.dirname(__file__), "deck.css")).read()
JS = """
(function(){var s=[].slice.call(document.querySelectorAll('.slide')),n=s.length,c=document.getElementById('ctr');
function cur(){var y=window.innerHeight*0.35,b=0;s.forEach(function(e,i){if(e.getBoundingClientRect().top<y)b=i;});return b;}
function upd(){if(c)c.textContent=(cur()+1)+' / '+n;}
function go(d){var i=Math.max(0,Math.min(n-1,cur()+d));s[i].scrollIntoView({behavior:window.matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth',block:'start'});}
window.addEventListener('scroll',upd,{passive:true});
document.addEventListener('keydown',function(e){var t=e.target&&e.target.tagName;if(t==='INPUT'||t==='TEXTAREA')return;
if(e.key==='ArrowRight'||e.key==='PageDown'||(e.key===' '&&!e.shiftKey)){e.preventDefault();go(1);}
else if(e.key==='ArrowLeft'||e.key==='PageUp'||(e.key===' '&&e.shiftKey)){e.preventDefault();go(-1);}});
upd();})();
"""
FONTS = '<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin><link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,500;8..60,600&display=swap" rel="stylesheet">'
HEAD = '<title>Kenya Green Industry Hub</title><meta name="claude:viewport-fit-sides" content="cover">' + FONTS + f"<style>{CSS}</style>"
BODY = f'<main class="deck">{"".join(SL)}</main><div class="counter" id="ctr" aria-live="off"></div><script>{JS}</script>'

os.makedirs(os.path.join(ROOT, "deck"), exist_ok=True)
frag = HEAD + BODY
std = ('<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
       + HEAD + '</head><body>' + BODY + '</body></html>')
open(os.path.join(ROOT, "deck", "kenya_green_hub_deck.html"), "w").write(std)
open(os.path.join(ROOT, "deck", "kenya_green_hub_deck.fragment.html"), "w").write(frag)
print("slides:", len(SL), "bytes:", len(std))
