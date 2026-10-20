# Kenya as East Africa's low-carbon industry hub

A hypothesis-driven strategy assessment. **Question:** should Kenya position itself as East Africa's hub for energy-intensive, low-carbon industry (data centres first), and what would it take?

**Answer in one line.** Yes, but in sequence: Kenya's clean power is a real advantage, yet the grid has almost no spare evening capacity (3.3% reserve margin in June 2026). Secure evening supply first, pilot with firmed and flexible load, and scale only while the reserve margin stays at 10% or more.

## Headline findings

| Finding | Result |
|---|---|
| Clean-power advantage | About 79% renewable, 40% geothermal (Jul–Dec 2025) |
| Spare capacity | Reserve margin 20.7% (Jan 2026) to 3.3% (Jun 2026); Base supply negative from 2029 with no hub |
| Prize (targeted hub, 2030) | 222 MW IT, $300m a year in fees, $1.4bn investment, about 950 operating jobs |
| Hub's share of the grid gap | 155 MW of evening load, about a fifth of the 754 MW gap |
| Campus economics (44 MW, Base) | IRR 9.1% against 10% hurdle; NPV −$15.8m; break-even fee $156 per kW-month |
| Industry fit | Agro-processing (4.30) and night-shift heavy industry (4.10) beat colocation data centres (3.15, sixth of seven) |
| What it would take | Five levers add 805 MW of firm evening capacity for about $287m a year; 2030 reserve margin 16.5% |

## Deliverables

| File | What it is |
|---|---|
| [`deck/kenya_green_hub_deck.html`](deck/kenya_green_hub_deck.html) | 15-slide executive deck plus 2 appendix slides. One self-contained file; open in a browser, use the arrow keys. |
| [`model/kenya_green_hub_model.xlsx`](model/kenya_green_hub_model.xlsx) | 11-sheet scenario model with live formulas, source register and integrity checks |
| [`docs/executive_summary.md`](docs/executive_summary.md) | One-page executive summary |
| [`docs/methods_appendix.md`](docs/methods_appendix.md) | Issue tree, hypotheses, calculations, scenarios, limits |
| [`docs/cv_and_application_kit.md`](docs/cv_and_application_kit.md) | CV bullets and interview preparation |

## Repository layout

```
deck/        HTML deck (standalone) and hosted fragment
model/       Excel model
docs/        Executive summary, methods appendix, application kit
analysis/    build_model.py, export_results.py, build_deck.py, deck_lib.py, deck.css, sources.py
outputs/     results.json (exported from the recalculated model; feeds the deck)
```

## Reproduce

```bash
pip install openpyxl
python analysis/build_model.py model/kenya_green_hub_model.xlsx
# recalculate so formulas have cached values, for example with LibreOffice:
soffice --headless --convert-to xlsx --outdir model_recalc model/kenya_green_hub_model.xlsx
python analysis/export_results.py model/kenya_green_hub_model.xlsx outputs/results.json
python analysis/build_deck.py
```

`export_results.py` reads cached values, so the workbook must be recalculated first. The deck contains no hand-typed figures; it reads `outputs/results.json` and `analysis/sources.py`.

## Method in brief

Issue tree, five hypotheses, three supply scenarios × three hub scenarios, a tenant cost bridge, a 44 MW campus NPV, five sized levers and a weighted industry screen. Details are in [`docs/methods_appendix.md`](docs/methods_appendix.md).

## Limits

Several grid inputs come from press reports quoting Kenya Power and EPRA briefings, and colocation pricing, utilisation and cost of capital are analyst assumptions. Every input is tagged **[S]** sourced, **[A]** assumption or **[S/A]**, and six conflicts between sources are logged in the `Sources` sheet. This is an independent, self-initiated project using public sources, not a client engagement.

## Author

Samatar Ahmed · BSc Economics · October 2026
