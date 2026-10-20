# Methods appendix

This appendix lets a reader follow the reasoning without a call. Every number in the deck is read from `outputs/results.json`, which is exported from the recalculated Excel model, so the deck cannot drift from the model.

## 1. Question and issue tree

**Key question.** Should Kenya position itself as East Africa's hub for energy-intensive, low-carbon industry (data centres first), and what would it take?

```
Should Kenya be the hub?
├── A. Is there a real advantage?        -> clean power mix, regional position (Mombasa cables)
├── B. Can the grid serve it?            -> peak demand vs firm evening supply, 2026-2032
├── C. Is the market big enough?         -> hub size, fees, investment, jobs in 2030
├── D. Does the economics work?          -> tenant cost vs Ethiopia benchmark; campus NPV, IRR
├── E. Which industries fit best?        -> weighted screen of seven industries
└── F. What would it take?               -> five levers, lead times, phased roadmap with gates
```

## 2. Hypotheses and verdicts

| # | Hypothesis | Verdict | Evidence |
|---|---|---|---|
| H1 | Kenya has a structural low-carbon advantage | Supported on carbon, not on capacity or price | Renewables about 79% of generation, geothermal 40%; but power costs $26 per kW-month more than an Ethiopia-linked price |
| H2 | The grid can serve a hub now | Rejected | Reserve margin 3.3% in June 2026; Base supply is negative from 2029; zero connectable MW in Base and Stress |
| H3 | Hub economics are attractive | Marginal | Campus IRR 9.1% against a 10% hurdle; break-even fee $156 against $150 |
| H4 | Data centres are the best anchor sector | Not supported | Rank sixth of seven in the weighted screen |
| H5 | The gap can be closed in time | Supported, with conditions | Five levers add 805 MW, 1.07 times the 2030 gap, but depend on auction delivery |

## 3. Data and evidence

- 28 cited sources (S1 to S28) are listed in the `Sources` sheet and `analysis/sources.py`, each with the figure taken and the date read (19 October 2026).
- Six conflicts between sources (C1 to C6) are recorded with how each is resolved. Example: installed capacity is quoted as 3,192 MW (EPRA), about 3,000 MW and 2,300 MW in different places; the model uses Kenya Power's peak-demand and reserve-margin series instead.
- Each input on the `Inputs` sheet is tagged **[S]** sourced, **[A]** analyst assumption, or **[S/A]** sourced with analyst adjustment, with its source ID or rationale.

## 4. Model architecture (11 sheets)

| Sheet | Purpose |
|---|---|
| README | How to read the workbook and its limits |
| Summary | Headline outputs, all linked |
| Inputs | Every assumption in three columns (Stress, Base, Favourable) |
| Grid | Peak demand against firm, evening-capable supply, 2026 to 2032 |
| Hub | Data-centre capacity, load, fees, investment, jobs in three hub cases |
| Matrix | Reserve margin, firm-MW gap and connectable load for all nine supply × hub combinations |
| UnitEconomics | Tenant cost per kW-month; 44 MW reference campus cash flow, NPV, IRR, break-even, sensitivities |
| Levers | Five levers sized against the 2030 gap |
| Screen | Weighted industry screen |
| Sources | Source register and conflict log |
| Checks | Ten integrity checks, all TRUE |

The workbook has about 1,080 live formulas and no formula errors. Blue cells are inputs; change them and everything recalculates.

## 5. Key calculations

- **Reserve margin** = (firm available capacity − peak demand) ÷ peak demand.
- **Firm capacity** counts solar at 0% (the peak is 6 to 10pm) and storage at its rated MW for a four-hour peak. Baseline firm capacity is set so the June 2026 margin reproduces the sourced figure (3.34% Base, −1.5% Stress, 20.73% Favourable).
- **Peak demand** grows from 2,514 MW (June 2026) at 6.5% a year in Base (5% Favourable, 8% Stress).
- **Hub IT capacity in 2030** = Kenya's share of Africa's 2030 capacity (4%, 12%, 20%) × 1,850 MW, the midpoint of McKinsey's 1.5 to 2.2 GW range.
- **Hub load** = incremental IT MW × utilisation (75%) × PUE (1.4).
- **Fee pool** = IT MW × utilisation × infrastructure fee × 12.
- **Firm MW gap** = peak × (1 + target margin) + hub load × (1 + target margin) − firm capacity, with a 10% minimum and a 15% target.
- **Tenant cost** per IT kW-month = infrastructure fee + PUE × 730 hours × all-in tariff + evening self-supply premium (battery shifting cost from a BNEF benchmark with a 1.5× emerging-market premium, 85% round-trip efficiency).
- **Campus case**: 44 MW in two 22 MW phases, pre-tax, unlevered, discounted at 10%, with a 7× EBITDA terminal value. A linear decomposition check confirms the NPV sensitivities.
- **Industry screen**: scores 1 to 5 on five criteria with weights 30/20/20/15/15.

## 6. Scenarios

| Dimension | Stress | Base | Favourable |
|---|---|---|---|
| Reserve margin, June 2026 | −1.5% | 3.34% | 20.73% |
| Peak growth | 8% | 6.5% | 5% |
| Ethiopia import uplift | +100 MW from 2028 | +200 MW from 2027 | +200 MW from 2027 |
| Auction capacity | none | 150 MW a year from 2030 | 300 MW a year from 2030, plus 100 MW extra storage |
| All-in tariff | 11 US c/kWh | 9 US c/kWh | 6.5 US c/kWh |
| Build cost | $11.3m/MW | $8m/MW | $7m/MW |
| Infrastructure fee | $110 | $150 | $200 per kW-month |

## 7. Limits

1. Several grid inputs are press reports quoting Kenya Power and EPRA briefings. Confirm them against EPRA statistics before publishing.
2. Colocation pricing, utilisation, operating cost and the discount rate are analyst assumptions because no public Kenyan price data was available.
3. A negative reserve margin sizes the response. It does not forecast blackouts, because emergency capacity could be procured.
4. The campus case shows what price or cost of capital a campus needs. It does not value a specific project and ignores tax and financing.
5. The industry screen is judgement. Its value is in making the weights explicit and changeable.
6. Africa's 2030 data-centre capacity is a range from a consultancy report, not a forecast, and Kenya's share is a scenario input.

## 8. How I checked the work

- Ten integrity checks in the workbook, including that the Base margin reproduces the sourced 3.34%, scenarios are correctly ordered, and the campus NPV decomposition reproduces NPV.
- No formula errors across all sheets; every formula has a cached value.
- The deck is generated from the model's exported results, so no figure is typed by hand.
- Source conflicts are logged rather than silently resolved.
