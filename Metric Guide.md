# Metric Guide — What Each Number Means

A plain-finance explanation of every metric in the Segmented Track Record Analysis output.
No Excel formulas, no code — just what each figure is, how it's derived conceptually, and
how to read it when evaluating a GP's track record.

**Conventions used throughout**

- All money figures are in **millions**, in the fund's reporting currency (default USD).
  Raw GP files that report in thousands or absolute dollars are converted automatically.
- Everything here is **gross** — before management fees, carried interest, and fund expenses.
  Gross returns overstate what an LP actually receives; treat them as a measure of the GP's
  *investment skill*, not of LP outcomes.
- All metrics are **deal-level** first, then aggregated. Aggregations are **capital-weighted**,
  not simple averages (details below — this matters a lot).

---

## 1. Deal identity & timing

| Metric | Meaning |
|---|---|
| **Status** | Every deal is classified as **Realized** (fully exited — including write-offs, which are "realized at ~zero") or **Unrealized** (still held, including partially-realized deals with a residual stake). |
| **Vintage** | The year the deal was made (year of the investment date). Grouping by vintage shows how performance varies across market environments — a GP whose returns all come from one lucky vintage is riskier than one who performs across cycles. |
| **Exit Year** | Year of exit for realized deals; "n/a" while still held. |
| **Hold Period** | Years between entry and exit (or blank until exit). PE value creation typically needs 3–6 years; very short holds can flag quick flips, very long holds can flag stuck assets ("zombie" deals). |

## 2. Capital and value

| Metric | Meaning |
|---|---|
| **Initial Invested Capital** | Equity the fund put in at the original close of the deal. |
| **Total Invested Capital (IC)** | All equity the fund ultimately invested, including follow-ons. This is the denominator for almost every return metric. |
| **Realized Value** | Cash actually returned to the fund from this deal — sale proceeds, dividends, recapitalizations. This is *hard* value: it happened. |
| **Current (Unrealized) Value** | The GP's own estimate of what the remaining stake is worth today (fair value / NAV). This is *soft* value: it is a mark, not cash, and GPs have discretion in setting it. Scrutinize the valuation method for large unrealized positions. |
| **Total Value** | Realized + Current. Everything the deal has produced, cash plus paper. |

## 3. Deal-level returns

**Gross MOIC (Multiple of Invested Capital)** — Total Value ÷ Total Invested Capital.
"For every dollar invested, how many dollars came back (or are marked)?" 2.0x means the money
doubled. Below 1.0x means the deal has lost value. MOIC ignores *time* — a 2.0x in 3 years is a
great outcome, a 2.0x in 12 years is mediocre. Always read it next to the hold period or IRR.

**Gross IRR** — the annualized rate of return of the deal's cash flows, as reported by the GP.
IRR captures the time dimension MOIC misses, but it is manipulable (early small distributions,
subscription lines) — MOIC and IRR together are more honest than either alone.

**Gross TVPI** — as reported by the GP, Total Value to Paid-In. At deal level it is conceptually
the same thing as MOIC; it appears as an input so the GP's own figure can be compared with the
one recomputed from their value numbers.

**Performing flag (1 = Underperforming)** — marks any deal whose MOIC is below 1.0x, i.e. the
deal is currently worth less than the money put in.

## 4. Loss metrics

These answer the question every LP asks: *how does this GP behave when things go wrong?*

**Invested Capital in Loss Position** — the IC of every deal marked below 1.0x. It measures
*exposure*: how much of the money went into deals that are underwater at all.

**Impaired Value** — for each underwater deal, the *depth* of the hole:
(1 − MOIC) × invested capital. A deal at 0.8x with $100m invested is impaired by $20m; a
write-off at 0.0x is impaired by the full $100m.

**Loss Ratio** (in the pivots) — capital in loss-making deals ÷ total capital of the group.
A 15% loss ratio means 15 cents of every invested dollar sits in deals below 1.0x — regardless
of whether they're down 5% or down 95%.

**Impaired Loss Ratio** — impaired value ÷ total capital. This is the *severity-adjusted*
version: actual value destroyed as a share of capital. Read the two together: a high loss ratio
with a low impaired ratio means many small stumbles; a low loss ratio with a high impaired
ratio means a few blow-ups. The second pattern is usually more worrying.

## 5. Entry & exit operating metrics

These describe the *companies*, not the fund — what the GP bought, at what price, with how much
debt, and what changed under ownership. All are shown at entry and again at exit/current.

| Metric | Meaning |
|---|---|
| **LTM Revenue / LTM EBITDA** | The company's last-twelve-months sales and operating profit at the measurement date. Comparing entry to exit shows whether the GP actually *grew* the business or just financially engineered it. |
| **EBITDA Margin** | EBITDA ÷ Revenue. Profitability of the business. Margin expansion from entry to exit is operational value creation. |
| **Net Debt** | Debt minus cash. The leverage placed on the company. |
| **Leverage (Net Debt / EBITDA)** | The standard PE leverage yardstick. Entry leverage of 5–6x is typical of buyouts; much higher means returns depend heavily on debt paydown, and downside risk is larger. |
| **Enterprise Value (EV)** | The total price of the whole business (equity + net debt). |
| **Equity Value** | EV − Net Debt: the part the fund's shareholders own. |
| **EBITDA Multiple (EV / EBITDA)** | The headline *price paid* (entry) or *price achieved* (exit). Comparing entry and exit multiples shows how much of the return came from **multiple expansion** (buying low / selling high, or just a hotter market) versus earnings growth. |
| **Equity Multiple (Equity Value / EBITDA)** | Like the EBITDA multiple but only on the equity slice — sensitive to how much leverage is in the structure. |
| **EV / Sales** | Price relative to revenue — the fallback yardstick when EBITDA is small or negative (common for growth deals). |

**How to read the trio:** any deal's value creation decomposes into (a) revenue/EBITDA growth,
(b) margin change, (c) multiple expansion, (d) debt paydown. Deals that made money purely on
(c) tell you less about GP skill than deals that made money on (a) and (b).

*(The IC-weighted CAGR/multiple columns at the far right of the Deal List compute automatically —
they are the per-deal feedstock for the Op Performance pivots: each deal's metric × its invested
capital, plus the matching capital denominator, so the pivots can divide the two sums into a
capital-weighted average for any filter combination.)*

## 6. Buckets (segmentation labels)

Deals are tagged into standard size/outcome bands so the pivots can segment the portfolio:
hold-period buckets (≤2 yrs … ≥8 yrs), invested-capital buckets, MOIC buckets (≤1.0x, 1–2x,
2–3x, ≥3x), entry revenue / EBITDA / EV / entry-multiple / margin buckets. The band thresholds
sit in small helper tables above the Deal List and can be edited — the labels recompute.
Their purpose is diagnostic: e.g. "does this GP earn its returns on small deals but lose money
when it writes big checks?" (MOIC by invested-capital bucket answers that).

## 7. The Return & Loss Ratios tab — how the pivot numbers aggregate

Each breakdown (by sector, geography, vintage, fund, exit type, buckets, …) shows three figures
per category, and this is the part where the aggregation method really matters:

- **Count** — number of deals in the category.
- **MOIC** — **pooled, capital-weighted**: the *sum* of the category's Total Value divided by the
  *sum* of its Invested Capital. It is *not* the average of the deals' MOICs. A $500m deal at
  1.2x and a $10m deal at 5.0x pool to roughly 1.3x, not 3.1x — the pooled figure reflects what
  actually happened to the *money*, which is the LP-relevant view. A simple average would let
  tiny home-runs mask big mediocre deals.
- **Loss Ratio / Impaired Loss Ratio** — as defined in section 4, computed over the category's
  pooled capital.

**Capital Deployment & Returns by Vintage** adds the sum of invested capital per vintage —
showing the *pacing* of the fund: how much money went out each year. Heavy deployment at market
peaks is a classic red flag.

Every pivot carries **Fund / Status / Hold Period filters**. Filtering to *Realized only* is the
single most useful cut: it strips out the GP's own marks and shows the track record in cash terms.

**Blank categories**: deals missing a label (e.g. no sector) still count in the pivot and its
Grand Total, but are excluded from the charts. **Charts are fixed snapshots** of the full,
unfiltered portfolio — filtering a pivot deliberately does not alter any chart.

## 7b. The Return Dispersion tab — reading the spread, not the average

The Return & Loss Ratios tab tells you the *pooled* result; this tab tells you
**how that result is distributed** — the single most underrated question in
track record analysis, because two GPs with the same 2.2x pooled MOIC can be
completely different investments.

Each section (Gross MOIC, Gross IRR) buckets every deal by its return level
and shows three numbers per bucket:

- **Count** — how many deals landed in the bucket (the *hit rate* view).
- **% IC** — how much of the invested capital landed there (the *money* view).
  This is what the chart plots. When Count and % IC diverge, you're looking at
  sizing: many small winners with big losers, or vice versa.
- **Average of the metric** — the actual MOIC / IRR level within the bucket
  (the *severity/magnitude* view). "≤1.0x average 0.4x" means the losers lose
  badly; "≤1.0x average 0.9x" means they merely stall. On the top bucket it
  separates "a bit above 3x" from "a couple of 20x outliers dragging the
  whole track record up."

What to look for:

1. **Left tail first**: % IC in the ≤1.0x / ≤0% buckets is the capital-at-risk
   picture, and the bucket average tells you how deep those holes go.
2. **Outlier dependence**: if the pooled MOIC is high but % IC is concentrated
   in one or two deals' bucket with an extreme average, returns rest on a few
   home runs — ask whether that's repeatable (same partner? same vintage?).
3. **Consistency**: a fat middle (1–3x, 10–20% IR) with modest tails is the
   signature of a repeatable process; a barbell (big ≤1.0x and big ≥3x) is a
   high-variance strategy — not necessarily bad, but it must be priced and
   sized as such.
4. **MOIC vs IRR disagreement**: deals can sit high in the MOIC dispersion
   and low in the IRR one (long holds) or the reverse (quick flips) — reading
   the two sections together is a fast duration check.

## 8. The Portfolio Construction tab

- **Invested Capital by Fund and Sector / Geography** — for each fund, the share of its invested
  capital in each sector or region (each fund's row sums to 100%). This is a *concentration*
  view: it answers "is this manager actually diversified, and has their strategy drifted from
  fund to fund?" The Status filter lets you view the mix on realized deals versus the current book.
- **Deal Count Attributes** — deal counts by sector, geography, transaction type (buyout,
  carve-out, take-private…), GP role (lead / co-lead), and process type (auction vs proprietary).
  Count views complement the capital views: many small proprietary deals versus a few large
  auction deals are very different sourcing stories even if the capital split looks similar.
  The standalone **Count of Company** figure is the total number of deals in the track record.

## 9. Dates

- **Track Record Date** — the "as of" date of the GP's own reporting; every value and mark is as
  of this date. This is the date that matters analytically.
- **Data as of** — simply the date this file was generated by the tool; it has no financial meaning.

## 10. Reading caveats — worth repeating

1. **Gross, not net.** Subtract roughly 5–7 points of IRR and ~0.3–0.5x of MOIC to guess net.
2. **Unrealized value is an opinion.** The larger the unrealized share of Total Value, the softer
   every headline number is. Check MOIC filtered to Realized as the acid test.
3. **Pooled MOIC weights by capital.** One big bad deal moves it more than several small wins.
4. **MOIC has no clock.** Cross-check against hold periods and IRR.
5. **Blank fields are data gaps in the GP's own reporting** — worth asking the GP about, since
   what's *missing* from a track record is often as informative as what's in it.
