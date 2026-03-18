# CLAUDE.md — Large Cap Stock Screener

## Mission
Build an institutional-grade equity stock screener targeting beaten-down large-cap value stocks with asymmetric upside. The screener identifies stocks that have sold off significantly but retain strong fundamentals, clear catalysts, and defensible competitive positions.

## Strategy
**Contrarian Value** — Buy quality companies during maximum pessimism when the bad news is priced in but the good news is not.

### Hard Filters (All Must Pass)
1. **ATH Selloff**: Must be >30% below all-time high
2. **Recovery Filter**: `(price - 52w_low) / (52w_high - 52w_low)` must be <50% — catch early-stage bottoming, not stocks that have already recovered
3. **Balance Sheet**: ND/EBITDA < 3.5x AND interest coverage > 3x
4. **No Hard Disqualifiers** (see below)

### Hard Disqualifiers (Automatic Fail)
1. **Criminal DOJ Investigation**: Active DOJ/FBI criminal investigation into company or senior executives → DISQUALIFIED. Unbounded downside risk from fines, executive departures, customer defection, reputational damage. Override only if investigation formally closed with no charges AND stock hasn't already recovered.
2. **Chronic Underperformance**: 3+ years of underperforming sector peers with no confirmed inflection point → DISQUALIFIED. Requires: (a) specific management action taken (not just announced), (b) early quantitative evidence of improvement, (c) external validation (analyst upgrades based on data, not hope).

## 7-Criteria Scoring System (v2 — Post-Mortem Updated)

| Criterion | Weight | Key Question |
|---|---|---|
| Valuation | 25% | How cheap is the stock? |
| Balance Sheet | 20% | Can it survive the downturn? |
| Catalyst Specificity | 20% | Named, dated, binary catalysts? |
| Selloff Quality | 15% | Single-event overreaction or chronic decay? |
| Competitive Moat Integrity | 10% | Is the competitive position intact? |
| Macro Alignment | 10% | Is macro a tailwind or headwind? |
| Analyst Conviction | 5% | What does smart money think? |

### Score Interpretation
- **8.0-10.0**: STRONG BUY — 3-5% position
- **6.5-7.9**: BUY — 2-3% position
- **5.0-6.4**: SPECULATIVE BUY — 1% or watchlist
- **3.0-4.9**: PASS — Do not initiate
- **0-2.9**: AVOID
- **DISQUALIFIED**: Hard disqualifier triggered — do not invest

### Positive Patterns (Learned from BSX, MDT Deep Dives)
1. **Single-Event Overreaction**: Selloff caused by one discrete event (tariff announcement, trial miss, one-time charge) rather than slow structural decay. These resolve faster and more completely. *Example: BSX sold off on tariff fears, not company-specific issues.*
2. **Named/Dated/Binary Catalysts**: Catalysts with a specific name (FARAPULSE), date (FDA PDUFA Q4 2026), and binary outcome (approved/not). These create clear re-rating triggers. Score 9-10 on catalyst specificity.
3. **Analyst Support Despite Selloff**: Strong buy consensus maintained or upgraded during/after selloff. Smart money sees the opportunity. *Example: BSX maintained strong buy consensus throughout selloff.*
4. **Intact Franchise**: Core business franchise undamaged — the selloff is about fear, not fundamentals. Revenue/earnings still growing. *Example: BSX 12%+ organic growth continued through selloff.*
5. **Management Action**: New CEO, strategic review, cost restructuring with concrete actions (not just promises). *Example: MDT Hugo robotic surgery platform as strategic pivot.*

### Failure Patterns (Learned from ZBH, UNH Deep Dives)
1. **Criminal Investigation = Disqualifier**: DOJ/FBI criminal probes create unbounded tail risk. No matter how cheap the stock looks, the downside is unknowable. *Example: UNH DOJ Medicare Advantage fraud investigation.*
2. **Chronic Underperformance = Disqualifier**: 3+ years of losing share to peers without confirmed inflection. "This time is different" is almost never true. *Example: ZBH chronic share loss to SYK with ROSA robot not yet proven.*
3. **Competitor Structural Share Loss**: When a competitor is structurally gaining share (not a temporary blip), apply: -2pts to selloff quality, -1.5pts to competitive moat integrity. *Example: LLY gaining GLP-1 share from NVO.*
4. **Macro Hostility**: When the macro/regulatory environment is actively hostile (bipartisan political targeting, regulatory crackdown), cap macro alignment score at 3.0. *Example: UNH facing bipartisan healthcare reform pressure.*

## Sector Diversification

### Expanded Hunting Grounds
| Sector | Sub-Industries | Valuation Anchors |
|---|---|---|
| Healthcare | Pharma, Med-tech, MCO, Biotech | Fwd P/E 12-35x depending on sub-industry |
| Defense / Aerospace | Primes, suppliers, space | Fwd P/E ~22x; backlog/book-to-bill key |
| Consumer Staples | Food, beverage, household, personal care | Fwd P/E ~20x; volume vs. pricing mix |
| Industrials | Machinery, electrical, distribution | Fwd P/E ~18x; use mid-cycle earnings |
| Utilities | Regulated, renewable, water | Fwd P/E ~16x; dividend yield primary |
| International | ADRs, developed market value | Fwd P/E ~15x; country risk discount |

### Sector Balance Rule
- **Target**: Min 3 sectors with qualifying candidates
- **If unachievable**: Document why and proceed with available sectors. Do not force bad picks to meet sector targets.
- Each sector table gets its own color-coded section on the Summary Dashboard.

## Excel Model Structure (9 Sheets)

1. **SUMMARY DASHBOARD** — Master ranking + per-sector tables + screened-out table
   - New columns: Catalyst Name & Date, Selloff Type, Criminal/Regulatory Flag
   - Sector color-coding on ticker column background
   - DISQUALIFIED rows styled with red strikethrough
2. **STOCK SCORES** — 7-criteria scoring matrix with justifications
3. **VALUATION COMPS** — Multi-metric valuation comparison
4. **FUNDAMENTAL DATA** — Income statement, balance sheet, cash flow summary
5. **CATALYST TRACKER** — Named catalysts with dates, probabilities, expected impact
6. **RISK MATRIX** — Bull/base/bear scenarios with expected value
7. **NVO DEEP DIVE** — Institutional deep dive on primary candidate
8. **ASSUMPTIONS** — Macro assumptions, valuation methodology, peer groups, data sources
9. **SCREENER LOGIC** — Documents scoring methodology, weights, disqualifiers, positive/failure patterns, sector rules

## Sector Colors (Pastel, for Ticker Column Background)
- Healthcare: #BDD7EE (light blue)
- Defense / Aerospace: #E2EFDA (light green)
- Consumer Staples: #FFF2CC (light yellow)
- Industrials: #FCE4D6 (light orange)
- Utilities: #EDEDED (light gray)
- International: #E2D9F3 (light purple)

## Technical Notes

### FMP API
- **Base URL**: `https://financialmodelingprep.com/stable/`
- **API Key**: `YaPRXrJCQ4Uw8whrmdEbxnH22p2kPR7x`
- v3 API is deprecated; use stable endpoints only

### NVO Currency
- FMP returns NVO financials in DKK
- Conversion rate: 6.85 DKK/USD
- Apply to EPS, revenue, all per-share metrics

### File Structure
```
output/
  research_data.json          # Main research data (stocks + screened_out + assumptions)
  research_data_template.json # Template for new stock entries
  stock_screener.xlsx         # Excel model output
  NVO_deep_dive_*.txt         # NVO deep dive text
config/
  scoring_rubric.json         # Scoring criteria, weights, disqualifiers
scripts/
  build_model.py              # Excel model builder
  recalc.py                   # Formula audit and validation
examples/
  successful_trades.json      # Historical trade examples for pattern reference
```

### Build Commands
```bash
python scripts/build_model.py                    # Build Excel model
python scripts/recalc.py output/stock_screener.xlsx  # Validate formulas
```

### Git Branch
- Development branch: `claude/initial-setup-fQGvK`
