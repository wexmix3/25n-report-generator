# design.md — Financial Packet PDF

Scope: `src/lib/financial-packet-pdf.tsx` only. This is the rubric a visual eval
scores generated packets against (see `aios/scripts/month-close/visual_eval.py`).
Every rule here is checkable by looking at a rendered page — no rule should
require reading source code to verify.

## Palette

- Navy `#0F2044` — brand primary. Used for: header rule, section titles, table
  header background, KPI card background (neutral cards), brand name text.
- Gold `#C9A84C` — accent only. Used for: location subtitle under brand name,
  section-title underline rule (28×2px), KPI label text on navy cards.
- Green `#16A34A` / Red `#DC2626` — variance and balance-sensitive values only.
  Never used decoratively.
- Orange `#E07A3E` — reserved, currently unused in rendered output. Flag if it
  appears without a clear "why this row" reason.
- Grayscale (`#F9FAFB` → `#1F2937`) — table row striping, subtotal shading,
  secondary text (budget column, page footer).
- KPI cards use `#14532D` (dark green) / `#7F1D1D` (dark red) as background
  fills when a card represents a profit/loss figure whose sign should be
  visually obvious at a glance — never on non-financial KPI cards (e.g.
  Occupancy stays navy regardless of value).

## Typography

- Base font: Helvetica throughout — no other family should appear.
- Base body size: 8px. Table cells: 7px (bold variants use `Helvetica-Bold`
  at the same size, not a larger one). Section titles: 8.5px, uppercase,
  letter-spacing 1.0, centered. Footer text: 6px.
- No more than 2 font sizes on a single table (header row + body row) except
  the Monthly Trend table, which is allowed a smaller month-column size
  (5.5–6px) to fit more columns on one line.

## Spacing

- Page margins: 42px top, 44px bottom/sides — every page.
- Section title block: 16px margin-top, 8px margin-bottom, centered, with a
  gold 2px rule directly beneath it.
- Table rows: 3px vertical padding (5px horizontal) for body rows; 5px
  vertical for header rows. Subtotal rows get a top border (`gray300`) to
  separate them from the body rows above.
- KPI row: 10px gap between cards, 14px margin-bottom before the next section.

## Table conventions

- Every table has a navy header row with white bold text.
- Body rows alternate `white` / `gray50` — never two alternating rows in a row
  visually blend into the header or into a subtotal row.
- Subtotal/total rows always use `gray100` background + `gray300` top border,
  bold text — this must be visually distinct from a regular alternating row,
  not just "the last row before more text."
- Indented sub-line items (e.g. individual revenue categories under REVENUE)
  get 12px left padding — enough to read as a hierarchy at a glance, not a
  marginal nudge.

## Color-coded values (financial correctness signal, not decoration)

- A value is green if favorable, red if unfavorable — the definition of
  "favorable" flips for expense-type rows (actual ≤ budget = favorable) vs.
  revenue/income rows (actual ≥ budget = favorable). A reader should never see
  a red expense number that is actually good news, or vice versa.
- Aging-bucket values are red once they cross into the "overdue" range for
  that table — **31+ days for both AR and AP summary tables** (AR: Current/
  0-30 neutral, 31-60/61-90/Over 90 red; AP: 0-30 neutral, 31-60/61-90/Over 90
  red). This is a *lower* bar than the separate "Delinquent Accounts (61+
  Days)" drill-down section on the AR page, which only lists customers 61+
  days late — don't conflate the two thresholds, they're intentionally
  different (summary table warns earlier, drill-down lists the worse cases).
- Balance-sheet balance-check warnings use amber (`#FEF3C7` bg / `#92400E`
  text), not red — this is a "review" signal, not an error signal.

## What "looks wrong" means for this rubric

When scoring a rendered page, flag:
1. Any color used outside its defined role above (e.g. green on a neutral
   value, orange appearing without explanation).
2. Two adjacent rows that are hard to visually distinguish (missing
   alternating shade, subtotal row not standing out).
3. Text overflow, truncation, or a column visibly too narrow for its content.
4. Inconsistent font size within what should be one visual tier (e.g. two
   sibling KPI cards using different value font sizes).
5. Uneven whitespace — a section with visibly more/less breathing room than
   its siblings on the same page for no content reason.
6. A financial value whose color contradicts the favorable/unfavorable rule
   above (this doubles as a correctness check, not just aesthetics).

Findings are advisory. This file does not gate PDF generation — it's the
rubric a human (or the visual eval script) checks a packet against after
generation, per Max's rule that aesthetic findings get proposed and held for
approval rather than auto-fixed.
