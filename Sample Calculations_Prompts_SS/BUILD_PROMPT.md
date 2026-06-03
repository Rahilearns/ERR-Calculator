# Build Prompt — ERR (Effective Rate of Return) Calculator

Build a single-page web app for bank Relationship Managers to compute the Effective Rate of Return
(ERR) on loans, generate amortization schedules, and download verifiable Excel/Word/PDF outputs.

## Reference files attached with this prompt
**Required:**
- `Sample format for verification.xlsx` — exact layout/styling/formulas to mirror for the Verify Calculation Excel (§7).
- `Rate_Revision_Structured_Case 1_COF Data.xlsx` — the COF upload template; ship to users **verbatim** as a static asset (§3.7, §6).

**Validation ground-truth (match your output to these):**
- `LF-S (EMPP) - Rectified.xlsx` — Loan Facilities Structured, Equal Principal + Interest (Monthly).
- `LF-S (EQPP) - Rectified.xlsx` — Loan Facilities Structured, Equal Principal + Interest (Quarterly).
- `Sample Customized Loan.xlsx` — Loan Facilities Customized (multi-layer + EMI-after-moratorium security).
- `Rate_Revision_Structured_Case 1_Rectified Calculation.xlsx` — Rate Revision Structured with COF upload.

## 0. Tech / delivery
- **Vanilla HTML/CSS/JS only. No build step, no framework.** Runs by opening `index.html`; deployable to GitHub Pages.
- CDN libs only: `xlsx-js-style` (styled Excel), `jspdf` + `jspdf-autotable`, `docx`, `FileSaver`, `flatpickr`.
- ES modules. **Version-stamp every relative import (`?v=NNN`)** so each deploy busts the browser module cache.
- Keep logic factored (shared input components + one calc engine); the 4 pages reuse them.

## 1. Shell & global UX
- App opens directly on the first calculator. **4 tabs across the top**, switch instantly, no reload:
  `Loan Facilities — Structured`, `Loan Facilities — Customized`, `Rate Revision — Structured`, `Rate Revision — Customized`.
- **Dark/light toggle**, persisted (localStorage). In dark mode **no text may be dark** (set explicit light colors on every control).
- Compact, tidy layout (~70% of default scale; ~780px max content width; small paddings).
- **Compare** button in header, **hidden until ≥1 summary is saved**.
- **Drafts → sessionStorage** (cleared on browser close, preserved across in-app tab navigation / reload). **Saved summaries → localStorage.**

## 2. Shared input components & formatting (apply everywhere)
- **Number/amount fields:** live thousands commas while typing; format to 2 decimals on blur; **right-aligned**. Blank until typed.
- **Percent/rate fields:** `%` fixed at the **far right** of the box; blank until typed; 2 decimals enforced on blur; **right-aligned**. Store as decimal (`12.50%`→`0.125`).
- **Option fields:** native `<select>` dropdowns, **center-aligned** text.
- **Date fields:** flatpickr, display format **`DD-Mmm-YYYY`** (e.g. `04-May-2026`); **Year shown as a dropdown** (±100 yrs) like the month dropdown; works in dark mode.
- **Layered/table fields:** column **headers centered and aligned with their inputs**; numeric/percent cells right-aligned, date/dropdown cells centered; the header row is **hidden when there are 0 rows**.
- **Info icons:** small `(i)` that opens a popover **on click** (not hover); closes on outside-click/Escape.
- Everywhere a value is an amount → commas; a rate → trailing `%`.

## 3. Calculation engine — exact rules (this is where precision matters most)

### 3.1 Common
- Schedule row 0 = **disbursement**: all zero, `URPA = loan amount`. (URPA = un-realized principal = outstanding balance after a payment.)
- **Installment relationship is universal:** `Installment = Interest + Principal + Accrued_carried_from_previous_row`. **Principal is computed independently** — never `Installment − Interest` (that leaks accrued interest into principal).
- **Interest Expense accrues every month** sl=1..N: `Expense[sl] = URPA_at_start_of_month[sl] × COF/12 = URPA[sl−1] × COF/12`. Disbursement row expense = 0.
- **Avg Portfolio** = average of start-of-month balances = `URPA[sl]` for `sl = 0 … N−1` (**exclude the final post-payment zero row**).

### 3.2 Moratorium + "Interest During Moratorium Period" (IDP)
- Each moratorium month: `interest = URPA × offeredRate/12`.
- A **checked** month = borrower **pays** that month's interest **plus all accrued** from prior unchecked months; accrued resets to 0.
- An **unchecked** month = interest **accrues** (running balance), no payment.
- Any accrued still outstanding at moratorium end is **paid with the first installment after moratorium**.

### 3.3 Loan Facilities — Structured payment modes (single mode for whole regular tenor)
- `EMI` / `EQI`: level installment = `PMT(rate/ppy, regularPeriods, −balanceAtMoratoriumEnd)`; ppy=12 (EMI) or 4 (EQI). Principal = installment − interest.
- `Equal Principal + Interest (Monthly)`: **principal constant = `Loan / (tenor − mora)`**; interest = `URPA × rate/12`.
- `Equal Principal + Interest (Quarterly)`: **principal constant = `Loan / ((tenor − mora)/3)`**; payments only on **absolute month % 3 == 0**; interest = `URPA × rate/4`.
- Final month settles any remaining principal + accrued.

### 3.4 Loan Facilities — Customized "Payment Layers"
- Layers are absolute month ranges `[from..to]` (after moratorium). SL numbering continues sequentially through moratorium and layers.
- `Customized Principal`: principal = the entered amount each month (capped at URPA); interest = `URPA × rate/12`.
- `EMI`: PMT over the layer's months. `EQI`: PMT over quarters in the layer; payments on month%3==0.
- **`Equal Principal + Interest` layers — principal is CONSTANT and sized over the periods remaining to MATURITY, not the layer's own span** (so a layer ending before maturity only partially amortizes; the remaining balance flows into the next layer):
  - Monthly: `principal = URPA_at_layer_start / (tenor − from + 1)`
  - Quarterly: `principal = URPA_at_layer_start / ((tenor − from + 1)/3)`; payments on month%3==0.

### 3.5 Funded Security (Loan Facilities)
- `FDR` / `Cash Security`: user enters **Amount + Rate**.
- `EMI after Moratorium` / `EQI after Moratorium`: amount = `PMT(rate/12, tenor−mora, −loan)` (EMI) or `PMT(rate/4, (tenor−mora)/3, −loan)` (EQI), **× Number of Installments**; rate = 0. (Hypothetical post-moratorium installment, independent of actual layers.)
- `Installment` (for Equal-Principal modes): amount = first EPI installment × Number of Installments; rate = 0.
- CS Benefit = `(COF − securityRate) × securityAmount × tenorYears`.

### 3.6 Loan Facilities ERR/metrics (display only ERR, NIM%, NII$)
- `NII = TotalInterest + CSBenefit − TotalInterestExpense`
- `NIM = NII / AvgPortfolio / tenorYears`
- `ERR = COF + NIM`
- Results show: **ERR** (headline), **NIM** (%), **Net Interest Income** ($). If an installment-type security was used, also show the derived **Security Amount**.

### 3.7 Rate Revision — Structured (date-driven, monthly)
- Rows sl 0..tenor; **row date = disbursement + sl months**.
- **Loan maturity = disbursement + tenor months − 1 day** (e.g. 01-Jan-2020 + 60mo → **31-Dec-2024**). Used only for the **last To Date** of Lending Rate Layers & Loan Security Layers, and as the COF cutoff. (Schedule row dates still run to disbursement+tenor months.)
- **One-month lag:** the lending rate and COF for month sl use the value effective at the **start of the accrual month = date[sl−1]**. **Security** amount/rate use the **current** date[sl].
- Lending rate on a date = covering layer's rate; past the last layer, extend the last rate forward.
- **EMI recomputes at each rate change:** `PMT(newRate/12, tenor − slAtChange, −balanceAtChange)`, where slAtChange = last sl before the change and balance = URPA there.
- COF from the uploaded file: COF on a date = the last record with `effectiveDate ≤ date`; **0% if the date is before the first record** (uncovered → show a warning); **ignore records effective after maturity**; the latest record on/before maturity extends to maturity.
- `Expense[sl] = URPA[sl−1] × COF[sl]/12`; `CSBenefit[sl] = securityBalance[sl] × (COF[sl] − securityRate[sl])/12` (0 where no security covers date[sl]).
- `NII = TotalInterest + ΣCSBenefit − TotalInterestExpense`; `NIM = NII/AvgPort/tenorYears`; `effectiveCOF = TotalInterestExpense/AvgPort/tenorYears`; **`ERR = NIM + effectiveCOF`**.
- If no COF file uploaded → COF = 0 everywhere (no expense); still calculates.

### 3.8 Rate Revision — Customized (uploaded schedule)
- User uploads a filled amortization schedule (cols `Date, Installment Amount, Interest Amount, Principal Amount, URPA`; first row = disbursement: date + URPA only).
- ERR = **URPA-exposure-weighted average periodic rate, annualized**: for each gap, `rate = Interest_i / URPA_prev × (365/days)`, weighted by `URPA_prev × days`.

## 4. The four pages — fields & conditional logic

### 4.1 Loan Facilities — Structured
Rows (2 fields/row unless noted): **Loan Amount (left) | Offered Rate (right)** · Moratorium Available? (No/Yes) | Moratorium Period (Months, shown when Yes) · **Interest During Moratorium Period** (full row; shown only when months>0) · Loan Tenor (Months) | Payment Mode · Total Cost of Fund [COF/ISC + OPEX] | Funded Security Type · **Cash Security/FDR Amount | Cash Security/FDR Rate** (shown for FDR/Cash Security) · Number of Installments (full row; shown for EMI/EQI/Installment security) · **Calculate ERR**.
- Payment Mode options: `EMI, EQI, Equal Principal + Interest (Monthly), Equal Principal + Interest (Quarterly)`.
- Security options: `FDR, Cash Security`, plus `EMI after Moratorium` / `EQI after Moratorium` / `Installment` depending on payment mode.
- When moratorium: rename `Payment Mode → Payment Mode after Moratorium`, `Loan Tenor → Loan Tenor including Moratorium (Months)`.
- **Interest During Moratorium Period:** label + `(i)` tooltip; **"Select all" checkbox on its own line below the label**; selectable **month boxes, equal width, max 6 per row**.

### 4.2 Loan Facilities — Customized
Same as Structured but **Payment Layers** replaces the single Payment Mode:
- **Payment Layers** table, columns: `From Date | To Date | Payment Type | Custom Principal`. **2 undeletable default rows.**
- Months shown as `Month NN`. **To-Date dropdown has `Last Month (Month NN)` at the top** = loan tenor.
- Payment Types: `Customized Principal, EMI, EQI, Equal Principal + Interest (Monthly), Equal Principal + Interest (Quarterly)`.
- **Custom Principal cell editable only when "Customized Principal" is selected.**
- Layer rules: Layer-1 `From = mora+1`; each subsequent `From = previous To + 1` (auto, editable); **last `To` defaults to tenor (editable until a new layer is added)**; **Add Layer disabled** until tenor entered AND all rows complete AND last To < tenor; **max layers = tenor**; allow `From == To`; **block overlaps & gaps; highlight conflicting fields and show a reason** before calculating.

### 4.3 Rate Revision — Structured
Rows: Initial Loan Amount | **Disbursement Date** (flatpickr; **disallow Fri/Sat**) · Moratorium Given at Disbursement? | Period · Interest During Moratorium Period (month boxes) · Payment Modality | Loan Tenor at Disbursement (Months) · **Lending Rate Layers** · **Loan Security Layers** · **COF Data Upload** · Calculate ERR.
- **Lending Rate Layers:** `From Date | To Date | Active Rate`. 2 undeletable rows; **first From = Disbursement Date**; cascade `From = previous To + 1 day`; **last To = maturity** (auto). Validate last To ≤ maturity.
- **Loan Security Layers:** `From Date | To Date | Amount | Active Rate`. Same first-From=disbursement, cascade +1 day, last To = maturity.
- **COF Data Upload:** `Get Sample COF Excel` (downloads the exact template) + `Upload COF Data`. (Replaces any "NIM comparison" toggle.)

### 4.4 Rate Revision — Customized
- `Get Sample Excel` (amortization template) + `Upload Excel`. · Loan Security Layers · *(optional NIM-margin comparison + COF layers)* · Calculate ERR.

## 5. Results panel (every page, after Calculate)
- Cards: **Effective Rate (ERR)** (highlighted), **NIM**, **Net Interest Income** (+ derived **Security Amount** when applicable).
- Action buttons: **Download Report** (PDF) and **Verify Calculation** (Excel).
- Below: **Amortization Schedule** table (Sl, [Date], Installment, Interest, Principal, URPA, [Accrued]); a **TOTAL** row; three summary cards **Total Principal Paid / Total Interest Paid / Total Payment** (centered); and a download row labeled **"Download the Schedule →"** with **Excel / Word / PDF** buttons.

## 6. Downloads
- **Schedule Excel / Word / PDF:** formatted amortization table + summary header.
- **Download Report (PDF):** all user inputs (including every layer/COF/IDP detail), all results, and the schedule — enough to **fully replicate the scenario**; all tables **center-aligned**.
- **Verify Calculation (Excel)** — see §7.
- **Sample files:** Rate Revision Customized ships a sample amortization template; Rate Revision Structured ships an **exact COF template** (an Excel *Table* with `Year, Month, Day, Date(formula), COF` columns) served **verbatim** as a static asset (preserve formatting/formulas/structure; user edits only values). COF parser reads Year/Month/Day + COF.

## 7. Verify Calculation Excel — exact, formula-driven (mirror the provided sample format)
Two sheets: **"Inputs & Results"** and **"Schedule"**. **Every number must be a traceable cell-referenced formula**, not a hardcoded value (except where an amount is a genuine input).
- **Styling:** A1 = navy fill, white bold, "ERR Calculator – Verification"; page-title + section headers = green fill, white bold; numbers in accounting format `_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)`; rates `0.00%` (inputs) / `0.0000%` (NIM, ERR); dates via `DATE(y,m,d)`.
- **Inputs block** (label | value), each at a known row so Results can reference it (e.g. Loan, Tenor, Mora, COF, Rate).
- **Schedule formulas:** `Interest = URPA_prev × rate/(12 or 4)`; `Principal` independent (Equal-Principal = constant formula referencing inputs/layer-start cells; EMI/EQI = base; Custom = the amount; final = remaining `URPA_prev`); `Installment = Interest + Principal + Accrued_prev`; `URPA = URPA_prev − Principal`; `Int Expense = URPA_prev × COF/12` (every month); **TOTAL** row = SUMs.
- **Results:** Total Interest Received, Total Interest Expense, CS Benefit, Net Interest Income, **Avg Portfolio = AVERAGE(URPA over sl 0..N−1, excluding the final zero row)**, NIM, **ERR** — all formulas linking the two sheets.
- **Yearly Summary** table: `Year | Interest Income | Interest Expense | Loan Security Benefit | Net Interest Income | Average Portfolio | NIM`, one row per loan year, formula-based.
- **Rate Revision — Structured variant** adds columns to the Schedule: `Date | Total COF | CS Balance | CS Benefit`; `Date=DATE()`; EMI installment uses the rate-block `PMT` recompute; `Int Expense = URPA_prev × COF/12`; `CS Benefit = CS Balance × (COF − securityRate)/12`; **ERR = NIM + (TotalExpense/AvgPort/tenor×12)**.

## 8. Save & Compare
- **Auto-save on Calculate ERR** (no manual save button): store inputs + results. **Skip exact duplicates** (fingerprint of pageType + inputs); different amounts/rates count as new. **Max 5** saved.
- **Compare page:** only compare **within the same module group** (Loan Facilities ↔ Loan Facilities, Rate Revision ↔ Rate Revision). Show ERR, NIM, NII; show inputs with **% (2 decimals)** and **comma-separated amounts**; rename `CS Amount → Security Amount`, `CS Rate → Offered Rate on Security Amount`; **hide IDP-flags**; allow delete per card.

## 9. Edge rules / gotchas (don't relearn these)
- flatpickr is configured `dateFormat:'d-M-Y'` → when setting a date programmatically, pass a **Date object** (an ISO string would be mis-parsed). Use `clear(false)` so programmatic clears don't fire the change handler (which would mark a field user-edited).
- Equal-Principal `urpa/(remaining)` per month is mathematically constant — but for Customized layers you must divide by **periods to maturity**, not periods in the layer.
- Disbursement-row interest expense = 0; interest expense accrues every month thereafter including the last.
- COF & lending rate in Rate Revision lag one month (start-of-accrual-month value); security uses the current month.
