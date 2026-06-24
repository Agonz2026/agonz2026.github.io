# PI Case Valuation Tool — Florida

A first-pass personal-injury case valuation tool for case managers preparing demands and
evaluating opening offers from insurers. Two interchangeable interfaces, identical math.

**Live tool:** https://agonz2026.github.io/pi-case-valuation/

## Files

- **`index.html`** — the web calculator (this is what loads at the live URL above). Self-contained;
  works offline too — just open the file in any browser. Includes a Print/Save-as-PDF button.
- **`PI_Case_Valuation_Calculator_FL.xlsx`** — Excel version. Case managers type into the input
  cells; all formula/result cells are locked. Has a `Reference` tab (editable factor tables) and a
  `Guide` tab (methodology + Florida law). To edit locked cells: Review → Unprotect Sheet (password `pi`).
- **`build_calc.py`, `lock_cells.py`, `verify.py`** — Python source that generates and tests the
  workbook (`pip install openpyxl`).

## How it values a case

1. **Economic damages (specials)** — past/future medicals, lost wages, out-of-pocket. Recoverable
   and summed directly. (Property damage is **not** included — see below.)
2. **Non-economic damages** estimated two ways and blended into a low–high range:
   - **Multiplier method** — medical specials × a severity multiplier (1.5×–10×).
   - **Per-diem method** — daily pain rate × number of recovery days.
   - Blended low = lower of the two methods; blended high = higher.
3. **Case-characteristic modifier** (applied to non-economic only, capped 0.4×–2.0×):
   permanent impairment rating (1 + whole-person % × 1.5, capped 1.6×), client age, prior injuries
   (same body part reduces value), and vehicle/impact severity.
4. **Property damage** is not a recoverable element of the BI claim. It is used only as a
   **causation signal**: low property damage against high medicals flags causation exposure;
   a totaled vehicle supports causation.
5. **Risk adjustments** — gross value × liability factor, reduced by the client's comparative
   fault %, then capped at available policy limits.
6. Outputs a **settlement target range**, an **opening demand**, a **settlement floor**, and an
   **offer-evaluation** score.

## Florida law baked in

- Modified comparative negligence (Fla. Stat. §768.81, amended 2023 by HB 837): a plaintiff
  **more than 50% at fault recovers nothing**; at/below 50%, damages reduce by the fault %.
- No-fault / PIP for auto ($10k); a bodily-injury claim generally requires meeting the
  permanent-injury threshold (Fla. Stat. §627.737) before non-economic damages are recoverable.
- No general statutory cap on non-economic damages in ordinary auto/PI cases.

## Disclaimer

Estimates only — **not legal advice**. Every case turns on its specific facts, venue, witnesses,
and the people involved. An attorney must review before any demand is sent or offer accepted.
