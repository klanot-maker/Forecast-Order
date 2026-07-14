# Unit Conversion Calculator (Google Apps Script)

Converts each order line's total quantity into KG (or its native unit for
liquids) based on the pack size embedded in the item Name, and writes the
result into Column T.

## How it works

- **Input:** Column C (Name, e.g. `TOFU EXTRA FIRM MORI-NU 1 X 349GM`) and
  Column P (Quantity).
- **Output:** Column T (`Converted Total`).
- Parses the first `<number><unit>` found in the Name (`KG`, `KGS`, `G`,
  `GM`, `GRAM(S)`, `LTR(S)`, `ML`, `CL`, `L`), and an adjacent pack-count
  multiplier if present (`1X16KG`, `1Kg X 12 pcs`, `1 X 349GM`).
- `size × pack count × Quantity(P)` = total.
  - Weight units (G/GM/GRAM/KG) → expressed in **KG**.
  - Volume units (LTR/ML/CL) → left in their **own unit** (no density
    assumption is made to turn liters into kilograms).
- If no size/unit can be found in the Name (e.g. `... SPA KGS` or
  `... PAYSAN KG` with no leading number), Column T is left untouched.

## Runs automatically

- `onOpen()` backfills Column T for every existing row when the sheet is
  opened, and adds a **Unit Tools > Recalculate All (KG Conversion)** menu
  item for a manual re-run.
- `onEdit()` recalculates a single row automatically whenever its Name
  (Column C) or Quantity (Column P) is edited.

## Install

1. Open the Google Sheet.
2. **Extensions > Apps Script**.
3. Replace/add the contents of `Code.gs` with this repo's `Code.gs`.
4. Save, then reload the spreadsheet (or run `recalcAll` once from the
   Apps Script editor) to grant the required permissions.

## Configuration

Edit the constants at the top of `Code.gs` if your sheet layout differs:

- `SHEET_NAME` — restrict to one tab (leave `null` for the active sheet).
- `NAME_COL`, `QTY_COL`, `RESULT_COL` — column numbers (C=3, P=16, T=20).
