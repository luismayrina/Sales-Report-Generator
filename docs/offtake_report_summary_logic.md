# Offtake Report Summary — Generation Logic

## Purpose

This sheet provides a high-level monthly sales breakdown by channel. Each row is a sales channel (Shopify, Shopee, Lazada, physical stores, etc.) and each column group represents a calendar month.

---

## Column Structure

The template uses a **repeating 10-column block per month**, starting at column B. The pattern is:

| Offset | Column Purpose        | Written By  |
|--------|-----------------------|-------------|
| 0      | % cont (monthly)      | App         |
| 1      | Actual sales (month)  | App (online) / Template (stores) |
| 2      | Target                | Template (preserved) |
| 3      | PY (Previous Year)    | Template (preserved) |
| 4      | vs PY                 | App (if PY exists) |
| 5      | (gap column)          | Template    |
| 6      | % cont (YTD)          | App         |
| 7      | YTD actual            | App         |
| 8      | YTD PY                | Template (preserved) |
| 9      | YTD vs PY             | App (if PY exists) |

**Example:**
- January section: columns B–K (cols 2–11)
- February section: columns L–U (cols 12–21)
- March section: cols 22–31 (if added to template)

The app detects month sections **dynamically** by finding `datetime` values in row 1. No month is hardcoded.

---

## Row Structure

| Row | Channel                | Type           |
|-----|------------------------|----------------|
| 2   | Common Room            | Physical store |
| 3   | Frankie and Friends    | Physical store |
| 4   | Craft Central          | Physical store |
| 5   | Simula PH              | Physical store |
| 6   | 9 Matters              | Physical store |
| 7   | Lazada                 | Online (App)   |
| 8   | Shopee                 | Online (App)   |
| 9   | Shopify                | Online (App)   |
| 10  | Pick-a-roo             | Physical store |
| ... | ...                    | ...            |
| 20  | TOTAL WITH INDUSTRIAL  | Calculated     |

---

## What the App Calculates

### Online Channel Rows (Shopify, Shopee, Lazada)

For each month section detected in the template:

1. **Actual sales**: Sum of `net` field for all non-cancelled orders in that month.
2. **YTD**: Cumulative sum of actual sales from January up to that month.
3. **% cont**: `row_actual / total_actual` for monthly; `row_ytd / total_ytd` for YTD.
4. **vs PY**: `actual - PY` if a PY value exists in the template for that row/column.
5. **YTD vs PY**: `ytd - ytd_PY` if a YTD PY value exists in the template.

> **Important**: The app always overwrites the online channel cells (actual and YTD) to prevent old template data from showing through. Even a value of 0 is written explicitly.

### Physical Store Rows

Physical store rows (Common Room, Frankie and Friends, etc.) are **not touched by the app**. Their values are copied directly from the template as-is. If you need to update them, edit the `Sample Reports.xlsx` template directly.

### TOTAL Row

The TOTAL row is recalculated by summing all channel rows for each month and YTD. This ensures it reflects both the physical store data from the template and the newly calculated online channel data.

---

## What the App Preserves

| Field  | Behavior                                                          |
|--------|-------------------------------------------------------------------|
| Target | Preserved from template; not overwritten                          |
| PY     | Preserved from template; used to calculate `vs PY` but not changed|
| Styling| All cell formatting (font, fill, border, alignment) is copied     |

---

## Formula for % Contribution

```
monthly % cont = row_actual / total_actual
YTD % cont     = row_ytd    / total_ytd
```

Division-by-zero is avoided: if the denominator is 0 or not a number, the % cont cell is not written.

---

## Formula for vs PY

```
vs PY     = actual     - PY
YTD vs PY = ytd_actual - ytd_PY
```

PY values are read from the template. If a PY cell is empty or 0, `vs PY` is not written.

---

## YTD Calculation

YTD is calculated as a **running cumulative sum** as the app iterates through months left-to-right. For each month section found in the template:

```
ytd_sum += monthly_actual_for_that_month
```

This means if the template only has January and February columns, YTD at February = Jan + Feb.
If the template is extended with March, YTD at March = Jan + Feb + Mar automatically.

---

## Dynamic Month Detection

The app reads row 1 and looks for `datetime` values. Each `datetime` marks the "actual" column of that month. The full 10-column section is derived from that position. This means:

- No month is hardcoded.
- Adding more months to the template automatically adds more populated columns.
- The existing `#DIV/0!` error values in the template (from empty % cont formula cells) are cleared and replaced with calculated values.

---

## Current Limitations

| Limitation | Detail |
|---|---|
| Target | Not calculated from uploaded data. Preserved from template. To update targets, edit `Sample Reports.xlsx`. |
| PY | Not calculated from uploaded data. Preserved from template. Historical year data is not an input file. |
| Physical stores | Not updated by the app. To update Common Room, Craft Central etc., edit the template directly. |
| Shopee/Lazada Dec data | If data exists for a month not in the template (e.g., December, but template only goes to February), those sales are calculated internally but cannot be written to the sheet. The template must be extended first. |

---

## Developer Reference

| Step | Description | Location |
|---|---|---|
| 1 | Copy entire template to output sheet | `write_offtake_summary` — Step 1 |
| 2 | Detect month sections from row 1 datetimes | Step 2 |
| 3 | Detect channel rows from column A values | Step 3 |
| 4 | Calculate monthly net totals per online channel | Step 4 |
| 5 | Write actuals + YTD for online channels | Step 6 |
| 6 | Recalculate TOTAL row | Step 7 |
| 7 | Calculate and write % cont | Step 8 |
| 8 | Calculate and write vs PY | Step 6 (within channel loop) |

All debug output is available in the app's **Generation Progress** panel when a report is generated.
