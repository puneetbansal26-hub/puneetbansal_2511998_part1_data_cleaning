# Cleaning Log — raw_orders.xlsx

**Project:** Business Data Cleaning, Validation & Excel Reporting  
**Dataset:** raw_orders.xlsx (932 rows, 21 columns)  
**Cleaned Output:** cleaned_orders.xlsx  

---

## Issues Found

### 1. Text Field Issues

This category had the most scattered problems. Almost every text column had at least some inconsistency.

**Whitespace issues:**
- `region` column: values like `"  North "` and `"  East "` with leading/trailing spaces
- `segment`: `"Corporate "`, `"  Small Business "`, `"Ananya Rao "` (trailing space in customer_name)
- `category`: `"  Furniture "`, `"Technology "` with trailing space
- `ship_mode`: `"Second Class "` — looked clean on the surface, failed exact-match lookups
- `payment_status`: `"Paid "` in several rows
- `order_status`: `"Completed "` with trailing space

**Case inconsistencies:**
- `segment`: "SMALL BUSINESS", "corporate", "consumer", "HOME OFFICE", "Small  Business" (double space)
- `region`: "NORTH", "WEST", "west", "east", "north" — basically all four regions had at least one all-caps or all-lowercase variant
- `category`: "OFFICE SUPPLIES", "office supplies", "Office  Supplies", "FURNITURE", "TECHNOLOGY"
- `sub_category`: "LABELS" (row 856), "storage" (row 378)
- `ship_mode`: "STANDARD CLASS" (row 38), "FIRST CLASS" (rows 424, a few others)
- `payment_status`: "PENDING", "failed"
- `order_status`: "completed", "COMPLETED", "cancelled", "CANCELLED"

Total affected cells: 30+ across the dataset (rough count, not exhaustive since TRIM handles most of these systematically).

---

### 2. Date Format Issues

This was the most time-consuming part. Five distinct date formats appeared across `order_date` and `ship_date`:

| Format | Example | Approx. Frequency |
|--------|---------|------------------|
| DD MMM YYYY | `21 Jul 2024` | ~25% of rows |
| MM/DD/YYYY | `07/27/2024` | ~25% of rows |
| YYYY-MM-DD | `2024-07-22` | ~30% of rows |
| DD-MM-YYYY | `28-11-2024` | ~20% of rows |
| Inconsistent mixing within same row | Various | Common |

Several rows had one format for `order_date` and a different format for `ship_date` on the same row, which made formula-based parsing unreliable without first converting everything to a consistent base.

**Ship date before order date (flagged as invalid):**
- ORD-2024-10073: order_date = 11 Feb 2024, ship_date = 07 Feb 2024
- ORD-2025-10128: order_date = 30 Jan 2025, ship_date = 26 Jan 2025
- ORD-2024-10143: order_date = 30 Mar 2024, ship_date = 26 Mar 2024

**Ambiguous date checked:**
- ORD-2024-10895: ship_date = 02/29/2024 — validated as legitimate. 2024 is a leap year, so Feb 29 exists. No flag needed.

---

### 3. Missing Values

| Column | Missing Count | Action |
|--------|-------------|--------|
| `region` | 26 | Filled as "Unknown", flagged in quality report |
| `ship_mode` | 22 | Filled as "Unknown", flagged in quality report |
| `discount` | 18 | Treated as 0 where sales fields were internally consistent; flagged otherwise |

The missing region records are spread across different states — this doesn't seem to follow a pattern, so it's likely a data export issue from the source system rather than a deliberate omission.

---

### 4. Duplicate Records

Found 33 rows with duplicate `order_id` values, covering 16 unique order IDs.

**Exact duplicates (same in every field) — 18 rows removed:**

| Order ID | Copies | Action |
|----------|--------|--------|
| ORD-2024-10503 | 2 | Kept 1, removed 1 |
| ORD-2024-10599 | 2 | Kept 1, removed 1 |
| ORD-2024-10705 | 2 | Kept 1, removed 1 |
| ORD-2024-10713 | 2 | Kept 1, removed 1 |
| ORD-2025-10007 | 2 | Kept 1, removed 1 |
| ORD-2024-10277 | 2 (from exact match group) | Kept 1, removed 1 |
| ORD-2025-10378 | 2 | Kept 1, removed 1 |
| ORD-2025-10400 | 2 | Kept 1, removed 1 |
| ORD-2025-10550 | 2 | Kept 1, removed 1 |
| ORD-2025-10633 | 2 | Kept 1, removed 1 |
| ORD-2025-10677 | 2 | Kept 1, removed 1 |
| ORD-2025-10695 | 2 | Kept 1, removed 1 |
| ORD-2025-10720 | 2 | Kept 1, removed 1 |
| ORD-2025-10730 | 2 | Kept 1, removed 1 |

**Conflicting duplicates (different values in key fields) — retained both, flagged:**

| Order ID | Conflict |
|----------|---------|
| ORD-2024-10124 | 3 copies — two with sales=634.13/Cancelled, one with sales=808.67/Returned |
| ORD-2024-10134 | Exact match — treated as exact duplicate (kept 1) |
| ORD-2024-10143 | Different sales value and order_status (Completed vs Returned) |
| ORD-2024-10251 | Different sales values |
| ORD-2024-10273 | sales=3510.42/Completed vs sales=3797.79/Returned |
| ORD-2024-10281 | Same sales, different sales formula outcome |
| ORD-2024-10332 | sales differ AND order_status differs |
| ORD-2024-10424 | sales differ slightly; both show FIRST CLASS (case inconsistency too) |
| ORD-2024-10850 | Minor trailing space difference in order_status (`"Completed "` vs `"Completed"`) — treated as exact duplicate after trimming |
| ORD-2025-10091 | sales=9445.10/Completed vs sales=9549.44/Returned — sales AND status mismatch |
| ORD-2025-10128 | Same sales but ship date is before order date in both — both flagged as invalid ship date |
| ORD-2025-10171 | sales differ and order_status differs (Completed vs Returned) |
| ORD-2025-10225 | sales differ; order_status: Cancelled vs Returned |
| ORD-2025-10315 | sales differ; one shows Completed, one Returned |
| ORD-2025-10336 | sales differ; one Completed, one Returned |
| ORD-2025-10374 | sales differ; one Completed, one Returned |
| ORD-2025-10572 | sales differ; one Returned, one Completed — also missing ship_mode |

All conflicting duplicates are marked with `data_quality_flag = "invalid"` and retained in `cleaned_orders.xlsx` for traceability. They are excluded from pivot summaries.

---

### 5. Discount Issues

**Negative discounts (16 rows, 15 unique orders — ORD-2024-10277 appears twice due to being a duplicate):**

ORD-2024-10020 (-0.19), ORD-2025-10022 (-0.23), ORD-2024-10094 (-0.09), ORD-2024-10125 (-0.15), ORD-2025-10202 (-0.16), ORD-2024-10277 (-0.14), ORD-2024-10355 (-0.06), ORD-2025-10401 (-0.10), ORD-2024-10495 (-0.13), ORD-2025-10516 (-0.07), ORD-2024-10657 (-0.05), ORD-2024-10664 (-0.24), ORD-2024-10819 (-0.14), ORD-2025-10844 (-0.02), ORD-2024-10857 (-0.16)

All flagged as `invalid` in `data_quality_flag`.

**High discounts above 50% (7 rows):**

ORD-2024-10003 (0.55), ORD-2024-10192 (0.65), ORD-2024-10239 (0.55), ORD-2024-10294 (0.65), ORD-2024-10306 (0.55), ORD-2025-10686 (0.55), ORD-2025-10708 (0.55)

Flagged as `warning`. Not impossible in retail contexts (clearance sales, bulk deals), but unusual enough to need review.

---

### 6. Order Status and Payment Contradictions

- **ORD-2025-10315**: `payment_status = Failed` but `order_status = Completed`. A failed payment should not result in a completed order. Flagged.
- Several records have `order_status = Returned` but `payment_status = Pending` — technically possible (payment wasn't made, item was returned), but flagged for awareness.
- Orders with `order_status = Cancelled` but `payment_status = Paid` — refund may have been issued. These are included in the refunded summary.

---

### 7. Sales/Profit Calculation Mismatches

Some rows have a `sales` value in the raw data that doesn't match `quantity × unit_price × (1 - discount)`. Possible causes: rounding differences, currency conversion, or data entry errors.

Where mismatches exceed ±1% of the expected value, those rows are flagged as `warning` in `data_quality_flag`. The new `calculated_sales` column always reflects the correct formula-based value; the original `sales` column is preserved for reference.

---

## Cleaning Actions Performed

1. **TRIM applied** to all text fields via formula columns, then pasted as values in the cleaned file
2. **PROPER/consistent casing applied** to: customer_name, city, state, product_name; Title Case applied to segment, region, category, sub_category, ship_mode, payment_status, order_status
3. **Date standardization**: all dates parsed and converted to DD/MM/YYYY format; shipping_delay_days calculated
4. **Exact duplicates removed**: one copy retained, duplicate row deleted
5. **Conflicting duplicates flagged**: both records kept, marked with `data_quality_flag = "invalid"`
6. **Missing region and ship_mode**: filled with "Unknown" text
7. **Missing discount**: set to 0 where sales fields were consistent, otherwise left null and flagged
8. **Negative discounts**: `cleaned_discount` set to 0, original value preserved, row flagged as invalid
9. **High discounts**: left as-is in `cleaned_discount`, flagged as warning
10. **Calculated columns added**: cleaned_discount, calculated_sales, calculated_profit, profit_margin, shipping_delay_days, order_month, order_year, data_quality_flag

---

## Business Rules Applied

- **Missing region / ship_mode → fill "Unknown"**: Deleting these rows would lose valid financial data. Filling preserves the records and the quality report documents them clearly.
- **Missing discount → treat as 0**: In the context of this dataset, no discount value and a discount value of zero produce the same result mathematically, so as long as the rest of the row checks out, this is a safe assumption.
- **Negative discounts → invalid**: A negative discount would imply a surcharge or a data entry error. Either way, it's not a legitimate discount and shouldn't flow into any revenue calculation.
- **Discounts above 50% → warning**: There's no stated maximum in the business rules beyond flagging "unusually high discounts." I used 50% as the threshold based on what's reasonable for retail. This can be adjusted.
- **Cancelled, Failed, Refunded → excluded from completed-sales summaries**: These don't represent realized revenue. Pivot tables covering sales performance are filtered to `order_status = Completed` and `payment_status = Paid`.
- **Refunded orders → separate summary**: Summarized by region and category in the pivot report.
- **Ship date before order date → invalid shipping record**: This is physically impossible and indicates a data entry error. Records are flagged but retained.
- **Conflicting duplicates → flag, don't delete**: Without source system access, there's no way to know which version of a duplicate is correct. Both are retained with a flag.

---

## Assumptions Made

- Date format ambiguity: where a date could be read as either MM/DD or DD/MM (e.g., `07/08/2024`), the MM/DD/YYYY interpretation was assumed, as it matches the majority pattern in this dataset and aligns with US-format exports.
- The 50% discount threshold is a judgment call. It's not specified in the assignment brief. If the company has a different policy, this threshold should be updated accordingly.
- For the `calculated_sales` formula, I assumed: `sales = quantity × unit_price × (1 - discount)`. This is the standard multiplicative discount formula. Some rows in the raw data use a different rounding approach, causing minor mismatches.
- When a row has both a conflicting status and a negative discount, both flags are applied — the row gets `data_quality_flag = "invalid"` and both issues are noted separately in the quality report.
- February 29, 2024 (ORD-2024-10895) was treated as a valid date. 2024 is a leap year, so this date exists.

---

## Records Removed

- **Exact duplicate rows removed**: approximately 18 rows (one copy of each exact duplicate was deleted)
- No other rows were permanently removed

---

## Records Flagged

| Flag Type | Approximate Count |
|-----------|------------------|
| Missing region (Unknown filled) | 26 |
| Missing ship_mode (Unknown filled) | 22 |
| Missing discount (set to 0) | 18 |
| Negative discount (invalid) | 16 |
| High discount >50% (warning) | 7 |
| Ship date before order date (invalid) | 3 |
| Conflicting duplicate (invalid) | ~15 rows |
| Payment/status contradiction (warning) | ~5 rows |
| Sales calculation mismatch (warning) | Varies — documented in quality report |

All flagged records are visible in `cleaned_orders.xlsx` via the `data_quality_flag` column.

---

## Limitations of This Cleaning Process

- **Excel-only approach**: Everything was done manually in Excel. This works fine for 932 rows, but it's not reproducible or scalable. Any changes to the raw file would require re-doing the cleaning steps by hand.
- **Date parsing in Excel**: Excel's date recognition is inconsistent with mixed-format columns. Some conversions required manual intervention and may not catch every edge case automatically.
- **Conflicting duplicates are unresolved**: The correct version of a conflicting duplicate can only be determined by checking the source system. That's not something that can be fixed from the data alone.
- **No validation against external reference data**: City/state/region combinations weren't verified against an official geography master. A state listed under the wrong region (if it happened) would not be caught.
- **Discount threshold is subjective**: The 50% warning threshold is an assumption. Without a documented discount policy from the company, this could be wrong in either direction.
