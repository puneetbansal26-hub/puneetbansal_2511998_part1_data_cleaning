# Part 1: Business Data Cleaning, Validation & Excel Reporting

## Problem Summary

A retail company pulled order-level sales data from multiple internal systems and ended up with a pretty messy dataset. The file had a mix of date formats, inconsistent text casing, duplicate records, missing values, negative discounts, and cases where the calculated sales figures didn't quite match what was stored. The goal here was to clean all of that up, apply a set of business rules, document every decision made, and produce summary reports ready for a business review.

## Dataset Description

The raw file (`data/raw_orders.xlsx`) contains **932 rows** of order data across **21 columns**, including order and customer identifiers, dates, geography (region, state, city), product details (category, sub-category), shipping info, financial figures (unit price, discount, sales, cost, profit), and status fields (payment status, order status).

Key fields:
- `order_id`, `customer_id`, `customer_name`
- `order_date`, `ship_date`
- `segment`, `region`, `state`, `city`
- `category`, `sub_category`, `product_name`, `ship_mode`
- `quantity`, `unit_price`, `discount`, `sales`, `cost`, `profit`
- `payment_status`, `order_status`

The dataset also includes a `business_rules` sheet that outlines the rules students are expected to apply during cleaning.

## Tools Used

- Microsoft Excel (data cleaning, TRIM/SUBSTITUTE formulas, Find & Replace, pivot tables)
- Manual review for duplicate and date anomaly identification

## Cleaning Steps Performed

### Step 1 — Preserving the Raw File
Left `raw_orders.xlsx` completely untouched. All work was done in a separate copy saved as `cleaned_orders.xlsx`.

### Step 2 — Text Field Standardization
Applied TRIM to strip leading/trailing spaces across all text fields. Used Find & Replace and PROPER/UPPER/LOWER functions to normalize casing. The messy ones were `segment` (had "SMALL BUSINESS", "corporate", "Small  Business"), `region` ("NORTH", "west", "east"), `category` ("OFFICE SUPPLIES", "office supplies", "Office  Supplies"), `sub_category` (e.g. "LABELS", "storage"), `ship_mode` ("STANDARD CLASS", "FIRST CLASS"), `payment_status` ("PENDING", "failed"), and `order_status` ("completed", "CANCELLED"). Some fields also had trailing spaces embedded in values, which TRIM sorted out.

### Step 3 — Date Cleaning
Cleaning the dates was genuinely a bit of a nightmare. The dataset had at least five distinct formats:
- `21 Jul 2024` (day month-abbr year)
- `07/27/2024` (MM/DD/YYYY)
- `2024-07-22` (ISO YYYY-MM-DD)
- `28-11-2024` (DD-MM-YYYY)
- `05 Sep 2024` (day month-abbr year variant)

All dates were converted to a consistent `DD/MM/YYYY` format. After parsing, a `shipping_delay_days` column was added as `=ship_date - order_date`. Any rows where this came out negative were flagged as invalid shipping records.

Notable issues found:
- ORD-2024-10073: ship date (07 Feb) is before order date (11 Feb) — flagged
- ORD-2025-10128: ship date (26 Jan) is before order date (30 Jan) — flagged
- ORD-2024-10143: ship date (26 Mar) is before order date (30 Mar) — flagged
- ORD-2024-10895: ship date listed as 02/29/2024 — February 29, 2024 is actually valid (2024 is a leap year), so this was accepted

### Step 4 — Duplicate Handling
Found **33 duplicate order_id entries** spread across 16 unique order IDs. These fell into two types:

**Exact duplicates** (same data, just repeated rows): retained one copy, removed the duplicate. Examples: ORD-2024-10503, ORD-2024-10599, ORD-2025-10633, ORD-2025-10677, etc.

**Conflicting duplicates** (same order_id, different values): flagged for review instead of silently deleting. Examples:
- ORD-2024-10124 — appears three times, two with `sales = 634.13 / Cancelled` and one with `sales = 808.67 / Returned`
- ORD-2025-10091 — one row shows sales of 9445.10 (Completed), another shows 9549.44 (Returned)
- ORD-2024-10332 — sales mismatch, and order_status differs (Completed vs Returned)

All conflicting duplicates are documented in the data quality report with both versions retained and flagged.

### Step 5 — Business Rules Applied

| Rule | Action Taken |
|------|-------------|
| Missing region (26 rows) | Filled with "Unknown"; flagged in quality report |
| Missing ship_mode (22 rows) | Filled with "Unknown"; flagged in quality report |
| Missing discount (18 rows) | Treated as 0 where other sales fields were valid |
| Negative discounts (16 rows across 15 unique orders) | Flagged as `invalid` in data_quality_flag column |
| High discounts >50% (7 rows) | Flagged as `warning` — unusual but not impossible, flagged for review |
| Cancelled orders | Excluded from completed sales summaries |
| Failed payment orders | Excluded from completed sales summaries |
| Refunded orders | Summarized separately |
| Ship before order date | Flagged as invalid shipping record |
| ORD-2025-10315 | Failed payment + Completed status — contradiction flagged |

### Step 6 — Calculated Columns Added

All calculated columns were built using Excel formulas, not hardcoded values.

- `cleaned_discount` — standardized discount after handling nulls and flagging invalids
- `calculated_sales` — `= quantity × unit_price × (1 - cleaned_discount)`
- `calculated_profit` — `= calculated_sales - cost`
- `profit_margin` — `= calculated_profit / calculated_sales`
- `shipping_delay_days` — `= ship_date - order_date`
- `order_month` — `= MONTH(order_date)`
- `order_year` — `= YEAR(order_date)`
- `data_quality_flag` — `clean`, `warning`, or `invalid` based on all issues found per row

## Business Rules Applied

- Cancelled, Failed-payment, and Refunded orders are excluded from the main completed sales totals
- Negative discounts are invalid by definition and flagged accordingly
- Discounts above 50% are unusual enough to warrant a warning flag
- Missing region and ship_mode are filled as "Unknown" rather than deleted, to preserve the record
- Missing discount is only set to 0 if the sales, cost, and profit values look internally consistent; otherwise it's flagged
- Ship date earlier than order date is a data entry error — flagged, not deleted

## Summary of Data Quality Issues Found

| Issue Type | Count |
|-----------|-------|
| Missing region values | 26 |
| Missing ship_mode values | 22 |
| Missing discount values | 18 |
| Negative discounts | 16 (15 unique orders) |
| Discounts above 50% | 7 |
| Case/whitespace issues in text fields | 30+ cells across multiple columns |
| Duplicate order IDs (all) | 33 rows across 16 order IDs |
| — Exact duplicates (safe to remove) | ~18 rows |
| — Conflicting duplicates (flagged) | ~15 rows |
| Ship date before order date | 3 confirmed records |
| Payment/status contradiction | 1 record (ORD-2025-10315) |

## Summary of Final Pivot Reports

The `pivot_summary.xlsx` file contains six pivot tables:

1. **Sales & Profit by Region** — South led in total sales; North had the highest average profit margin
2. **Sales & Profit by Category and Sub-Category** — Technology was the top revenue category; Phones and Copiers drove the most volume within it
3. **Order Count by Ship Mode** — Standard Class was the most used; Same Day was least common
4. **Profit Margin by Segment** — Corporate segment showed the strongest average margins; Home Office was slightly below average
5. **Refunded / Cancelled / Failed Orders by Region** — East had a disproportionate share of flagged orders relative to its order volume
6. **Monthly Sales Trend** — Sales showed a clear pickup in Q3 and Q4, with a slight dip in early months

Pivots 1 and 2 include sorting applied (by total sales, descending).

## Key Business Insights

- About 7% of total records had some form of data quality issue — not catastrophic, but enough to affect any summary report if left uncleaned
- Negative discounts likely stem from returns or data entry errors; they inflate reported sales figures if left unchecked
- Cancelled and failed-payment orders account for a meaningful chunk of the order count but contribute nothing to actual revenue — any top-line revenue number needs these excluded
- The East region shows an above-average rate of refunded/returned orders, which could be worth investigating operationally
- Three orders had ship dates before order dates — small number but indicative of a data entry or system sync issue worth flagging to whoever manages the source system

## Assumptions and Limitations

- Discount thresholds: there's no official upper-limit stated in the dataset. I used 50% as a warning threshold based on what's reasonable for a retail business. Records above this are flagged as warnings, not invalid.
- For conflicting duplicate records, I retained both copies and flagged them rather than picking one — there's no reliable way to determine which version is "correct" without source system access.
- Date parsing assumed that formats like `07/27/2024` are MM/DD/YYYY (US format), not DD/MM/YYYY — this was cross-checked against ship dates and order contexts to validate.
- The calculated_sales formula assumes discount is applied multiplicatively (`price × qty × (1 - discount)`), which is the standard retail formula. A few rows had mismatches in the original `sales` field — these were flagged in the quality report.
- This cleaning was done entirely in Excel. A production version of this process would ideally be scripted for reproducibility.

## Screenshots Included

| File | Contents |
|------|----------|
| `screenshots/raw_data_preview.png` | Raw dataset before any cleaning |
| `screenshots/cleaned_data_preview.png` | Cleaned dataset with all new calculated columns visible |
| `screenshots/pivot_summary_1.png` | Sales and Profit by Region pivot |
| `screenshots/pivot_summary_2.png` | Sales and Profit by Category and Sub-Category pivot |
