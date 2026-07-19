# 📊 Vendor Price Analytics Dashboard

**A spreadsheet-native pricing database system built entirely inside Google Sheets.**

The Vendor Price Analytics Dashboard tracks, normalizes, and analyzes vendor pricing across multiple markets over time. It behaves like a miniature relational database implemented with native Sheets formulas — using dimension tables, a fact table, a performance cache layer, a current-state snapshot, and an automated analytics/reporting layer.

> **Primary business question:**
> *"For Product X, among all vendors' CURRENT prices, what is the minimum, maximum, and average market price?"*

---

## Table of Contents

- [Project Overview](#project-overview)
- [System Architecture](#system-architecture)
- [Key Architectural Features](#key-architectural-features)
- [Performance Optimization Strategy](#performance-optimization-strategy)
- [Data Dictionary](#data-dictionary)
  - [`products`](#sheet-products)
  - [`markets`](#sheet-markets)
  - [`vendors`](#sheet-vendors)
  - [`price_log`](#sheet-price_log)
  - [`latest_stamps`](#sheet-latest_stamps)
  - [`current_snapshot`](#sheet-current_snapshot)
  - [`analytics`](#sheet-analytics)
- [Formula Engine](#formula-engine)
- [Data Entry Workflow](#data-entry-workflow)
- [Design Principles](#design-principles)
- [Future Extensions](#future-extensions)

---

## Project Overview

The workbook solves the problem of managing dynamic, multi-vendor pricing data where:

- The same **product** can have multiple **vendors**
- Each **vendor** belongs to a specific **market** / location
- Vendor prices **change periodically**
- **Historical prices must be preserved** — never overwritten
- **Current competitive pricing** must be summarized automatically

It combines:

- **Relational modeling** — dimension tables + foreign keys
- **Historical event logging** — an append-only fact table
- **Cached computation** — a precomputed lookup layer to keep formulas fast
- **Automated transformation pipelines** — `ARRAYFORMULA`, `MAP`, `LAMBDA`, `BYROW`
- **Dynamic reporting** — a self-updating analytics dashboard

---

## System Architecture

The workbook follows a simplified **star schema** data model, with `products` as the central dimension referenced by the transactional log, which in turn feeds a performance cache, a current-state snapshot, and the analytics layer.

```
                products
                    │
                    │
markets ──▶ vendors ──▶ price_log
                          │
                          │
                    latest_stamps
                          │
                          │
                  current_snapshot
                          │
                          │
                      analytics
```

---

## Key Architectural Features

### 1. Separation of Dimension Tables and Fact Tables

The workbook separates descriptive entities from transactional records, following a standard analytics engineering principle:

> **Store business entities once, then reference them through keys.**

| Table type | Sheets | Purpose |
|---|---|---|
| **Dimension tables** | `products`, `markets`, `vendors` | Reusable master/descriptive data |
| **Fact table** | `price_log` | Transactional price history |
| **Cache layer** | `latest_stamps` | Precomputed "latest date" lookups for performance |
| **Snapshot layer** | `current_snapshot` | Current-state view derived from the fact table |
| **Reporting layer** | `analytics` | Aggregated, product-level competitive metrics |

### 2. Dimension Tables

**`products`** — Master list of products (name, category, unit of measurement, status).

**`markets`** — Master list of geographical market locations, e.g.:

- Oyingbo
- Ojuelegba
- Mile 12

**`vendors`** — Maps vendors to their operating markets, e.g.:

```
Vendor A → Yaba
Vendor B → Lagos Island
Vendor C → Oyingbo
```

This allows market-level pricing analysis **without duplicating market data inside transactions**.

### 3. Fact Table — Immutable Historical Logging

**`price_log`** is the core transaction table. Every price update creates a **new row** — the system intentionally never overwrites a previous price.

| Product | Vendor | Price | Date |
|---|---|---|---|
| Rice | Vendor A | ₦50,000 | Jan 1 |
| Rice | Vendor A | ₦55,000 | Feb 1 |

Both records remain available. This enables:

- Historical price tracking
- Trend analysis
- Price movement detection
- Vendor comparison over time

### 4. Many-to-Many Relationships

- A single **product** can be sold by multiple **vendors**.
- A single **vendor** can sell multiple **products**.
- The system handles this seamlessly without data duplication.

### 5. Hierarchical Scoping

Vendors are strictly classified under regional `markets`, creating a multi-tiered filtering layer for downstream analytics.

---

## Performance Optimization Strategy

### The Problem

A naive approach to finding "the latest price" would require every row in `price_log` to repeatedly scan the entire table:

```
For every row:
    Find all matching products/vendors
    Find maximum date
    Compare current row
```

With *N* records, this becomes approximately:

> **O(N²)** complexity

As the log grows:

- Formulas slow down
- Recalculation becomes expensive
- Circular references become more likely

Checking the latest state directly from a live filtered snapshot also risks a **circular dependency loop** (`#REF!`), because the snapshot requires the `is_latest` flag, and that flag would otherwise require the snapshot.

### The Caching Strategy (`latest_stamps`)

Instead of repeatedly asking *"What is the latest date?"* on every row, the workbook asks **once**: *"Give me the latest date for each combination."* The result is stored as a single indexed lookup table.

```
[price_log raw data] ──▶ [latest_stamps (single QUERY cache)] ──▶ [price_log Column J (fast XLOOKUP)] ──▶ [current_snapshot]
```

1. **Isolate & aggregate** — `latest_stamps` executes a single, global `QUERY` across the raw data table, returning the maximum (most recent) `date_updated` per unique `product|vendor` combination.
2. **Point reference lookup** — Column J (`is_latest`) in `price_log` drops the heavy group-by search entirely, instead running a lightweight `XLOOKUP`/`VLOOKUP` against the single pre-calculated array on the helper sheet.
3. **Result** — computational complexity is reduced to **O(N)**, keeping the workbook responsive even with large datasets.

---

## Data Dictionary

### Sheet: `products`

**Purpose:** Master product dimension table.

| Column | Field | Type | Description |
|---|---|---|---|
| A | `product_id` | String (Key) | Unique, auto-generated product key |
| B | `product_name` | String | Product description |
| C | `category` | String | Product classification |
| D | `base_unit` | String | Measurement unit |
| E | `is_active` | Boolean | Product availability status |

**Product ID generation** (cell `A2`, expands automatically):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    "PR-" & TEXT(ROW(B2:B)-1, "000")
  )
)
```

Example output: `PR-001`, `PR-002`, `PR-003` …

---

### Sheet: `markets`

**Purpose:** Master market/location dimension table.

| Column | Field | Type | Description |
|---|---|---|---|
| A | `market_id` | String (Key) | Unique market identifier |
| B | `market_name` | String | Market location |

**Market ID generation** (cell `A2`):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    "MKT-" & TEXT(ROW(B2:B)-1, "000")
  )
)
```

---

### Sheet: `vendors`

**Purpose:** Vendor master table, mapped to its parent market.

| Column | Field | Type | Description |
|---|---|---|---|
| A | `vendor_id` | String (Key) | Unique, auto-generated vendor identifier (`V-001`, `V-002`, …) |
| B | `vendor_name` | String | Vendor name |
| C | `market_name` | String | Vendor location — dropdown sourced from `markets!B:B` |
| D | `market_id` | String (FK) | Linked market key |
| E | `is_active` | Boolean | Vendor status |

**Vendor ID generation** (cell `A2`):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    "V-" & TEXT(ROW(B2:B)-1, "000")
  )
)
```

**Market mapping (foreign key resolution)** — cell `D2`:

```excel
=ARRAYFORMULA(
  IF(C2:C1000="", "",
    XLOOKUP(
      C2:C1000,
      markets!B:B,
      markets!A:A
    )
  )
)
```

This automatically converts a typed market name like `Lagos` into its corresponding key, e.g. `MKT-001`.

---

### Sheet: `price_log`

**Purpose:** Central transaction history table — every price update is **appended** as a new record, never overwritten.

| Column | Field | Type | Description |
|---|---|---|---|
| A | `entry_id` | String (Key) | Unique transaction ID |
| B | `product_name` | String | Selected product — dropdown from `products!B:B` |
| C | `vendor_name` | String | Selected vendor — dropdown from `vendors!B:B` |
| D | `product_id` | String (FK) | Product foreign key |
| E | `vendor_id` | String (FK) | Vendor foreign key |
| F | `cost_price` | Numeric | Vendor cost |
| G | `selling_price` | Numeric | Market selling price |
| H | `date_updated` | Date/Time | Update timestamp |
| I | `combo_key` | String | Product/vendor relationship key |
| J | `is_latest` | Boolean | Current price indicator |
| K | `captured_by` | String | Data contributor |
| L | `notes` | String | Additional information |

**Entry ID** (cell `A2`):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    "ENT-" & TEXT(ROW(B2:B)-1, "000")
  )
)
```

**Product ID foreign key** (cell `D2`):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    XLOOKUP(
      B2:B,
      products!B2:B,
      products!A2:A,
      ""
    )
  )
)
```

**Vendor ID foreign key** (cell `E2`):

```excel
=ARRAYFORMULA(
  IF(C2:C="", "",
    XLOOKUP(
      C2:C,
      vendors!B2:B,
      vendors!A2:A,
      ""
    )
  )
)
```

**Composite key** — a unique vendor/product relationship key, e.g. `PR-001|V-002` (cell `I2`):

```excel
=ARRAYFORMULA(
  IF(D2:D="", "",
    D2:D & "|" & E2:E
  )
)
```

**`is_latest` flag** — resolved via fast lookup against the `latest_stamps` cache (cell `J2`):

```excel
=ARRAYFORMULA(
  IF(I2:I="", "",
    H2:H =
    VLOOKUP(
      I2:I,
      QUERY(
        {I2:I, H2:H},
        "select Col1, max(Col2)
         where Col1 is not null
         group by Col1
         label max(Col2) ''",
        0
      ),
      2,
      FALSE
    )
  )
)
```

> ⚠️ **Operational guardrail:** Never type over Columns A, D, E, I, or J — these contain structural array formulas. Overwriting them halts data flow into the analytics dashboard.

---

### Sheet: `latest_stamps`

**Purpose:** Calculation cache. Precomputes the latest timestamp for every product/vendor combination, so downstream sheets never need to scan the full `price_log` history.

Example output shape:

| combo_key | latest_date |
|---|---|
| PR-001\|V-001 | *(latest date)* |
| PR-001\|V-002 | *(latest date)* |

**Cache formula** (cell `A1`):

```excel
=QUERY(
  price_log!I2:H,
  "select Col1, max(Col2)
   where Col1 is not null
   group by Col1
   label max(Col2) ''",
  0
)
```

---

### Sheet: `current_snapshot`

**Purpose:** A clean, current-state table — the latest price per product/vendor combination only, instead of the entire historical log.

```
price_log
    │
    ▼
 latest only
    │
    ▼
current_snapshot
```

**Filter formula** (cell `A1`):

```excel
=FILTER(
  price_log!A:L,
  price_log!J:J = TRUE
)
```

This layer separates **historical data** (`price_log`) from **reporting data** (`current_snapshot`).

---

### Sheet: `analytics`

**Purpose:** Reporting dashboard layer. Provides product-level competitive pricing metrics, including:

- Average vendor cost
- Average selling price
- Cheapest vendor
- Highest / lowest market prices
- Vendor availability

| Column | Field | Description |
|---|---|---|
| A | `product_id` | Product foreign key |
| B | `product_name` | Dropdown sourced from `products!B:B` |
| C | `avg_cost_price` | Average cost price across active vendors |
| D | `avg_sell_price` | Average selling price across active vendors |
| E | `min_cost_price` | Minimum cost price |
| F | `min_sell_price` | Minimum selling price |
| G | `max_cost_price` | Maximum cost price |
| H | `max_sell_price` | Maximum selling price |
| I | `vendor_count` | Number of distinct active vendors |
| J | `cheapest_vendor_id` | ID of the cheapest current vendor |
| K | `cheapest_vendor_name` | Resolved name of the cheapest vendor |

**Product ID** (cell `A2`):

```excel
=ARRAYFORMULA(
  IF(B2:B="", "",
    XLOOKUP(
      B2:B,
      products!B:B,
      products!A:A,
      ""
    )
  )
)
```

**Average cost price** (cell `C2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          AVERAGE(
            FILTER(current_snapshot!F:F, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Average selling price** (cell `D2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          AVERAGE(
            FILTER(current_snapshot!G:G, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Minimum cost price** (cell `E2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          MIN(
            FILTER(current_snapshot!F:F, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Minimum selling price** (cell `F2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          MIN(
            FILTER(current_snapshot!G:G, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Maximum cost price** (cell `G2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          MAX(
            FILTER(current_snapshot!F:F, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Maximum selling price** (cell `H2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          MAX(
            FILTER(current_snapshot!G:G, current_snapshot!D:D = pid)
          ),
        0)
      )
    )
  )
)
```

**Vendor count** (cell `I2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          ROWS(
            UNIQUE(
              FILTER(current_snapshot!E:E, current_snapshot!D:D = pid)
            )
          ),
        0)
      )
    )
  )
)
```

**Cheapest vendor ID** (cell `J2`):

```excel
=ARRAYFORMULA(
  IF(A2:A="", "",
    MAP(A2:A,
      LAMBDA(pid,
        IFERROR(
          INDEX(
            SORT(
              FILTER(current_snapshot!E:F, current_snapshot!D:D = pid),
              2,
              TRUE
            ),
          1, 1),
        "")
      )
    )
  )
)
```

**Cheapest vendor name** (cell `K2`):

```excel
=ARRAYFORMULA(
  IF(J2:J="", "",
    IFERROR(
      VLOOKUP(J2:J, vendors!A:B, 2, FALSE),
    "")
  )
)
```

---

## Formula Engine

The workbook runs on a **hands-off maintenance architecture**: every major operation is calculated down the entire column using top-row array engines, so there's never a need to manually drag or copy formulas.

| Function | Role |
|---|---|
| **`ARRAYFORMULA`** | Lets a single formula automatically expand across an entire column, replacing "copy formula down 500 rows" with "one formula → entire column." |
| **`MAP` + `LAMBDA`** | Performs row-by-row calculations inside array pipelines, enabling dynamic per-row logic without manual dragging. |
| **`BYROW` + `LAMBDA`** | An alternative row-by-row iteration pattern used for the dashboard's aggregate metrics (average, min, max, cheapest vendor). |
| **`QUERY`** | Powers the `latest_stamps` cache via SQL-like aggregation (`group by`, `max`). |
| **`XLOOKUP` / `VLOOKUP`** | Resolves all foreign-key relationships and the fast `is_latest` cache lookup. |
| **`FILTER` / `SORT` / `UNIQUE` / `INDEX`** | Power the `current_snapshot` filter and the `analytics` aggregation formulas. |

---

## Data Entry Workflow

### Phase 1 — Onboarding New Structural Dimensions (Setup)

**Adding a new product**

1. Open `products`.
2. Add a new row with `product_name`, `category`, and `base_unit`.
3. `product_id` generates automatically.

**Adding a new market**

1. Open `markets`.
2. Enter the market name (e.g. `Lagos`).
3. `market_id` generates automatically.

**Adding a new vendor**

1. Open `vendors`.
2. Enter the vendor name (e.g. `Vendor A`).
3. Select the parent market from the Column C dropdown (e.g. `Lagos`).
4. `market_id` populates automatically.

### Phase 2 — Ongoing Pricing Logging Operations

**Logging a new price**

1. Open `price_log` and go to the first blank row at the bottom.
2. Select the **product** from the Column B dropdown.
3. Select the **vendor** from the Column C dropdown.
4. Enter `cost_price` and `selling_price` as **raw numbers** — do not include currency symbols (`$`, `₦`, etc.).
5. Enter the capture date under `date_updated`.
6. Optionally fill in `captured_by` and `notes`.

The system then automatically:

- Generates the `entry_id`
- Maps the `product_id`
- Maps the `vendor_id`
- Creates the `combo_key` relationship key
- Determines the `is_latest` price status

> ⚠️ **Critical operational guardrail:** Never type over columns containing structural formulas (Columns A, D, E, I, J in `price_log`). Overwriting these cells destroys the array formulas and halts data flow to the analytics dashboard.

---

## Design Principles

### Single Source of Truth

- Each entity exists exactly once.
- Products are not duplicated.
- Markets are not manually re-typed across rows.

### Append-Only History

- Historical prices are preserved.
- Old prices are **never** updated in place.
- New records are always **inserted**.

### Layered Architecture

Each sheet has exactly one responsibility:

| Layer | Sheet(s) | Responsibility |
|---|---|---|
| **Dimensions** | `products`, `markets`, `vendors` | Master data |
| **Fact table** | `price_log` | Transactions |
| **Cache** | `latest_stamps` | Performance |
| **Snapshot** | `current_snapshot` | Current state |
| **Analytics** | `analytics` | Reporting |

---

## Future Extensions

Possible improvements under consideration:

- Price trend charts
- Market inflation tracking
- Vendor reliability scores
- Automated alerts when prices change
- Google Forms ingestion for data capture by supplier reps
- BigQuery migration
- Looker Studio reporting

---

## Summary

This workbook is a **relational pricing tracker** built entirely in Google Sheets. It combines relational modeling, historical event logging, cached computation, automated transformation pipelines, and dynamic reporting to create a scalable vendor price intelligence platform.
