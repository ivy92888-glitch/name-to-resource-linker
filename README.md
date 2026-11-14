# name-to-resource-linker

A small, dependency-free **VBA macro for Excel-compatible spreadsheets** that,
for each free-text item name in a source sheet, finds the best-matching
resource records from one or more configurable reference tables and writes up
to N ranked reference links (with a clickable hyperlink) back into the sheet.

> **Unofficial.** Not affiliated with, endorsed by, or sponsored by Microsoft
> (Excel) or any other spreadsheet software vendor. The word "workbook" is used
> only as a generic reference to a spreadsheet file.

## Why this exists

Many everyday tasks share the same shape: you have a list of messy, free-text
names and you want each one linked to the **best** supporting resource — a
document, a supplier page, a knowledge-base article, a spec sheet — ranked by
some metric such as price or relevance. Examples that fit this exact pattern:

- attaching the cheapest supplier link to each requested item,
- suggesting the top knowledge-base articles for each ticket,
- linking the most relevant spec sheet to each product in an order.

Doing this by hand is slow and inconsistent. This macro automates it with a
robust matching engine and writes the result straight back into the sheet.

## What the engine does

The macro is fully generic. It implements seven reusable ideas:

1. **Name-keyed resource registry.** Each reference name maps to a list of
   resource records, where every record carries several attributes plus a URL.
2. **Two-tier name matching.** An exact match is preferred; a boundary-aware
   fuzzy match (so `leg` does **not** match `legacy`) is used as a fallback.
3. **Link filtering.** Records whose link contains any excluded keyword
   (e.g. "out of stock") are dropped before ranking.
4. **Cross-field normalization.** A numeric rate is coerced to a percentage
   string; invalid metric values are pushed to the end of the sort.
5. **Metric-based ranking.** Each name's records are sorted by a numeric metric
   (e.g. lowest unit price first).
6. **Top-N priority blend.** Exact matches fill first, fuzzy matches fill the
   rest, capped at `MAX_LINKS` (default 3).
7. **Composite output + hyperlink.** A readable attribute string
   (`DisplayName=Code=Rate=Weight=Link`) plus a clickable hyperlink is written
   per link.

All of the above is driven by configuration constants at the top of the module;
nothing about a specific domain is hard-coded.

## How it works

1. The macro reads the source sheet (the active sheet, or a named one).
2. For each data row it takes the item name.
3. It scans the configured resource sheets, building a name → records index,
   applying the link filter and sorting each name's records by metric.
4. For each item it collects exact and fuzzy hits, blends them (exact first,
   capped at `MAX_LINKS`), and writes the composite string + hyperlink.
5. Rows with no match get cleared output cells.

### Illustrative example (fictional)

With the bundled sample, the source item `Desk leg` matches:

- `Desk leg` **exact** (two records, sorted by metric → cheapest first),
- `X-shaped desk leg` (intelligent: item appears at the **end** of the name),
- `Desk leg-black` (intelligent: item at the **start** with a `-` suffix),

but it does **not** match `Desk leg kit`, because `kit` is in the generic-suffix
blacklist and the user input did not contain it.

## Input data format

The source sheet layout is configurable. Defaults used by the sample:

| Column (default) | Meaning | Required |
|------------------|---------|----------|
| C (NAME_COL)     | Item name (free text) | Yes |
| C3 (header)      | Expected header text `Item name` | No (configurable) |

- `HEADER_ROW` = `3`, `DATA_START_ROW` = `4`.
- Resource tables live in the **same workbook that holds the macro**
  (`ThisWorkbook`), each with a header row and data starting at row 2.
- A resource name may use `-` to drop a suffix and `/` to list aliases; each
  alias is matched independently.

### Resource table columns (A..J)

| Col | Field | Role |
|-----|-------|------|
| A | ResourceName | key name (`/` aliases, `-` suffix) |
| B | DisplayName | human-readable label |
| C | Code | optional identifier |
| D | Rate | numeric rate → percentage |
| E | Price | generic price |
| F | Weight | generic weight / measure |
| G | Link | URL (filtered by keyword) |
| H | Category | output only |
| I | Group | output only |
| J | Metric | numeric sort key |

## Configuration

All knobs are constants at the top of `src/NameToResource.bas`:

| Constant | Default | Meaning |
|----------|---------|---------|
| `SRC_SHEET` | `""` | `""` = active sheet, else the sheet name |
| `NAME_COL` | `"C"` | Item name column |
| `HEADER_ROW` | `3` | Row holding the name header |
| `DATA_START_ROW` | `4` | First data row |
| `NAME_HEADER_TEXT` | `"Item name"` | Header to expect (`""` to skip) |
| `RESOURCE_SHEETS` | `"Resources,ResourcesExtra"` | Comma-separated resource sheets |
| `RES_NAME_COL` … `RES_METRIC_COL` | `A` … `J` | Resource table column letters |
| `OUTPUT_LINK_COL_START` | `"AG"` | First output column |
| `MAX_LINKS` | `3` | Max links per item |
| `HYPERLINK_ON` | `True` | Attach clickable links |
| `FORMAT_ON` | `True` | Style output columns |
| `EXCLUDE_LINK_KEYWORDS` | `"out of stock,discontinued,sold out"` | Links containing these are skipped |
| `EXCLUDE_SUFFIX_WORDS` | `"accessory,kit,cover,base,cap,decor,part"` | Over-broad name suffixes |
| `COMPOSITE_SEP` | `"="` | Field separator in output string |

### Matching rules at a glance

- **Exact**: `dbName` equals the item, or one of its `/` aliases does.
- **Intelligent**: `dbName` *contains* the item, and the position check passes:
  - item at the **end** of the name → OK,
  - item at the **start** + `-suffix` → OK,
  - item at the **start** + extra word → NO,
  - item in the **middle** → NO.
- Plus a **suffix blacklist** so `Desk leg` won't match `Desk leg kit`.

## Installation

1. Open your workbook in Excel or any compatible spreadsheet program.
2. Press `Alt + F11` to open the VBA editor, then `File ▸ Import File…` and
   select `src/NameToResource.bas`.
3. Place your resource tables in the **same workbook** that now holds the macro.
4. Edit the configuration constants at the top of the module to match your layout.
5. Save the workbook as a macro-enabled file (`.xlsm`).

## Usage

1. Prepare the source sheet in the [input format](#input-data-format) above.
2. Select the sheet you want to process (or leave `SRC_SHEET = ""` for the active
   sheet).
3. Run the `LinkResources` macro.
4. Review the links in the output columns (AG..AI by default); adjust manually if
   needed.

A ready-to-run sample lives in
[`examples/sample-data.xlsx`](examples/sample-data.xlsx).

## File structure

```
name-to-resource-linker/
├── src/
│   └── NameToResource.bas    # the macro (clean-room, generic reimplementation)
├── docs/
│   └── Configuration.md      # configuration walkthrough with a second domain
├── examples/
│   └── sample-data.xlsx      # neutral, fictional sample workbook
├── LICENSE                   # MIT
├── README.md                 # this file
├── PUSH_GUIDE.md             # how this repo was published to GitHub
├── .gitignore
└── .gitattributes
```

## Performance notes

Indexing scans each resource sheet once; linking scans the index per source row.
For typical lookup tables (thousands of rows) this is fast. Each name's record
list is sorted independently, so cost grows roughly with (source rows) × (distinct
matched names) plus the per-name sort.

## Limitations

- Resource tables must reside in the same workbook as the macro (`ThisWorkbook`).
- The engine matches on whole-name equality or boundary-aware substring; it does
  not do fuzzy/phonetic matching.
- The final candidate list fills exact matches first then fuzzy, in index order;
  it is not globally re-sorted across different matched names.
- This macro was not executed in the author's environment; test it on a copy of
  your data before relying on it.

## Non-association & Legal

This repository is an **independent, clean-room reimplementation** of a generic
matching pattern. It contains no proprietary data, no internal nomenclature, and
no business-specific logic from any third party or former employer. The author
is not affiliated with, and does not speak for, any company whose data may have
originally inspired a similar workflow.

The MIT license covers only the author's own contributions; it does not purport
to license any rights of others. Users are responsible for ensuring their use of
this tool complies with applicable laws and their own organization's policies.

## License

Released under the [MIT License](LICENSE). Copyright (c) 2026 ivy92888-glitch.
