# Configuration walkthrough

Everything in `name-to-resource-linker` is controlled by constants at the top of
`src/NameToResource.bas`. This page shows the defaults and then walks through a
completely different domain to prove the tool is domain-agnostic.

## Defaults (as shipped)

```vb
SRC_SHEET        = ""                 ' active sheet
NAME_COL         = "C"
HEADER_ROW       = 3
DATA_START_ROW   = 4
NAME_HEADER_TEXT = "Item name"

RESOURCE_SHEETS  = "Resources,ResourcesExtra"

RES_NAME_COL     = "A"                ' key name
RES_DISPLAY_COL  = "B"                ' display / alias name
RES_CODE_COL     = "C"                ' optional code
RES_RATE_COL     = "D"                ' rate (numeric -> percentage)
RES_PRICE_COL    = "E"                ' price
RES_WEIGHT_COL   = "F"                ' weight / measure
RES_LINK_COL     = "G"                ' URL
RES_CATEGORY_COL = "H"                ' category (output only)
RES_GROUP_COL    = "I"                ' group (output only)
RES_METRIC_COL   = "J"                ' numeric sort metric

OUTPUT_LINK_COL_START = "AG"
MAX_LINKS       = 3
HYPERLINK_ON    = True
FORMAT_ON       = True

EXCLUDE_LINK_KEYWORDS = "out of stock,discontinued,sold out"
EXCLUDE_SUFFIX_WORDS  = "accessory,kit,cover,base,cap,decor,part"
COMPOSITE_SEP   = "="
```

## Worked example: ticket summary → knowledge-base articles

Suppose you receive support tickets and want each ticket's short summary linked
to the top-3 most relevant KB articles, ranked by a helpfulness score.

### 1. Source sheet (`Tickets`)

| C = Summary (row 3 header `Item name`) | rows 4+ |
|----------------------------------------|---------|
| Password reset not working | … |
| Export to CSV fails on Safari | … |

### 2. Resource sheets (in the macro workbook)

Each resource sheet uses the same A..J layout, but the fields are reinterpreted:

- `A` ResourceName — a keyword/phrase that a ticket summary might contain
  (e.g. `Password reset`, `CSV export`, `Safari`).
- `B` DisplayName — article title.
- `C` Code — article ID (e.g. `KB-1201`).
- `D` Rate — article rating (numeric → shown as `%`).
- `E` Price — leave blank or repurpose as "views".
- `F` Weight — leave blank or repurpose as "age in days".
- `G` Link — the article URL.
- `H` Category — topic (output only).
- `I` Group — product area (output only).
- `J` Metric — helpfulness score used for ranking (higher = better, so store
  `1 / score` or invert the sort by negating).

An entry might be:

```
Password reset | Password Reset Guide | KB-1201 | 96 |  |  | https://kb.example.com/1201 | Account | Core | 0.04
```

(Here `Metric = 0.04` = `1 / 25` helpfulness, so the best-rated article sorts
first. If you prefer "higher is better", set the metric directly and flip the
comparison in `SortIndexesByMetric`.)

### 3. Edit the constants

```vb
NAME_COL         = "C"
HEADER_ROW       = 3
DATA_START_ROW   = 4
NAME_HEADER_TEXT = "Item name"

RESOURCE_SHEETS  = "KBArticles"

OUTPUT_LINK_COL_START = "AG"
MAX_LINKS       = 3

EXCLUDE_LINK_KEYWORDS = "deprecated,draft"
EXCLUDE_SUFFIX_WORDS  = "guide,faq,template"
```

Notes:

- `RESOURCE_SHEETS` accepts a comma-separated list, so you can layer multiple
  article libraries (e.g. `KBArticles,CommunityArticles`) and the index merges
  them by name.
- A ticket summary `Password reset not working` contains `Password reset`, so it
  matches the KB entry exactly (the item appears at the **start** of the summary,
  and the entry is an exact name match). A summary `Safari CSV export bug` would
  match both `CSV export` and `Safari` entries.
- `EXCLUDE_SUFFIX_WORDS` here prevents `Password reset` from over-matching
  `Password reset guide` as if "guide" were part of the core term.

### 4. Result

Each ticket row gets up to three article links in columns AG/AH/AI, cheapest (or
highest-scored) first, each a clickable hyperlink of the form:

```
Password Reset Guide=KB-1201=96.00%=N/A=https://kb.example.com/1201
```

Same engine, different vocabulary — no code changes beyond the constants.

## Tuning the matcher

- Raise `MAX_LINKS` (to `5`) to surface more candidates per item.
- Set `HYPERLINK_ON = False` to write only the composite text (no clickable link).
- Set `FORMAT_ON = False` to skip header/column styling when you only want data.
- Add your own keywords to `EXCLUDE_LINK_KEYWORDS` to drop stale or banned links.
- Add your own over-broad terms to `EXCLUDE_SUFFIX_WORDS` to stop generic
  suffixes from swallowing a more specific source term.
