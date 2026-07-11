# Data Dictionary

Release **v0.2.0-alpha.3** defines the production category dictionary in
`config/categories.yaml`.

## Approved merchant defaults

| Workbook label | Normalized category | Master category |
|---|---|---|
| Costco | Groceries | Food |
| Canadian Tire | Household Supplies | Household |
| Walmart | Household Supplies | Household |
| Amazon | Household Supplies | Household |

## Behaviour

- Matching is case-insensitive.
- Curly apostrophes and dash variants are normalized.
- Ordinal due dates such as `(21st)` are removed for matching.
- Unknown labels remain unmapped and must be reviewed; they are not guessed.

Configured normalized categories: **64**.
