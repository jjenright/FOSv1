# FOS v1.1.1 — Visa transaction integration

FOS now recognizes a copied `Import_2025`-style Visa sheet inside the private source workbook. Detection is based on the transaction-table headers, so the sheet may be named `Import_2025`, `Visa Transactions 2025`, or another descriptive name.

Imported rows must include a transaction date, amount, transaction type, and either a valid FOS `CategoryID` or a recognizable suggested category. Credit-card payments are excluded as transfers. Purchases, refunds, interest, and fees are merged into the matching annual year without changing the annual pay-period count.

The integration assumes the detailed Visa purchases were not also entered separately in the annual sheet. FOS can prevent duplicate transaction IDs within the Visa table, but it cannot reliably identify a manually entered annual expense as the same purchase when the annual row has no transaction ID.
