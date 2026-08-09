# SharePoint Implementation Guide — A&H Lawn Care Invoicing

Step-by-step instructions for standing up the SharePoint site and lists that back the Power App. Build lists in the order below — later lists depend on lookups into earlier ones.

## 0. Prerequisites
- A SharePoint site (Team or Communication site) dedicated to this app, e.g. `AH-LawnCare-Ops`.
- Modern list experience enabled (default in SharePoint Online).
- You as sole owner/full-control member (single-user app — no additional permission groups needed).

## 1. Build Order
1. `Customers`
2. `Services`
3. `Documents`
4. `DocumentLines`

(Lookups require the target list to exist first, hence Customers/Services before Documents, and Documents before DocumentLines.)

---

## 2. List: `Customers`

Create list → name it **Customers**. Rename the default `Title` column display name to **Customer Name** (internal name stays `Title`).

| Column (display name) | Internal type | Settings |
|---|---|---|
| Customer Name | Title (rename display) | Required |
| Property Name | Single line text | Optional — e.g. "Don Kennedy Property" |
| Billing Address | Multiple lines of text | Plain text, no versioning history needed |
| City | Single line text | |
| State | Single line text | |
| Zip | Single line text | **Text, not Number** — preserves leading zeros/formatting |
| Contact Name | Single line text | |
| Phone | Single line text | |
| Email | Single line text | (use text, not the built-in hyperlink type, for simpler formulas) |
| Property Number | Single line text | Internal account/property reference |
| Notes | Multiple lines of text | Plain text |
| Active | Yes/No | Default value: Yes |

**View:** Default view — add columns Customer Name, Property Name, City, Phone, Active. Sort by Customer Name ascending. Add a filter toggle view "Active Customers" (`Active = Yes`).

---

## 3. List: `Services`

Create list → **Services**. Rename `Title` display to **Service Name**.

| Column (display name) | Internal type | Settings |
|---|---|---|
| Service Name | Title (rename) | Required |
| Category | Choice | Values: `Lawn Care`, `Landscaping`, `Irrigation`, `Labor`, `Materials`, `Other`. Default: none. Format as dropdown. |
| Unit | Choice | Values: `Each`, `Hour`, `Yard`, `Sq Ft`, `Visit`, `Month`. Dropdown. |
| Default Rate | Currency | 2 decimal places, USD |
| Description | Multiple lines of text | Plain text — boilerplate wording to pre-fill line item descriptions |
| Active | Yes/No | Default: Yes |

**View:** Default view grouped by Category, showing Service Name, Unit, Default Rate, Active.

---

## 4. List: `Documents`

Create list → **Documents**. Rename `Title` display to **Doc Number** (this will be populated by a flow, not typed by users).

| Column (display name) | Internal type | Settings |
|---|---|---|
| Doc Number | Title (rename) | Populated by flow after creation — see Power Automate guide |
| Doc Type | Choice | Values: `Estimate`, `Invoice`. Required. No default (force selection on New). Radio buttons recommended for form clarity. |
| Customer | Lookup | Source list: `Customers`, show column `Title` (Customer Name). Required. Allow single value only. |
| Bill To Snapshot | Multiple lines of text | Optional — frozen billing address text at time of send |
| Ship To Snapshot | Multiple lines of text | Optional — used for Invoices only |
| Doc Date | Date only | Required. Default: Today |
| Due Date | Date only | Optional — Invoices only |
| PO Number | Single line text | Optional |
| Status | Choice | Values: `Draft`, `Sent`, `Approved`, `Invoiced`, `Paid`, `Void`. Default: `Draft`. |
| Linked Document | Lookup | Source list: `Documents` (self-lookup), show column `Title`. Used when converting an Estimate to an Invoice. **Note:** SharePoint allows self-referencing lookups; create this column after the list already has at least one item, or SharePoint may block it until the list exists with data — if you hit this, create the column, save the list, then add the first Estimate. |
| Subtotal | Currency | 2 decimals. **Do not** mark as calculated column — it's set by a Power Automate flow reading `DocumentLines`. |
| Total | Currency | 2 decimals. Same value as Subtotal (no tax) — set by the same flow, kept as a separate column for clarity/future-proofing if tax is ever reintroduced. |
| Notes | Multiple lines of text | Footer/terms text shown on the printed document |
| Prepared By | Person or Group | Default: current user; single value |

**Indexing:** Add an index on `Doc Type` and `Status` (List Settings → Indexed columns) — keeps filtered views/galleries fast as the list grows.

**Views:**
- **All Estimates** — filter `Doc Type = Estimate`, sort by Doc Date desc.
- **All Invoices** — filter `Doc Type = Invoice`, sort by Doc Date desc.
- **Open Invoices** — filter `Doc Type = Invoice AND Status ne Paid AND Status ne Void`.

---

## 5. List: `DocumentLines`

Create list → **DocumentLines**. Rename `Title` display to **Line Label** (optional short label; can stay unused/hidden in forms since `Item Label` below is the real field).

| Column (display name) | Internal type | Settings |
|---|---|---|
| Line Label | Title (rename) | Optional — leave default text on create, not shown on the Power Apps line-item UI |
| Document | Lookup | Source list: `Documents`, show column `Title`. Required. This is the parent/header link. |
| Service | Lookup | Source list: `Services`, show column `Title`. Optional. |
| Item Label | Single line text | e.g. "Annual Maintenance Contract", "Landscaping" — required |
| Description | Multiple lines of text | Plain text — free-form detail matching the wrapped description text seen on paper estimates |
| Qty | Number | 2 decimals, no thousands separator needed |
| Rate | Currency | 2 decimals |
| Amount | Currency | 2 decimals — set via Power Apps formula (`Qty * Rate`) when the row is saved, **not** a SharePoint calculated column |
| Sort Order | Number | Integer, no decimals — controls display order |

**Indexing:** Add an index on `Document` (the lookup) — this list will be queried/filtered by parent Document constantly from both the app and the Power Automate totals flow.

**View:** Default view sorted by `Document`, then `Sort Order` ascending.

---

## 6. Permissions
Single-user app — keep it simple:
- Site collection: you as Owner.
- No SharePoint groups needed beyond the default Owners group.
- Turn off "Allow items from this list to appear in search results" is optional; not required for a single-user tool.

## 7. Post-Build Checklist
- [ ] All 4 lists created in the correct order
- [ ] All lookup columns point to the correct target list/column
- [ ] `Doc Type`, `Status` indexed on `Documents`
- [ ] `Document` lookup indexed on `DocumentLines`
- [ ] Choice values match exactly what's referenced in the Power Automate and Power Apps guides (case-sensitive matching in formulas)
- [ ] Default views created as described
- [ ] Test data: add 1 Customer, 2–3 Services, then manually add 1 Document + 2 DocumentLines rows to confirm lookups resolve correctly before moving to Power Automate/Power Apps build-out
