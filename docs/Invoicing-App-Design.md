# Lawn Care Invoicing & Estimates App — Design Document

## 1. Overview
A simple Power Apps (Canvas app) backed by SharePoint lists to manage **Customers**, a **Service/Price list**, and **Estimates/Invoices** (with line items). Built around SharePoint **Form controls** for fast, low-maintenance CRUD, with a lightweight editable gallery for line items (the one place a plain form can't handle repeating rows).

Based on the sample InvoiceEstimates provided:
- **Estimate #148** — Bill To (Customer + Property name), Date, line items (Item / Description / Qty / Rate / Amount), Subtotal, Tax %, Total.
- **Invoice #2835** — Invoice #, Date, Due Date, Bill To, Ship To, P.O. Number, line items, Subtotal, Tax %, Total.

---

## 2. SharePoint List Schema

### 2.1 `Customers`
| Column | Type | Notes |
|---|---|---|
| Title | Single line text | Customer/Company name (e.g. "Gabriel Properties, LLC") |
| PropertyName | Single line text | e.g. "Don Kennedy Property" — optional site/property label |
| BillingAddress | Multiple lines (plain) | Street |
| City | Single line text | |
| State | Single line text | |
| Zip | Single line text | Keep as text (leading zeros, formatting) |
| ContactName | Single line text | |
| Phone | Single line text | |
| Email | Single line text | |
| PropertyNumber | Single line text | Internal property/account ref (e.g. "422505") |
| Notes | Multiple lines | |
| Active | Yes/No | Default Yes; hide inactive customers from pickers |

### 2.2 `Services` (price list)
| Column | Type | Notes |
|---|---|---|
| Title | Single line text | Service name ("Weekly Lawn Cutting", "Mulch & Pre-emergent", "Irrigation Labor") |
| Category | Choice | Lawn Care / Landscaping / Irrigation / Labor / Materials / Other |
| Unit | Choice | Each, Hour, Yard, Sq Ft, Visit, Month |
| DefaultRate | Currency | Pre-fills line item rate; still editable per line |
| Description | Multiple lines | Optional boilerplate description |
| Active | Yes/No | Default Yes |

### 2.3 `InvoiceEstimates` (Estimate/Invoice header — one list for both types)
| Column | Type | Notes |
|---|---|---|
| Title | Single line text | Doc number, e.g. "EST-0148" / "INV-2835" (set by Power Automate on create) |
| DocType | Choice | Estimate / Invoice |
| Customer | Lookup → Customers | |
| BillToSnapshot | Multiple lines | Optional: frozen copy of billing address at time of send |
| ShipToSnapshot | Multiple lines | Invoices only; optional |
| DocDate | Date only | |
| DueDate | Date only | Invoices only |
| PONumber | Single line text | |
| Status | Choice | Draft / Sent / Approved / Invoiced / Paid / Void (Approved & Invoiced apply to Estimates) |
| LinkedDocument | Lookup → InvoiceEstimates | When an Estimate is converted to an Invoice, link back for traceability |
| Subtotal | Currency | Calculated (Flow or Power Apps) from line items |
| Total | Currency | Same as Subtotal — no tax tracked (services, not taxable products) |
| Notes | Multiple lines | Footer/terms text, e.g. "Any extra work will be approved & billed..." |
| PreparedBy | Person | Defaults to current user (Harold) |

> Note: SharePoint calculated columns can't reference another list, so Subtotal/Total are best set by a small Power Automate flow (or Power Apps `Patch`) whenever line items change — see §4.

### 2.4 `DocumentLines`
| Column | Type | Notes |
|---|---|---|
| Title | Single line text | Optional short item name |
| Document | Lookup → InvoiceEstimates | Parent header |
| Service | Lookup → Services | Optional — pulls DefaultRate/Unit on selection |
| ItemLabel | Single line text | e.g. "Annual Maintenance Contract", "Landscaping" |
| Description | Multiple lines | Free text detail (matches the wrapped descriptions in your samples) |
| Qty | Number | |
| Rate | Currency | Editable even if Service selected |
| Amount | Currency | Calculated: Qty × Rate (set via app formula, not SP calc column) |
| SortOrder | Number | Controls display order in the gallery/PDF |

---

## 3. Power Apps Structure

### Screens
1. **HomeScreen** — Dashboard: counts of Draft estimates, Unpaid invoices; buttons to "New Estimate", "New Invoice", browse Customers/Services.
2. **CustomerListScreen / CustomerFormScreen** — Standard SharePoint **Form control** (New/View/Edit) bound to `Customers`.
3. **ServiceListScreen / ServiceFormScreen** — Standard **Form control** bound to `Services`.
4. **DocumentListScreen** — Gallery of `InvoiceEstimates`, filterable by DocType/Status/Customer.
5. **DocumentDetailScreen** — **Form control** for the header fields (Customer, Date, PO#, Status, Notes) + a **line-items section** below:
   - Editable/scrollable gallery bound to `Filter(DocumentLines, Document.ID = ThisItem.ID)`, sorted by SortOrder.
   - Each row: Service combo box (optional), ItemLabel/Description text inputs, Qty, Rate, read-only Amount label.
   - "+ Add Line" button → `Patch(DocumentLines, Defaults(DocumentLines), {Document: header, SortOrder: count+1})`.
   - Delete icon per row → `Remove()`.
   - Footer label for Total computed live with `Sum()` over the filtered lines (and pushed to the header record on Save via Patch).
6. **PDF/Print Screen** (optional) — A print-styled screen matching your paper layout, exported via `PDF Export` or Power Automate + Word template for emailing to customers.

### Why hybrid instead of pure Form control everywhere
The built-in Form control is 1 form = 1 list item, so it's perfect for Customers, Services, and the Document *header*. But an invoice needs an unknown number of line items, which forms can't repeat — that's inherently a one-to-many relationship, so a small custom gallery + Patch pattern is the standard, low-code way to handle it while keeping everything else simple.

---

## 4. Power Automate Flows (keep app logic thin)
1. **On InvoiceEstimates create** — Generate sequential Title (EST-#### / INV-####) per DocType.
2. **On DocumentLines create/update/delete** — Recalculate parent Document's Subtotal/Total (no tax to compute).
3. **Convert Estimate → Invoice** (button-triggered) — Clone header + lines into a new Invoice-type Document, set LinkedDocument, set Estimate Status = Invoiced.
4. **(Optional) Send Invoice/Estimate PDF** — Generate PDF (Word template or Power Apps PDF export) and email to Customer.Email; update Status to Sent.

---

## 5. Build Order (suggested)
1. Create the 4 SharePoint lists per schema above.
2. Build Customers & Services screens (pure Form controls) — quick wins, validates the pattern.
3. Build InvoiceEstimates list screens: gallery + header Form control.
4. Add DocumentLines gallery/Patch logic on the detail screen.
5. Add the numbering + totals-calculation flows.
6. Add PDF export / email flow.
7. (Later) Convert-to-Invoice flow once basic CRUD is solid.

---

## 6. Confirmed Decisions
- **Single `InvoiceEstimates` list** for both Estimates and Invoices, distinguished by `DocType`. Numbering (EST-#### / INV-####) handled per-type by the create flow (§4.1).
- **No tax tracking** — these are services, not taxable products, so `TaxRatePct`/`TaxAmount` are omitted; `Total` simply mirrors `Subtotal`.
- **Single-user app** — no role-based security, sharing, or per-crew filtering needed. Standard SharePoint list permissions (owner-only edit) are sufficient.
