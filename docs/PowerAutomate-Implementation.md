# Power Automate Implementation Guide — A&H Lawn Care Invoicing

Four flows support the app. Build and test them in order — later flows assume earlier ones exist and work.

---

## Flow 1: Generate Document Number on Create

**Purpose:** Auto-populate `Title` (Doc Number) as `EST-####` or `INV-####` right after a new `Documents` item is created, using the SharePoint auto-incrementing `ID` for uniqueness (no separate counter list needed, no race conditions).

**Trigger:** `When an item is created` — SharePoint connector, site: your site, list: `Documents`.

**Steps:**
1. **Get item** (SharePoint) — Site: same site, List: `Documents`, Id: `ID` from trigger. (Ensures you have the freshly-committed `Doc Type` value; avoids trigger payload timing issues.)
2. **Condition** — `Doc Type` (from Get item) is equal to `Estimate`.
   - **If yes:** Compose `DocNumber` = `concat('EST-', formatNumber(triggerOutputs()?['body/ID'], '0000'))`
   - **If no:** Compose `DocNumber` = `concat('INV-', formatNumber(triggerOutputs()?['body/ID'], '0000'))`
3. **Update item** (SharePoint) — Site/List: `Documents`, Id: `ID`, Title: output of the Compose from step 2.

**Notes:**
- Using the list's internal `ID` guarantees no duplicate numbers even with concurrent creates — no need for a locking/counter pattern.
- If you later want numbers to reset per year (e.g. `EST-2026-0001`), swap the format expression to include `formatDateTime(utcNow(),'yyyy')`.
- Set flow run-after settings on Update item: run even `has failed` on Get item is not needed — default (only run after success) is fine.

---

## Flow 2: Recalculate Document Totals

**Purpose:** Whenever a line item is added, edited, or deleted, recompute the parent `Documents` record's `Subtotal` and `Total` (Total = Subtotal since there's no tax).

**Trigger:** `When an item is created, modified, or deleted` — SharePoint connector, list: `DocumentLines`. (Use the combined trigger, or three separate triggers pointing to the same child flow logic if your Automate environment doesn't support the combined trigger — combined is preferred to keep it to one flow.)

**Steps:**
1. **Condition** — check trigger type isn't relevant here since we need the parent ID either way. For **create/modify**, `Document` field is present on the triggering item. For **delete**, the deleted item's `Document` lookup value is still available in the trigger outputs (SharePoint retains it in the delete trigger payload) — capture it into a variable immediately:
   - **Initialize variable** `DocumentId` (Integer) = `triggerOutputs()?['body/Document/Id']` (fallback to `triggerOutputs()?['body/DocumentId']` depending on connector version — check the dynamic content picker for the exact lookup ID token).
2. **Get items** (SharePoint) — List: `DocumentLines`, Filter Query: `Document/Id eq {DocumentId}` (Odata), Select columns: `Amount`.
   - This requires the `Document` lookup column to be indexed (done in the SharePoint guide) so the filter query is efficient.
3. **Initialize variable** `RunningTotal` (Float) = `0`.
4. **Apply to each** item in Get items output:
   - **Set variable** `RunningTotal` = `add(variables('RunningTotal'), coalesce(items('Apply_to_each')?['Amount'], 0))`
5. **Update item** (SharePoint) — List: `Documents`, Id: `DocumentId`, Subtotal: `RunningTotal`, Total: `RunningTotal`.
   - Wrap this Update item in a **Condition**: only run if `DocumentId` is not null/empty (guards against edge cases where a line item is created before `Document` is set, e.g. via API).

**Concurrency setting:** On the trigger, set **Concurrency Control** to a degree > 1 (e.g. 20) so rapid successive line-item edits (adding several lines quickly in the app) don't queue up and slow down the UI feedback loop. Since updates are idempotent (recompute-from-scratch each time), concurrent runs are safe.

---

## Flow 3: Convert Estimate → Invoice

**Purpose:** Button-triggered from Power Apps. Clones an Estimate's header and line items into a new Invoice-type Document, links them, and marks the Estimate as `Invoiced`.

**Trigger:** `PowerApps (V2)` trigger. Add one input parameter: `DocumentId` (Number) — the ID of the Estimate to convert.

**Steps:**
1. **Get item** (SharePoint) — List: `Documents`, Id: `DocumentId` (the Estimate).
2. **Create item** (SharePoint) — List: `Documents`:
   - `Doc Type` = `Invoice`
   - `Customer` = Customer Id from step 1
   - `Bill To Snapshot` = from step 1
   - `Ship To Snapshot` = from step 1 (or leave blank if Estimates don't carry ship-to)
   - `Doc Date` = `utcNow()` (today, since this is when the invoice is generated)
   - `Due Date` = `addDays(utcNow(), 30)` (30-day terms; adjust default as needed)
   - `PO Number` = from step 1
   - `Status` = `Draft`
   - `Linked Document` = `DocumentId` (points back to the source Estimate)
   - `Notes` = from step 1
   - *(Title/Doc Number left blank — Flow 1's create-trigger will populate it automatically as `INV-####` since `Doc Type = Invoice` on this new item)*
3. **Get items** (SharePoint) — List: `DocumentLines`, Filter Query: `Document/Id eq {DocumentId}`.
4. **Apply to each** line from step 3:
   - **Create item** (SharePoint) — List: `DocumentLines`:
     - `Document` = new Invoice's Id (from step 2)
     - `Service`, `Item Label`, `Description`, `Qty`, `Rate`, `Amount`, `Sort Order` = copy directly from the current line
5. **Update item** (SharePoint) — List: `Documents`, Id: `DocumentId` (the original Estimate), `Status` = `Invoiced`, `Linked Document` = new Invoice's Id (from step 2) — cross-links both directions.
6. **Respond to PowerApps** — return `NewInvoiceId` (Number, from step 2's Id) so the app can navigate straight to the new invoice.

**Notes:**
- Flow 2 (totals recalculation) will fire automatically as each new `DocumentLines` item is created in step 4 — no need to manually copy Subtotal/Total; they'll self-populate. If you want to avoid N redundant recalculation runs during the copy, you can instead directly copy `Subtotal`/`Total` from the Estimate in step 2 as a shortcut, accepting Flow 2 will simply confirm the same value once the last line is created.

---

## Flow 4 (Optional): Generate & Email PDF

**Purpose:** Produce a PDF matching the paper layout and email it to the customer; update Status to `Sent`.

**Trigger:** `PowerApps (V2)` trigger, input: `DocumentId` (Number).

**Steps:**
1. **Get item** (SharePoint) — `Documents`, Id: `DocumentId`.
2. **Get items** (SharePoint) — `DocumentLines`, Filter: `Document/Id eq {DocumentId}`, sorted by `Sort Order` (use `$orderby` in the ODataQuery if available, or sort in a subsequent **Apply to each** using a Compose + `sort` array function).
3. **Populate a Word template** (Word Online / Power Automate action) using a `.docx` template stored in a SharePoint document library, with content controls for Customer, Dates, PO#, and a repeating table row bound to the line items collection.
   - Build the template first: create a Word doc styled like the sample Estimate/Invoice, insert content controls (Developer tab → Insert Controls) named to match the flow's field mapping, and save it to a `Templates` library in the same site.
4. **Convert Word document to PDF** (built-in Power Automate action, no separate connector needed).
5. **Send an email (V2)** (Office 365 Outlook connector) — To: `Customer.Email` (from a **Get item** on `Customers` using the Customer lookup Id), Subject: `"{Doc Type} {Doc Number} — A&H Lawn Care Services"`, Attachment: PDF from step 4.
6. **Update item** (SharePoint) — `Documents`, Id: `DocumentId`, `Status` = `Sent`.
7. **Respond to PowerApps** — return success boolean.

---

## Build & Test Order
1. Build and test Flow 1 in isolation — create a test Estimate and Invoice manually in SharePoint, confirm Doc Number populates correctly for both types.
2. Build and test Flow 2 — manually add/edit/delete DocumentLines rows against your test Document, confirm Subtotal/Total update correctly, including after a delete.
3. Build Flow 3 only once Power Apps has a working "Convert to Invoice" button to call it from (or test via Power Automate's built-in "Test" pane with a manually supplied `DocumentId`).
4. Build Flow 4 last, once the app and core flows are stable — it depends on a Word template that takes extra design time.

## Naming Convention
Name flows clearly for future maintenance:
- `DOC - Generate Document Number`
- `DOC - Recalculate Totals`
- `DOC - Convert Estimate to Invoice`
- `DOC - Generate and Send PDF`
