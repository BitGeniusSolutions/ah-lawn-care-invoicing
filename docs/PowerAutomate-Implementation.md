# Power Automate Implementation Guide — A&H Lawn Care Invoicing

Four flows support the app. Build and test them in order — later flows assume earlier ones exist and work.

## Conventions Used Throughout This Guide

All SharePoint list read/write actions in these flows use **Send an HTTP request to SharePoint** instead of the native `Get item(s)`/`Create item`/`Update item` actions (triggers are unaffected — they remain the native SharePoint trigger actions). This avoids `Update item`'s full-row required-field re-validation, gives direct control over `$filter`/`$select`/`$expand`, and is unaffected by lists being renamed later since it addresses lists by GUID.

**Environment variables** (create these once, referenced by GUID across all flows):
- `SiteUrl` — the site's absolute URL
- `CustomersListId` — GUID of the `Customers` list
- `ServicesListId` — GUID of the `Services` list
- `InvoiceEstimatesListId` — GUID of the `InvoiceEstimates` list
- `DocumentLinesListId` — GUID of the `DocumentLines` list

> Get each GUID from List Settings → scroll to the bottom → the `List=%7B...%7D` segment of the URL (URL-decode the `%7B`/`%7D` to `{`/`}` and drop the braces), or via `_api/web/lists/GetByTitle('ListName')?$select=Id` in a browser while signed in.

**Standard headers by method** (paste directly into the action's Headers field — it accepts JSON):

GET (read one/many items):
```json
{
  "Accept": "application/json;odata=nometadata"
}
```

POST (create an item):
```json
{
  "Accept": "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata"
}
```

PATCH (update an item):
```json
{
  "Accept": "application/json;odata=nometadata",
  "Content-Type": "application/json;odata=nometadata",
  "IF-MATCH": "*"
}
```

**Reading responses:** The HTTP action returns a raw JSON string body — follow every HTTP action with a **Parse JSON** action (`Content: body('<HTTP action name>')`) so downstream steps can reference fields normally instead of via manual string parsing. Generate each schema from a sample response the first time you run the action in test mode.

**Lookup fields in REST:** A SharePoint lookup column named e.g. `Customer` exposes its id in REST as `CustomerId` (SharePoint appends `Id` to the internal name) — use that suffixed name in `$select`, `$filter`, and JSON request bodies. Confirm exact internal names via `_api/web/lists(guid'...')/fields?$select=InternalName,Title`. **Choice fields** (like `DocType`, `Status`) return as a plain string in REST (e.g. `"DocType": "Estimate"`) — no `/Value` needed, unlike the native connector's output shape.

---

## Flow 1: Generate Document Number on Create

**Purpose:** Auto-populate `Title` (Doc Number) as `EST-####` or `INV-####` right after a new `InvoiceEstimates` item is created, using the SharePoint auto-incrementing `ID` for uniqueness (no separate counter list needed, no race conditions).

**Trigger:** `When an item is created` — SharePoint connector, site: your site, list: `InvoiceEstimates`.

**Steps:**
1. **Send an HTTP request to SharePoint** — use this instead of **Get item** to fetch the freshly-committed `DocType` value (avoids trigger payload timing issues).
   - **Site Address:** your site (or `variables('SiteUrl')`)
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{triggerOutputs()?['body/ID']})?$select=Id,DocType`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
2. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint')`, Schema:
   ```json
   {
     "type": "object",
     "properties": {
       "Id": { "type": "integer" },
       "DocType": { "type": "string" }
     }
   }
   ```
3. **Compose** `DocNumber` — a single expression (no separate Condition action needed) using `if()`:
   ```
   if(
       equals(body('Parse_JSON')?['DocType'], 'Estimate'),
       concat('EST-', formatNumber(triggerOutputs()?['body/ID'], '0000')),
       concat('INV-', formatNumber(triggerOutputs()?['body/ID'], '0000'))
   )
   ```
   > `equals()` takes two arguments — the value to check and `'Estimate'` — then `if()`'s 2nd/3rd arguments are the true/false results. Note `DocType` is read as a plain string here (REST Choice field shape), not `.../Value` as the native connector would return it.
4. **Send an HTTP request to SharePoint** — use this instead of the **Update item** action to set `Title`. `Update item` re-validates *all* required/mandatory fields on the row (even ones you're not touching) and can throw spurious "required field missing" errors on partially-saved items; a scoped HTTP PATCH only touches the field(s) you specify.
   - **Site Address:** your site (or environment variable, if you're using one for the site URL)
   - **Method:** `PATCH`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{triggerOutputs()?['body/ID']})`
     - Use the list's **GUID** (from your environment variable) rather than `GetByTitle('InvoiceEstimates')` — GUID-based references aren't affected by the list being renamed later and match the pattern you're already using elsewhere for site/list environment variables. Swap `variables('InvoiceEstimatesListId')` for however you reference that environment variable in this flow (e.g. `outputs('Get_environment_variable...')` if pulled via the Environment Variable connector, or `variables('...')` if copied into a variable at the top of the flow).
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata",
       "Content-Type": "application/json;odata=nometadata",
       "IF-MATCH": "*"
     }
     ```
   - **Body:**
     ```json
     {
       "Title": "@{outputs('Compose')}"
     }
     ```
     (Replace `outputs('Compose')` with your actual Compose action name if renamed.)

**Notes:**
- Using the list's internal `ID` guarantees no duplicate numbers even with concurrent creates — no need for a locking/counter pattern.
- Collapsing the branch into a single Compose (`if()` expression) instead of a Condition action keeps the flow to 3 steps and is easier to read/maintain than duplicating the Compose across Condition branches.
- The HTTP PATCH approach avoids `Update item`'s behavior of re-validating the entire row against required-field rules on every update — since this flow only needs to set `Title`, a scoped PATCH is more reliable, especially before other required fields (e.g. `Customer`, `Doc Date`) may have been filled in yet by the user in Power Apps.
- If you later want numbers to reset per year (e.g. `EST-2026-0001`), swap the format expression to include `formatDateTime(utcNow(),'yyyy')`.
- No "run after" configuration is needed for the HTTP action beyond the default (only run after Compose succeeds).

---

## Flow 2: Recalculate Document Totals

**Purpose:** Whenever a line item is added, edited, or deleted, recompute the parent `InvoiceEstimates` record's `Subtotal` and `Total` (Total = Subtotal since there's no tax).

**Trigger:** `When an item is created, modified, or deleted` — SharePoint connector, list: `DocumentLines`. (Use the combined trigger, or three separate triggers pointing to the same child flow logic if your Automate environment doesn't support the combined trigger — combined is preferred to keep it to one flow.)

**Steps:**
1. **Condition** — check trigger type isn't relevant here since we need the parent ID either way. For **create/modify**, `Document` field is present on the triggering item. For **delete**, the deleted item's `Document` lookup value is still available in the trigger outputs (SharePoint retains it in the delete trigger payload) — capture it into a variable immediately:
   - **Initialize variable** `DocumentId` (Integer) = `triggerOutputs()?['body/Document/Id']` (fallback to `triggerOutputs()?['body/DocumentId']` depending on connector version — check the dynamic content picker for the exact lookup ID token).
2. **Send an HTTP request to SharePoint** — use this instead of **Get items** to retrieve the sibling line items. This avoids the `Get items` action's ODataQuery quirks with lookup-field filters (its `Filter Query` box sometimes needs the lookup's internal field name, which isn't always what the picker shows) and gives you direct control over the `$filter`/`$select` REST query.
   - **Site Address:** your site (or environment variable)
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('DocumentLinesListId')}')/items?$filter=DocumentId eq @{variables('DocumentId')}&$select=Id,Amount`
     - `DocumentId` here is the REST internal name for the lookup field's Id value (SharePoint auto-appends `Id` to a lookup column's internal name — confirm the exact internal name via `_api/web/lists(guid'...')/fields?$filter=Title eq 'Document'` if your column was created with a different internal name than `Document`).
     - Use the `DocumentLines` list's GUID environment variable, matching the pattern used in Flow 1.
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
3. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint')` (or your action's actual name), Schema (generate from a sample response, or use):
   ```json
   {
     "type": "object",
     "properties": {
       "value": {
         "type": "array",
         "items": {
           "type": "object",
           "properties": {
             "Id": { "type": "integer" },
             "Amount": { "type": ["number", "null"] }
           }
         }
       }
     }
   }
   ```
4. **Initialize variable** `RunningTotal` (Float) = `0`.
5. **Apply to each** — Items: `body('Parse_JSON')?['value']`:
   - **Set variable** `RunningTotal` = `add(variables('RunningTotal'), coalesce(items('Apply_to_each')?['Amount'], 0))`
6. **Condition** — only proceed if `DocumentId` is not null/empty (guards against edge cases where a line item is created before `Document` is set, e.g. via API).
7. **Send an HTTP request to SharePoint** (inside the Condition's Yes branch) — use this instead of **Update item** to persist the recalculated totals.
   - **Site Address:** your site (or `variables('SiteUrl')`)
   - **Method:** `PATCH`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{variables('DocumentId')})`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata",
       "Content-Type": "application/json;odata=nometadata",
       "IF-MATCH": "*"
     }
     ```
   - **Body:**
     ```json
     {
       "Subtotal": @{variables('RunningTotal')},
       "Total": @{variables('RunningTotal')}
     }
     ```

**Concurrency setting:** On the trigger, set **Concurrency Control** to a degree > 1 (e.g. 20) so rapid successive line-item edits (adding several lines quickly in the app) don't queue up and slow down the UI feedback loop. Since updates are idempotent (recompute-from-scratch each time), concurrent runs are safe.

---

## Flow 3: Convert Estimate → Invoice

**Purpose:** Button-triggered from Power Apps. Clones an Estimate's header and line items into a new Invoice-type Document, links them, and marks the Estimate as `Invoiced`.

**Trigger:** `PowerApps (V2)` trigger. Add one input parameter: `DocumentId` (Number) — the ID of the Estimate to convert.

**Steps:**
1. **Send an HTTP request to SharePoint** — use this instead of **Get item** to fetch the source Estimate's header.
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{triggerBody()['DocumentId']})?$select=Id,CustomerId,BillToSnapshot,ShipToSnapshot,PONumber,Notes,DocType`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
2. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint')`, Schema generated from a sample response (properties: `Id` integer, `CustomerId` integer, `BillToSnapshot`/`ShipToSnapshot`/`PONumber`/`Notes`/`DocType` strings).
3. **Send an HTTP request to SharePoint** — use this instead of **Create item** to create the new Invoice header.
   - **Method:** `POST`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata",
       "Content-Type": "application/json;odata=nometadata"
     }
     ```
   - **Body:**
     ```json
     {
       "DocType": "Invoice",
       "CustomerId": @{body('Parse_JSON')?['CustomerId']},
       "BillToSnapshot": "@{body('Parse_JSON')?['BillToSnapshot']}",
       "ShipToSnapshot": "@{body('Parse_JSON')?['ShipToSnapshot']}",
       "DocDate": "@{utcNow()}",
       "DueDate": "@{addDays(utcNow(), 30)}",
       "PONumber": "@{body('Parse_JSON')?['PONumber']}",
       "Status": "Draft",
       "LinkedDocumentId": @{triggerBody()['DocumentId']},
       "Notes": "@{body('Parse_JSON')?['Notes']}"
     }
     ```
     *(`Title`/Doc Number left unset — Flow 1's create-trigger will populate it automatically as `INV-####` since `DocType = Invoice` on this new item. `DueDate` defaults to 30-day terms; adjust as needed.)*
4. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint_1')` (the create response), Schema: `{ "type": "object", "properties": { "Id": { "type": "integer" } } }`. This captures the new Invoice's `Id`.
5. **Send an HTTP request to SharePoint** — use this instead of **Get items** to retrieve the Estimate's line items.
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('DocumentLinesListId')}')/items?$filter=DocumentId eq @{triggerBody()['DocumentId']}&$select=Id,ServiceId,ItemLabel,Description,Qty,Rate,Amount,SortOrder`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
6. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint_2')`, Schema: array under `value` with `Id`, `ServiceId`, `ItemLabel`, `Description`, `Qty`, `Rate`, `Amount`, `SortOrder`.
7. **Apply to each** — Items: `body('Parse_JSON_2')?['value']`:
   - **Send an HTTP request to SharePoint** — use this instead of **Create item** to clone each line onto the new Invoice.
     - **Method:** `POST`
     - **Uri:** `_api/web/lists(guid'@{variables('DocumentLinesListId')}')/items`
     - **Headers:**
       ```json
       {
         "Accept": "application/json;odata=nometadata",
         "Content-Type": "application/json;odata=nometadata"
       }
       ```
     - **Body:**
       ```json
       {
         "DocumentId": @{body('Parse_JSON_1')?['Id']},
         "ServiceId": @{items('Apply_to_each')?['ServiceId']},
         "ItemLabel": "@{items('Apply_to_each')?['ItemLabel']}",
         "Description": "@{items('Apply_to_each')?['Description']}",
         "Qty": @{items('Apply_to_each')?['Qty']},
         "Rate": @{items('Apply_to_each')?['Rate']},
         "Amount": @{items('Apply_to_each')?['Amount']},
         "SortOrder": @{items('Apply_to_each')?['SortOrder']}
       }
       ```
8. **Send an HTTP request to SharePoint** — use this instead of **Update item** to mark the original Estimate as converted.
   - **Method:** `PATCH`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{triggerBody()['DocumentId']})`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata",
       "Content-Type": "application/json;odata=nometadata",
       "IF-MATCH": "*"
     }
     ```
   - **Body:**
     ```json
     {
       "Status": "Invoiced",
       "LinkedDocumentId": @{body('Parse_JSON_1')?['Id']}
     }
     ```
     — cross-links both directions (original Estimate → new Invoice, and new Invoice → original Estimate, set in step 3).
9. **Respond to PowerApps** — return `NewInvoiceId` (Number) = `body('Parse_JSON_1')?['Id']` so the app can navigate straight to the new invoice.

**Notes:**
- Flow 2 (totals recalculation) will fire automatically as each new `DocumentLines` item is created in step 7 — no need to manually copy Subtotal/Total; they'll self-populate. If you want to avoid N redundant recalculation runs during the copy, you can instead directly copy `Subtotal`/`Total` from the Estimate into step 3's body as a shortcut, accepting Flow 2 will simply confirm the same value once the last line is created.

---

## Flow 4 (Optional): Generate Document PDF

**Purpose:** Produce a PDF matching the paper layout and make it available for the user to view/download from Power Apps. This flow does **not** email or change `Status` — printing and sending to the customer is a manual step the user does themselves after reviewing the PDF.

**Trigger:** `PowerApps (V2)` trigger, input: `DocumentId` (Number).

**Steps:**
1. **Send an HTTP request to SharePoint** — use this instead of **Get item** to fetch the header.
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('InvoiceEstimatesListId')}')/items(@{triggerBody()['DocumentId']})?$select=Id,Title,DocType,CustomerId,DocDate,DueDate,PONumber,Notes,Total`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
2. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint')`, Schema generated from a sample response.
3. **Send an HTTP request to SharePoint** — use this instead of **Get items** to fetch the line items, sorted server-side.
   - **Method:** `GET`
   - **Uri:** `_api/web/lists(guid'@{variables('DocumentLinesListId')}')/items?$filter=DocumentId eq @{triggerBody()['DocumentId']}&$orderby=SortOrder asc&$select=Id,ItemLabel,Description,Qty,Rate,Amount`
   - **Headers:**
     ```json
     {
       "Accept": "application/json;odata=nometadata"
     }
     ```
4. **Parse JSON** — Content: `body('Send_an_HTTP_request_to_SharePoint_1')`, Schema: array under `value`.
5. **Populate a Word template** (Word Online / Power Automate action) using a `.docx` template stored in a SharePoint document library, with content controls for Customer, Dates, PO#, and a repeating table row bound to `body('Parse_JSON_1')?['value']`.
   - Build the template first: create a Word doc styled like the sample Estimate/Invoice, insert content controls (Developer tab → Insert Controls) named to match the flow's field mapping, and save it to a `Templates` library in the same site.
6. **Convert Word document to PDF** (built-in Power Automate action, no separate connector needed).
7. **Create file** (SharePoint) — save the PDF to a document library so it has a stable, shareable URL. (This one native action is kept as-is rather than converted to raw HTTP — file/library uploads require multipart binary handling that the native connector manages for you far more simply than a raw HTTP PUT/POST would.)
   - **Site Address:** your site
   - **Folder Path:** `/GeneratedDocuments` (create this document library once, ahead of time)
   - **File Name:** `@{body('Parse_JSON')?['Title']}.pdf` (e.g. `EST-0148.pdf` / `INV-2835.pdf`)
   - **File Content:** output of step 6 (Convert Word document to PDF)
   - Set **If file already exists** behavior to overwrite, so re-generating a PDF for the same document replaces the old copy rather than erroring or duplicating.
8. **Respond to PowerApps** — return:
   - `PdfUrl` (Text) = `concat(variables('SiteUrl'), '/GeneratedDocuments/', body('Parse_JSON')?['Title'], '.pdf')` — a direct link the app can open in a browser tab or PDF viewer control for viewing/downloading.
   - `DocTitle` (Text) = `body('Parse_JSON')?['Title']` — handy for labeling the link/button in the app.

**Notes:**
- No email step and no `Status` update — the app simply shows/opens the generated PDF; the user decides when and how to send or print it.
- Because step 7 overwrites on re-run, clicking "Generate PDF" again after editing line items produces an up-to-date file at the same URL — no cleanup needed.
- In Power Apps, wire the button that calls this flow to open the returned `PdfUrl` (e.g. `Launch(pdfUrl)` to open it in a new browser tab, where the user can view, print, or save/download it using their browser's built-in PDF controls) — see the Power Apps guide for the exact formula.

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
- `DOC - Generate Document PDF`
