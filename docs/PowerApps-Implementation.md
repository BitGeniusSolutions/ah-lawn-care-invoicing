# Power Apps Implementation Guide — A&H Lawn Care Invoicing

Build order and concrete formulas for the Canvas app. Connect all four SharePoint lists (`Customers`, `Services`, `InvoiceEstimates`, `DocumentLines`) as data sources, and add the `DOC - Get Invoice Estimates` Power Automate flow as a data source (see the Power Automate guide, Flow 5), before starting.

---

## 0. App Setup
1. Power Apps Studio → Create → Blank canvas app (Tablet format recommended — more room for the line-items table).
2. **Data** → Add data → SharePoint → connect to your site → add all four lists.
3. **App.OnStart** (App object formula bar):
   ```
   Set(varCurrentDoc, Blank());
   Set(varIsNewDoc, false);
   Set(varFilterDocType, Blank());
   Set(varFilterStatus, Blank());
   Set(varFilterSince, Blank());
   Set(varFilterDateTo, Blank());
   Set(varSearchText, "");
   Set(varDocsRaw, 'DOC - Get Invoice Estimates'.Run());
   Set(varDocsResult, ParseJSON(varDocsRaw.response));
   If(
       Boolean(varDocsResult.success),
       ClearCollect(
           colDocuments,
           ForAll(
               varDocsResult.data,
               {
                   ID: Value(ThisRecord.Id),
                   DocNumber: Text(ThisRecord.Title),
                   DocType: Text(ThisRecord.DocType),
                   CustomerId: Value(ThisRecord.CustomerId),
                   CustomerName: Text(ThisRecord.Customer.Title),
                   DocDate: DateValue(Text(ThisRecord.DocDate)),
                   DueDate: If(IsBlank(ThisRecord.DueDate), Blank(), DateValue(Text(ThisRecord.DueDate))),
                   PONumber: Text(ThisRecord.PONumber),
                   Status: Text(ThisRecord.Status),
                   Subtotal: Value(ThisRecord.Subtotal),
                   Total: Value(ThisRecord.Total),
                   Notes: Text(ThisRecord.Notes)
               }
           )
       ),
       Notify("Couldn't load documents. Please try again.", NotificationType.Error)
   )
   ```
   (Add `DOC - Get Invoice Estimates` as a data source first: **Power Automate** pane → Add flow.) See section 0.5 below for why this collection exists and where else to re-run this block.
4. Set a simple theme: App → Settings → color, or apply a Fluent theme for a clean, modern look. Not critical to function.

---

## 0.5 Data Loading Pattern (Avoiding Delegation)

`InvoiceEstimates` is a Choice-heavy list, and the list/dashboard screens need OR'd conditions (`DocType = X Or IsBlank(...)`) plus cross-record aggregates (`CountRows`, date-range filters). Filtering/counting a SharePoint data source directly with that shape of formula reliably trips Power Apps' delegation warning — past roughly 500-2000 rows (depending on tenant setting) the results silently become incomplete.

The fix used throughout this guide: **Flow 5 (`DOC - Get Invoice Estimates`)** fetches every row once (with the `Customer` display name already joined in via `$expand`), and the app caches it in a plain collection, `colDocuments`, via the block in **App.OnStart** above. From then on:
- `HomeScreen`'s dashboard cards (§1.2) and `DocumentListScreen`'s gallery/search (§4) all `Filter()`/`CountRows()` against `colDocuments`, not the `InvoiceEstimates` data source directly. **Collections are never subject to delegation limits** — Power Apps holds the whole thing in memory and evaluates formulas locally, so any combination of `Or`, `CountRows`, date-range, and text-search conditions works correctly regardless of row count (up to Power Apps' general collection-size practicality, which is far beyond what this business will ever produce).
- Because the collection's `Status`/`DocType` columns are plain `Text` (flattened out of the Choice fields by the `ForAll` in App.OnStart), formulas against `colDocuments` also drop the `.Value` suffix needed everywhere else against the live `InvoiceEstimates` source — e.g. `Status = "Draft"` instead of `Status.Value = "Draft"`.
- **This is a snapshot, not live data.** Re-run the App.OnStart block (wrap it in a reusable pattern, or just duplicate the same `Set`/`ClearCollect` calls) after anything that changes `InvoiceEstimates` server-side, so `colDocuments` doesn't go stale:
  - After `frmDocument.OnSuccess` (header saved) — see §5.1.
  - After a successful "Convert to Invoice" call — see §5.3.
  - On a manual **Refresh** button on `DocumentListScreen` (§4), for peace of mind after any external change (e.g. someone editing the list directly in SharePoint).
- The header **Form control** on `DocumentDetailScreen` (§5.1) and the **line items gallery** (§5.2) still read/write `InvoiceEstimates`/`DocumentLines` directly — those work against a single record by `ID`, which is always delegable, so there's no reason to route them through the cache. Only the list/search/dashboard views (many-row aggregate operations) need this pattern.

---

## 1. HomeScreen

### 1.1 HTML header banner
Add an **HTML text** control (`htmlHeader`) pinned at the top of the screen (Height: ~90-110px, Width: `Parent.Width`). The HTML text control renders a safe subset of HTML/CSS (no `<script>`, no external JS) — good enough for a styled banner without needing an image asset.

- `HtmlText`:
  ```
  "<div style='display:flex;align-items:center;justify-content:space-between;height:100%;width:100%;margin:0;padding:0 24px;box-sizing:border-box;background:linear-gradient(90deg,#1e4d2b 0%,#2e7d32 60%,#4caf50 100%);font-family:Segoe UI,Arial,sans-serif;color:#ffffff;'>" &
  "<div>" &
    "<div style='font-size:22px;font-weight:600;letter-spacing:0.3px;'>A&amp;H Lawn Care Services</div>" &
    "<div style='font-size:13px;opacity:0.85;margin-top:2px;'>Estimates &amp; Invoicing</div>" &
  "</div>" &
  "<div style='font-size:13px;opacity:0.9;text-align:right;'>" & Text(Today(), "dddd, mmmm d, yyyy") & "</div>" &
  "</div>"
  ```
- Since `HtmlText` is a string property, build it with `&` concatenation (as above) rather than a raw multi-line literal, so `Text(Today(), ...)` can be embedded dynamically — Power Fx doesn't support `@{}` interpolation inside HTML text like Power Automate does. If you prefer a fully static banner, replace the date `Text(...)` piece with plain text and drop the concatenation.
- Optional: swap the gradient colors for whatever green/brand shade you like — this is plain inline CSS, easy to retheme later without touching app logic.

### 1.2 Dashboard cards
Four equal-width card containers below the header, laid out in a horizontal container (or 4 separate containers with fixed `X`/`Width` if not using a layout container). Each card is a rectangle/container with a big number label, a caption label, and an `OnSelect` that drills into `DocumentListScreen` with a preset filter.

All four cards filter/count against **`colDocuments`** (see §0.5) rather than the `InvoiceEstimates` data source directly — this is what keeps these formulas delegation-safe despite the `Or`/inequality/date-range conditions.

Two context variables (initialized in App.OnStart, §0) support drill-down from these cards into `DocumentListScreen`:
- `varFilterStatus` (Text) — blank = no status filter.
- `varFilterSince` (Date) — blank = no date filter; when set, `DocumentListScreen` shows only items with `DocDate >= varFilterSince`.

**Card 1 — Drafts**
- Count label `Text`: `CountRows(Filter(colDocuments, Status = "Draft"))`
- Caption label `Text`: `"Drafts"`
- `OnSelect`:
  ```
  Set(varFilterDocType, Blank());
  Set(varFilterStatus, "Draft");
  Set(varFilterSince, Blank());
  Navigate(DocumentListScreen, ScreenTransition.Cover)
  ```

**Card 2 — Open Invoices**
- Count label `Text`: `CountRows(Filter(colDocuments, DocType = "Invoice", Status <> "Paid", Status <> "Void"))`
- Caption label `Text`: `"Open Invoices"`
- `OnSelect`:
  ```
  Set(varFilterDocType, "Invoice");
  Set(varFilterStatus, Blank());
  Set(varFilterSince, Blank());
  Navigate(DocumentListScreen, ScreenTransition.Cover)
  ```
  > "Open" here means not `Paid`/`Void`; the list screen itself doesn't need a separate status filter for this case since it's really a Doc Type + exclusion rule best expressed on the list screen's own default view. If you want the card to strictly hand off its own definition of "open," instead pass both `varFilterDocType = "Invoice"` and skip `varFilterStatus`, and add an `Open Invoices` toggle option alongside the Estimate/Invoice/All tabs on `DocumentListScreen`.

**Card 3 — Recent Estimates (last 30 days)**
- Count label `Text`: `CountRows(Filter(colDocuments, DocType = "Estimate", DocDate >= Today() - 30))`
- Caption label `Text`: `"Recent Estimates (30 days)"`
- `OnSelect`:
  ```
  Set(varFilterDocType, "Estimate");
  Set(varFilterStatus, Blank());
  Set(varFilterSince, Today() - 30);
  Navigate(DocumentListScreen, ScreenTransition.Cover)
  ```

**Card 4 — Recent Invoices (last 30 days)**
- Count label `Text`: `CountRows(Filter(colDocuments, DocType = "Invoice", DocDate >= Today() - 30))`
- Caption label `Text`: `"Recent Invoices (30 days)"`
- `OnSelect`:
  ```
  Set(varFilterDocType, "Invoice");
  Set(varFilterStatus, Blank());
  Set(varFilterSince, Today() - 30);
  Navigate(DocumentListScreen, ScreenTransition.Cover)
  ```

Styling tip: give each card container a `Fill` of white, `BorderRadius` ~8, a subtle `DropShadow`, the count label large/bold (~36px) in the brand green (`#2e7d32`), and the caption label smaller/gray beneath it. Set each card's `Hover.Fill` slightly darker and `OnSelect` as above so the whole card is clickable, not just a button inside it.

### 1.3 Buttons
- "New Estimate" → `Set(varIsNewDoc, true); Navigate(DocumentDetailScreen, ScreenTransition.Cover, {NewDocType: "Estimate"})`
- "New Invoice" → `Set(varIsNewDoc, true); Navigate(DocumentDetailScreen, ScreenTransition.Cover, {NewDocType: "Invoice"})`
- "Customers" → `Navigate(CustomerListScreen)`
- "Services" → `Navigate(ServiceListScreen)`
- "All Documents" → `Set(varFilterDocType, Blank()); Set(varFilterStatus, Blank()); Set(varFilterSince, Blank()); Navigate(DocumentListScreen)`

---

## 2. Customer Screens (pure Form control)

### CustomerListScreen
- **Gallery** (vertical, blank layout) — `Items: SortByColumns(Filter(Customers, Active.Value = true), "Title", SortOrder.Ascending)`
- Show: Customer Name, Property Name, City, Phone.
- Gallery `OnSelect`: `Navigate(CustomerFormScreen, ScreenTransition.Cover, {FormMode: FormMode.View, SelectedCustomer: ThisItem})`
- "New Customer" button: `Navigate(CustomerFormScreen, ScreenTransition.Cover, {FormMode: FormMode.New})`

### CustomerFormScreen
- **Form control** (Edit form), named `frmCustomer`:
  - `DataSource`: `Customers`
  - `Item`: `If(CustomerFormScreen.FormMode = FormMode.New, Defaults(Customers), SelectedCustomer)`
  - `DefaultMode`: passed-in `FormMode` param (`New` or `View`/`Edit`)
  - Fields to include (drag from field pane): Customer Name, Property Name, Billing Address, City, State, Zip, Contact Name, Phone, Email, Property Number, Notes, Active.
- Toolbar buttons:
  - **Edit** (visible only in View mode): `EditForm(frmCustomer)`
  - **Save**: `SubmitForm(frmCustomer)`
  - **Cancel**: `ResetForm(frmCustomer); Navigate(CustomerListScreen, ScreenTransition.Cover)`
  - `frmCustomer.OnSuccess`: `Navigate(CustomerListScreen, ScreenTransition.Cover)`

This is the "pure form control" pattern — repeat identically for Services.

---

## 3. Service Screens (identical pattern to Customers)

### ServiceListScreen
- Gallery `Items`: `SortByColumns(Filter(Services, Active.Value = true), "Category", "Title", SortOrder.Ascending)` — group visually by Category if desired using a grouped gallery layout.

### ServiceFormScreen
- Form control `frmService`, DataSource `Services`, fields: Service Name, Category, Unit, Default Rate, Description, Active.
- Same Edit/Save/Cancel button pattern as Customers.

---

## 4. DocumentListScreen

- **Gallery**, `Items`:
  ```
  SortByColumns(
    Filter(
      colDocuments,
      (DocType = varFilterDocType Or IsBlank(varFilterDocType)),
      (Status = varFilterStatus Or IsBlank(varFilterStatus)),
      (DocDate >= varFilterSince Or IsBlank(varFilterSince)),
      (IsBlank(varFilterDateTo) Or DocDate <= varFilterDateTo),
      (
        IsBlank(Trim(varSearchText))
        Or varSearchText in DocNumber
        Or varSearchText in CustomerName
        Or varSearchText in PONumber
        Or varSearchText in Notes
      )
    ),
    "DocDate", SortOrder.Descending
  )
  ```
  This filters the local `colDocuments` collection (§0.5), not `InvoiceEstimates` directly — fully delegation-safe regardless of how many conditions are combined, since collection filtering is always evaluated client-side. `varFilterStatus`/`varFilterSince` let the HomeScreen dashboard cards (§1.2) drill into a pre-filtered list; when navigating here any other way, reset all filter/search variables first (e.g. the "All Documents" button in §1.3) so stale filters don't linger.
- Add a toggle/tab control at top to set `varFilterDocType` to `"Estimate"`, `"Invoice"`, or blank (All) — also clear `varFilterStatus`/`varFilterSince`/`varFilterDateTo`/`varSearchText` when the user manually changes this toggle, since it represents starting a fresh filter.
- Show per row: `DocNumber`, `CustomerName`, `DocDate`, `Status`, `Total` (all plain fields off `colDocuments` — no `.Value` needed, see §0.5).
- `OnSelect`:
  ```
  Set(varCurrentDoc, LookUp(InvoiceEstimates, ID = ThisItem.ID));
  Set(varIsNewDoc, false);
  Navigate(DocumentDetailScreen, ScreenTransition.Cover)
  ```
  Note the `LookUp` against the real `InvoiceEstimates` data source (not `ThisItem` from the gallery directly) — `ThisItem` here is a flattened row from `colDocuments` with plain-text `Status`/`DocType`, but `DocumentDetailScreen`'s Form control (§5.1) needs the actual SharePoint record (with Choice-typed fields) to bind and edit correctly. This single-record `LookUp` by `ID` is always delegable.
- "New Estimate"/"New Invoice" buttons same as HomeScreen.

### 4.1 Search box
- Text input `txtSearch`, placeholder "Search doc #, customer, PO, notes…". `OnChange`: `Set(varSearchText, txtSearch.Text)`.
- Search is a simple substring match (`in` operator, case-insensitive) across `DocNumber`, `CustomerName`, `PONumber`, and `Notes` — see the gallery formula above. Since it runs against the cached collection, there's no OData syntax to fight and no delegation limit on how many fields you `Or` together; add more fields to the search (e.g. `Description` on line items, if you later flatten that in too) the same way.

### 4.2 Advanced filters (optional panel)
Add a collapsible container (toggle button "Advanced ▾") revealing extra filter controls, all wired to context variables already read by the gallery's `Filter()`:
- **Date To** picker `dtpDateTo` → `OnChange`: `Set(varFilterDateTo, dtpDateTo.SelectedDate)`. (`varFilterSince`/`varFilterDateTo` together give a full date-range filter.)
- **Status** dropdown `ddStatus` (`Items: ["", "Draft", "Sent", "Approved", "Invoiced", "Paid", "Void"]`) → `OnChange`: `Set(varFilterStatus, If(ddStatus.Selected.Value = "", Blank(), ddStatus.Selected.Value))`.
- "Clear Filters" button → resets everything and clears the search box:
  ```
  Set(varFilterDocType, Blank());
  Set(varFilterStatus, Blank());
  Set(varFilterSince, Blank());
  Set(varFilterDateTo, Blank());
  Set(varSearchText, "");
  Reset(txtSearch)
  ```

### 4.3 Refresh button
`colDocuments` is a snapshot (§0.5) — add a small refresh icon button near the search box that re-runs the same load block from App.OnStart:
```
Set(varDocsRaw, 'DOC - Get Invoice Estimates'.Run());
Set(varDocsResult, ParseJSON(varDocsRaw.response));
If(
    Boolean(varDocsResult.success),
    ClearCollect(
        colDocuments,
        ForAll(
            varDocsResult.data,
            {
                ID: Value(ThisRecord.Id),
                DocNumber: Text(ThisRecord.Title),
                DocType: Text(ThisRecord.DocType),
                CustomerId: Value(ThisRecord.CustomerId),
                CustomerName: Text(ThisRecord.Customer.Title),
                DocDate: DateValue(Text(ThisRecord.DocDate)),
                DueDate: If(IsBlank(ThisRecord.DueDate), Blank(), DateValue(Text(ThisRecord.DueDate))),
                PONumber: Text(ThisRecord.PONumber),
                Status: Text(ThisRecord.Status),
                Subtotal: Value(ThisRecord.Subtotal),
                Total: Value(ThisRecord.Total),
                Notes: Text(ThisRecord.Notes)
            }
        )
    ),
    Notify("Couldn't refresh documents. Please try again.", NotificationType.Error)
)
```

---

## 5. DocumentDetailScreen (the hybrid screen)

### 5.1 Header — Form control
- **Form control** `frmDocument`:
  - `DataSource`: `InvoiceEstimates`
  - `Item`:
    ```
    If(
      varIsNewDoc,
      Patch(Defaults(InvoiceEstimates), {'Doc Type': {Value: NewDocType}, Status: {Value: "Draft"}, DocDate: Today()}),
      varCurrentDoc
    )
    ```
  - `DefaultMode`: `If(varIsNewDoc, FormMode.New, FormMode.Edit)`
  - Fields: Doc Type (lock/disable editing after first save — see note below), Customer, Doc Date, Due Date, PO Number, Status, Notes.
  - **Note:** Once a Document is saved, prevent changing `Doc Type` (it drives numbering and downstream flows) — set that card's `DisplayMode` to `If(varIsNewDoc, DisplayMode.Edit, DisplayMode.View)`.
- Toolbar: **Save Header** → `SubmitForm(frmDocument)`; `frmDocument.OnSuccess`: `Set(varCurrentDoc, frmDocument.LastSubmit); Set(varIsNewDoc, false); Set(varDocsRaw, 'DOC - Get Invoice Estimates'.Run()); Set(varDocsResult, ParseJSON(varDocsRaw.response)); If(Boolean(varDocsResult.success), ClearCollect(colDocuments, ForAll(varDocsResult.data, {ID: Value(ThisRecord.Id), DocNumber: Text(ThisRecord.Title), DocType: Text(ThisRecord.DocType), CustomerId: Value(ThisRecord.CustomerId), CustomerName: Text(ThisRecord.Customer.Title), DocDate: DateValue(Text(ThisRecord.DocDate)), DueDate: If(IsBlank(ThisRecord.DueDate), Blank(), DateValue(Text(ThisRecord.DueDate))), PONumber: Text(ThisRecord.PONumber), Status: Text(ThisRecord.Status), Subtotal: Value(ThisRecord.Subtotal), Total: Value(ThisRecord.Total), Notes: Text(ThisRecord.Notes)})))` — the trailing re-load keeps `colDocuments` (§0.5) in sync so the just-saved/updated header shows correctly if the user navigates back to `DocumentListScreen`/`HomeScreen` (so the line-items gallery below can now filter by the real Document ID once it exists).

### 5.2 Line Items — Editable Gallery
> Only enabled once the header has been saved at least once (`varCurrentDoc` has a real ID) — line items need a parent to attach to.

- **Gallery** `galLines`, `Items`:
  ```
  SortByColumns(
    Filter(DocumentLines, Document.Id = varCurrentDoc.ID),
    "SortOrder", SortOrder.Ascending
  )
  ```
- Per-row controls (all bound to `ThisItem`, patched on change — no nested form control needed):
  - **Service combo box** `cboService`: `Items: Filter(Services, Active.Value = true)`. `OnChange`:
    ```
    Patch(DocumentLines, ThisItem, {
      Service: {Id: cboService.Selected.ID, Value: cboService.Selected.Title},
      Rate: cboService.Selected.'Default Rate',
      'Item Label': cboService.Selected.Title
    })
    ```
  - **Item Label** text input `txtItemLabel`: `Default: ThisItem.'Item Label'`. `OnChange`: `Patch(DocumentLines, ThisItem, {'Item Label': txtItemLabel.Text})`
  - **Description** text input `txtDescription`: same pattern, field `Description`.
  - **Qty** text input `txtQty` (numeric): `OnChange`:
    ```
    Patch(DocumentLines, ThisItem, {
      Qty: Value(txtQty.Text),
      Amount: Value(txtQty.Text) * ThisItem.Rate
    })
    ```
  - **Rate** text input `txtRate` (numeric): `OnChange`:
    ```
    Patch(DocumentLines, ThisItem, {
      Rate: Value(txtRate.Text),
      Amount: ThisItem.Qty * Value(txtRate.Text)
    })
    ```
  - **Amount** label (read-only): `Text: Text(ThisItem.Amount, "[$-en-US]$#,##0.00")`
  - **Delete icon**: `Remove(DocumentLines, ThisItem)`
- **"+ Add Line" button** (below gallery):
  ```
  Patch(
    DocumentLines,
    Defaults(DocumentLines),
    {
      Document: {Id: varCurrentDoc.ID, Value: varCurrentDoc.Title},
      SortOrder: CountRows(Filter(DocumentLines, Document.Id = varCurrentDoc.ID)) + 1,
      Qty: 1,
      Rate: 0,
      Amount: 0
    }
  )
  ```
- **Footer total label** (live, before flow catches up): `Text: "Total: " & Text(Sum(Filter(DocumentLines, Document.Id = varCurrentDoc.ID), Amount), "[$-en-US]$#,##0.00")`
  - This is a live client-side sum shown immediately; Power Automate Flow 2 (Recalculate Totals) will separately persist `Subtotal`/`Total` on the `InvoiceEstimates` record moments later so other views (galleries, PDF) stay in sync without the app needing to write those fields itself.

### 5.3 Convert to Invoice button (Estimates only)
- Visible: `varCurrentDoc.'Doc Type'.Value = "Estimate" And varCurrentDoc.Status.Value <> "Invoiced"`
- `OnSelect`:
  ```
  Set(varConvertRaw, 'DOC - Convert Estimate to Invoice'.Run(varCurrentDoc.ID));
  Set(varConvertResult, ParseJSON(varConvertRaw.response));
  If(
      Boolean(varConvertResult.success),
      Set(varCurrentDoc, LookUp(InvoiceEstimates, ID = Value(varConvertResult.data.NewInvoiceId)));
      Set(varDocsRaw2, 'DOC - Get Invoice Estimates'.Run());
      Set(varDocsResult, ParseJSON(varDocsRaw2.response));
      If(
          Boolean(varDocsResult.success),
          ClearCollect(
              colDocuments,
              ForAll(
                  varDocsResult.data,
                  {
                      ID: Value(ThisRecord.Id),
                      DocNumber: Text(ThisRecord.Title),
                      DocType: Text(ThisRecord.DocType),
                      CustomerId: Value(ThisRecord.CustomerId),
                      CustomerName: Text(ThisRecord.Customer.Title),
                      DocDate: DateValue(Text(ThisRecord.DocDate)),
                      DueDate: If(IsBlank(ThisRecord.DueDate), Blank(), DateValue(Text(ThisRecord.DueDate))),
                      PONumber: Text(ThisRecord.PONumber),
                      Status: Text(ThisRecord.Status),
                      Subtotal: Value(ThisRecord.Subtotal),
                      Total: Value(ThisRecord.Total),
                      Notes: Text(ThisRecord.Notes)
                  }
              )
          )
      );
      Navigate(DocumentDetailScreen, ScreenTransition.Cover),
      Notify("Couldn't convert this Estimate to an Invoice. Please try again.", NotificationType.Error)
  )
  ```
  (Add the Power Automate flow as a data source first: **Power Automate** pane → Add flow → select `DOC - Convert Estimate to Invoice`.)
  - The flow returns a single `response` text output shaped `{ success, data }` (see Power Automate guide's "Standard response envelope" convention, and Flow 3's Try/Catch scopes). `ParseJSON` turns it into an untyped object; wrap numeric fields in `Value()` and text fields in `Text()` when reading them out of `.data`.
  - Note the second reload of `colDocuments` (§0.5, using a separate `varDocsRaw2` variable name to avoid clobbering the in-flight `varDocsRaw`/`varDocsResult` from App.OnStart if it somehow overlaps) — this converts a new Invoice header + cloned lines server-side, so the cache must be refreshed for the new Invoice and the now-`Invoiced` Estimate to show correctly back on `DocumentListScreen`/`HomeScreen`.

### 5.4 Generate PDF button (optional, once Flow 4 exists)
- `OnSelect`:
  ```
  Set(varPdfRaw, 'DOC - Generate Document PDF'.Run(varCurrentDoc.ID));
  Set(varPdfResult, ParseJSON(varPdfRaw.response));
  If(
      Boolean(varPdfResult.success),
      Launch(Text(varPdfResult.data.pdfURL)),
      Notify("Couldn't generate the PDF. Please try again.", NotificationType.Error)
  )
  ```
  `Launch()` opens the PDF's OneDrive sharing link in a new browser tab/window, where the user can view it and use their browser's built-in controls to print, download, or save it — the app itself does no emailing or status changes; the flow doesn't touch `Status`. As with 5.3, the flow returns `{ success, data }` as a JSON string in `response` — check `.success` before trusting `.data.pdfURL`/`.data.docTitle`.

---

## 6. Navigation Summary
```
HomeScreen ──▶ CustomerListScreen ──▶ CustomerFormScreen
           ──▶ ServiceListScreen  ──▶ ServiceFormScreen
           ──▶ DocumentListScreen ──▶ DocumentDetailScreen
           ──▶ DocumentDetailScreen (New Estimate/Invoice, direct)
```

---

## 7. Testing Checklist
- [ ] Create a new Customer, Service via forms — confirm data lands correctly in SharePoint
- [ ] Create a new Estimate — confirm header saves, Doc Number auto-populates (Flow 1) shortly after
- [ ] Add 3–4 line items — confirm Amount calculates correctly per row, footer total matches sum
- [ ] Edit a line item's Qty/Rate — confirm Amount and footer total update live
- [ ] Delete a line item — confirm gallery updates and footer total recalculates
- [ ] Wait a few seconds, refresh InvoiceEstimates list — confirm `Subtotal`/`Total` on the header match (Flow 2 working)
- [ ] Click "Convert to Invoice" on a test Estimate — confirm new Invoice created with cloned lines, correct Doc Number prefix, `Linked Document` set both directions, original Estimate Status = `Invoiced`
- [ ] (If built) Click "Generate PDF" — confirm a correctly formatted PDF opens in a new browser tab from the SharePoint library link, with no email sent and `Status` unchanged
- [ ] Confirm `App.OnStart` populates `colDocuments` (check row count roughly matches the `InvoiceEstimates` list) and that no delegation warning icon appears on the `HomeScreen` cards or `DocumentListScreen` gallery formulas
- [ ] Confirm the four dashboard card counts match manual counts in SharePoint (Drafts, Open Invoices, Recent Estimates/Invoices last 30 days)
- [ ] Type in the `DocumentListScreen` search box — confirm results narrow correctly across Doc Number, Customer, PO Number, and Notes
- [ ] Set an Advanced date range and/or Status filter — confirm they combine correctly with the search box (all conditions are AND'd together)
- [ ] Click the Refresh button on `DocumentListScreen` after editing a record directly in SharePoint — confirm the change appears without restarting the app
- [ ] Save a Document header, then navigate back to `DocumentListScreen`/`HomeScreen` — confirm the saved/updated record and counts reflect the change immediately (via the reload in `frmDocument.OnSuccess`)

---

## 8. Notes on Form Control Reuse
Instead of separate `frmCustomer`/`frmService`/`frmDocument` forms with fully manual field layouts, you can speed up initial build using **Power Apps' "Generate app from SharePoint list"** wizard once per list to scaffold a starting Form control layout, then delete the auto-generated screens you don't need and wire the resulting form into the screens/navigation described above. This is optional — building forms directly from the field pane (as described above) gives you more control over field order and grouping to match the paper document layout.
