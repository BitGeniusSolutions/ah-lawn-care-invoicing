# Power Apps Implementation Guide — A&H Lawn Care Invoicing

Build order and concrete formulas for the Canvas app. Connect all four SharePoint lists (`Customers`, `Services`, `InvoiceEstimates`, `DocumentLines`) as data sources before starting.

---

## 0. App Setup
1. Power Apps Studio → Create → Blank canvas app (Tablet format recommended — more room for the line-items table).
2. **Data** → Add data → SharePoint → connect to your site → add all four lists.
3. **App.OnStart** (App object formula bar):
   ```
   Set(varCurrentDoc, Blank());
   Set(varIsNewDoc, false);
   ```
4. Set a simple theme: App → Settings → color, or apply a Fluent theme for a clean, modern look. Not critical to function.

---

## 1. HomeScreen

Controls:
- **Label** — Title "A&H Lawn Care Services — Estimates & Invoices"
- **Gallery/Labels** — quick counts:
  - Draft Estimates: `CountRows(Filter(InvoiceEstimates, 'Doc Type'.Value = "Estimate", Status.Value = "Draft"))`
  - Open Invoices: `CountRows(Filter(InvoiceEstimates, 'Doc Type'.Value = "Invoice", Status.Value <> "Paid", Status.Value <> "Void"))`
- **Buttons:**
  - "New Estimate" → `Set(varIsNewDoc, true); Navigate(DocumentDetailScreen, ScreenTransition.Cover, {NewDocType: "Estimate"})`
  - "New Invoice" → `Set(varIsNewDoc, true); Navigate(DocumentDetailScreen, ScreenTransition.Cover, {NewDocType: "Invoice"})`
  - "Customers" → `Navigate(CustomerListScreen)`
  - "Services" → `Navigate(ServiceListScreen)`
  - "All Documents" → `Navigate(DocumentListScreen)`

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
    Filter(InvoiceEstimates, 'Doc Type'.Value = varFilterDocType Or IsBlank(varFilterDocType)),
    "DocDate", SortOrder.Descending
  )
  ```
- Add a toggle/tab control at top to set `varFilterDocType` to `"Estimate"`, `"Invoice"`, or blank (All).
- Show per row: Doc Number (Title), Customer.Value, Doc Date, Status, Total.
- `OnSelect`: `Set(varCurrentDoc, ThisItem); Set(varIsNewDoc, false); Navigate(DocumentDetailScreen, ScreenTransition.Cover)`
- "New Estimate"/"New Invoice" buttons same as HomeScreen.

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
- Toolbar: **Save Header** → `SubmitForm(frmDocument)`; `frmDocument.OnSuccess`: `Set(varCurrentDoc, frmDocument.LastSubmit); Set(varIsNewDoc, false)` (so the line-items gallery below can now filter by the real Document ID once it exists).

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
  Set(varConvertResult, 'DOC - Convert Estimate to Invoice'.Run(varCurrentDoc.ID));
  Set(varCurrentDoc, LookUp(InvoiceEstimates, ID = varConvertResult.NewInvoiceId));
  Navigate(DocumentDetailScreen, ScreenTransition.Cover)
  ```
  (Add the Power Automate flow as a data source first: **Power Automate** pane → Add flow → select `DOC - Convert Estimate to Invoice`.)

### 5.4 Generate PDF button (optional, once Flow 4 exists)
- `OnSelect`:
  ```
  Set(varPdfResult, 'DOC - Generate Document PDF'.Run(varCurrentDoc.ID));
  Launch(varPdfResult.PdfUrl)
  ```
  `Launch()` opens the PDF's SharePoint library URL in a new browser tab/window, where the user can view it and use their browser's built-in controls to print, download, or save it — the app itself does no emailing or status changes; the flow doesn't touch `Status`.

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

---

## 8. Notes on Form Control Reuse
Instead of separate `frmCustomer`/`frmService`/`frmDocument` forms with fully manual field layouts, you can speed up initial build using **Power Apps' "Generate app from SharePoint list"** wizard once per list to scaffold a starting Form control layout, then delete the auto-generated screens you don't need and wire the resulting form into the screens/navigation described above. This is optional — building forms directly from the field pane (as described above) gives you more control over field order and grouping to match the paper document layout.
