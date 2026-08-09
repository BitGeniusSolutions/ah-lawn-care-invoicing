# A&H Lawn Care Services — Invoicing & Estimates App

Documentation and design artifacts for a simple Power Platform (Power Apps + SharePoint) application to manage estimates, invoices, customers, and services for **A&H Lawn Care Services**.

## Contents

- [`docs/Invoicing-App-Design.md`](docs/Invoicing-App-Design.md) — Overall design document: SharePoint list schema (Customers, Services, InvoiceEstimates, DocumentLines), Power Apps screen structure, Power Automate flow plan, and confirmed scope decisions.
- [`docs/SharePoint-Implementation.md`](docs/SharePoint-Implementation.md) — Step-by-step build guide for the 4 SharePoint lists (columns, types, choice values, lookups, indexing, views, permissions).
- [`docs/PowerAutomate-Implementation.md`](docs/PowerAutomate-Implementation.md) — Build guide for the 4 flows: document numbering, totals recalculation, Estimate→Invoice conversion, and optional PDF/email.
- [`docs/PowerApps-Implementation.md`](docs/PowerApps-Implementation.md) — Screen-by-screen Canvas app build guide with concrete formulas for forms, the editable line-items gallery, and navigation.

## Project Summary

- **Stack:** Power Apps (Canvas app) + SharePoint Lists + Power Automate
- **Scope:** Single-user internal tool for creating/managing estimates and invoices, tracking customers, and maintaining a reusable service/price list
- **Design approach:** SharePoint Form controls for simple CRUD (Customers, Services, Document headers) with a lightweight editable gallery for invoice/estimate line items
- **Tax handling:** Not tracked — services are non-taxable in this business's use case
- **Numbering:** Single `InvoiceEstimates` list distinguishes Estimates vs. Invoices via a `DocType` choice column
