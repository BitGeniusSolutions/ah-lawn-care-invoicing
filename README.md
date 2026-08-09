# A&H Lawn Care Services — Invoicing & Estimates App

Documentation and design artifacts for a simple Power Platform (Power Apps + SharePoint) application to manage estimates, invoices, customers, and services for **A&H Lawn Care Services**.

## Contents

- [`docs/Invoicing-App-Design.md`](docs/Invoicing-App-Design.md) — Full design document: SharePoint list schema (Customers, Services, Documents, DocumentLines), Power Apps screen structure, Power Automate flow plan, and confirmed scope decisions.

## Project Summary

- **Stack:** Power Apps (Canvas app) + SharePoint Lists + Power Automate
- **Scope:** Single-user internal tool for creating/managing estimates and invoices, tracking customers, and maintaining a reusable service/price list
- **Design approach:** SharePoint Form controls for simple CRUD (Customers, Services, Document headers) with a lightweight editable gallery for invoice/estimate line items
- **Tax handling:** Not tracked — services are non-taxable in this business's use case
- **Numbering:** Single `Documents` list distinguishes Estimates vs. Invoices via a `DocType` choice column
