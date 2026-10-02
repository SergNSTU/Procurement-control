# Application modules

`RRFQComparer.ps1` remains the executable entry point. Shared functions are
dot-sourced from this directory so the application can be separated gradually
without changing its launch command or UI behaviour.

- `AppConfig.ps1` — supplier configuration and matching helpers.
- `UiTheme.ps1` — reusable WinForms styling and application branding.
- `UiFoundation.ps1` — common WinForms layout, icon, card, and toast helpers.
- `RrfqUi.ps1` — RRFQ preview helpers and manual matching dialogs.
- `RrfqEngine.ps1` — Excel/PDF parsing, quote matching, winner selection, and RRFQ export.
- `PurchaseStore.ps1` — SQLite access, documents, deals, suppliers, tasks, and reminders.
- `PurchaseDashboard.ps1` — dashboard metrics and monthly cohort calculations.
- `QuoteHistory.ps1` — quote history and Globalist import/query operations.
- `CompelParser.ps1` — Compel HTML parsing and Excel exports.
- `PriceSearch.ps1` — price-source queries, normalization, ranking, and GPT prompts.

New modules must not start the application or read command-line parameters.
They may depend on functions loaded by `RRFQComparer.ps1` at runtime.
