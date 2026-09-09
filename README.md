# Claims Interest Calculator

A C#/WPF desktop app that reproduces the "Daily Interest Calculator" Excel
workbook (Claims_Calculations_and_Cover_Letter.xlsm), including its cover
letter generation.

## What it does

- Calculates daily interest on a claim's outstanding loan balance from a
  **Start Date** to **Date of Death**, using either a single interest rate
  or a table of dated rate changes.
- Interest accrues daily (simple interest, `balance * rate% / 100 / 365`)
  and is capitalised into the balance at each **month-end** and on the
  **date of death** — matching the original spreadsheet's compounding rule.
- Deducts any payments received on their given dates.
- Produces the same two totals as the spreadsheet: **Total Interest** and
  **Claim Amount Incl. Interest**.
- Generates a formatted claim cover letter to Hollard Insurance Namibia,
  including the same document checklist logic as the original VBA
  (including the rule that suppresses "Burial Removal Order" from the
  Outstanding list under specific conditions).

## What's different from the spreadsheet

- **Client Name**, **Date of Death**, **Policy No.**, and **Loan Agreement
  Date** were originally auto-filled from an external "Hollard Data Base"
  workbook via `XLOOKUP`. That file wasn't available, so these are now
  manual entry fields in the **Claim Details** tab.
- The spreadsheet's daily table was hard-capped at 365 rows (365 days after
  the Start Date). This app has no such cap — it calculates the full
  Start Date → Date of Death span regardless of length.
- There were no "Fees" entries in the source workbook, so the Fees column
  is carried through as zero. It's easy to extend if you need it later.

## How to build and run

You'll need the .NET SDK installed (this project targets **.NET 10**, since
that's what was set up during this conversation).

From inside this folder:

```
dotnet build
dotnet run
```

Or open the folder in Visual Studio / VS Code and run from there.

## Project layout

| File | Purpose |
|---|---|
| `MainWindow.xaml` / `.cs` | The UI — tabs for details, rate changes, payments, checklist, results, and letter |
| `ClaimsCalculator.cs` | The interest calculation engine + document checklist logic |
| `CoverLetterGenerator.cs` | Builds the formatted cover letter text |
| `Models.cs` | Data classes used across the app |

## Using it

1. **Claim Details** — fill in client info, opening balance, dates, and base rate.
2. **Rate Changes** — only used if you check "Multiple Interest Rates" — add each rate change with its effective date.
3. **Payments Received** — add any payments made against the account during the period.
4. **Document Checklist** — tick off which claim documents you have.
5. Click **Calculate** (bottom left) — this fills in the **Results** tab.
6. Go to **Cover Letter** → **Generate Cover Letter**, then save to file or copy to clipboard.

## OneDrive Excel database

The **Load from SharePoint** button looks up a record by the `AZ` column in
the local OneDrive-synced `ClaimsData` Excel table and fills in the client
name, policy number, loan amount, loan start date, and incident date. The
**Database Record** tab exposes every column for editing, and the **Save to
Excel** button writes all fields back to the matching table row.
The local workbook path and table name are in `appsettings.json`, so they can
be changed later without changing the calculation code.

The app uses the installed desktop Excel application for this integration.
Excel must be closed for the workbook before saving through the app. The
workbook should have SharePoint version history and a server-side backup
enabled. Concurrent edits are not conflict-free: the app should be used by
one writer at a time until a SharePoint List, API, or database is available.
