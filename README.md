# Expense Splitter

A small desktop app for tracking shared expenses, splitting costs fairly, and calculating who owes whom after a group trip, household purchase, or shared event.

The project is implemented as a single Python script using Tkinter for the interface and `openpyxl` for Excel import/export.

## Features

- Add a list of participants
- Record each expense with a title, amount, description, and payer
- Automatically calculate each person's total paid, equal share, and net balance
- Suggest a settlement flow to clear debts between participants
- Import an existing workbook exported by the app
- Export the results and raw expense data to an Excel workbook

## Project structure

```text
Expense_Splitter/
├── Expense_Splitter.py   # Tkinter app and Excel logic
├── README.md             # Project documentation
```

## Prerequisites

- Python 3.9 or newer
- Windows/macOS/Linux with a desktop environment for Tkinter
- `openpyxl` package

## Installation

From the project root, install the required dependency:

```bash
py -m pip install openpyxl
```

If you are using `python` rather than `py`, replace that command with:

```bash
python -m pip install openpyxl
```

## Running the app

Start the app from the repository root:

```bash
py Expense_Splitter.py
```

Or:

```bash
python Expense_Splitter.py
```

## How it works

1. Enter participant names in the format `Alice, Bob, Charlie`.
2. Add expenses using the form with the title, amount, description, and payer.
3. Click `Calculate Settlement` to view the totals and suggested transfers.
4. Export the workbook to save the data and settlement summary to Excel.
5. Optionally load a previously exported workbook to continue editing an existing expense list.

## Excel workbook output

When you export, the app creates an Excel file with these sheets:

- `Expenses`: all recorded expenses and the running total
- `Individual Balances`: per-person paid amount, split share, and net balance
- `Settlement`: suggested payments to settle balances
- `Read Me`: notes explaining the calculation logic

## Notes

- All expenses are split equally among the participants currently entered in the app.
- The payer is the person who paid the expense and is treated as the owner of that expense.
- Settlement transfers are generated as a suggested way to clear all balances.
- The app rounds currency values to cents using standard financial rounding.

## Example workflow

```text
Set participants: Alice, Bob, Charlie
Add expense: "Groceries" - €120.00 - paid by Alice
Add expense: "Dinner" - €90.00 - paid by Bob
Calculate settlement
Export Excel workbook
```

## Dependencies

This project depends on:

- `tkinter` (included with most Python installs)
- `openpyxl` (for reading and writing Excel files)

## Known limitation

This is a local desktop application and does not include network storage, multi-user syncing, or a web interface.
