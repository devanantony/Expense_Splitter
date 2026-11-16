"""
Expense Splitter & Settlement Calculator
----------------------------------------
Install dependency:
    py -m pip install openpyxl

Run:
    py expense_splitter.py

All expenses are split equally among the participants listed at the top.
The "Owner" of an expense is the person who paid it.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import datetime
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("Missing dependency. Install it with: py -m pip install openpyxl")


CENT = Decimal("0.01")


def money(value):
    return Decimal(str(value)).quantize(CENT, rounding=ROUND_HALF_UP)


class ExpenseSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Expense Splitter & Settlement Calculator")
        self.root.geometry("1120x820")
        self.root.minsize(900, 620)

        self.expenses = []
        self.participants = []
        self.loaded_file = None

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 17, "bold"))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview", rowheight=26, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Expense Splitter", style="Title.TLabel").pack(anchor="w")
        header_row = ttk.Frame(outer)
        header_row.pack(fill="x", pady=(2, 12))
        ttk.Label(
            header_row,
            text="Record shared expenses, calculate who owes whom, and export everything to Excel."
        ).pack(side="left", anchor="w")
        # Import an existing workbook and export the updated version.
        ttk.Button(
            header_row, text="Load Existing Excel", command=self.load_excel
        ).pack(side="right", padx=(8, 0))
        ttk.Button(
            header_row, text="Export Excel", command=self.export_excel
        ).pack(side="right", padx=(8, 0))

        people_frame = ttk.LabelFrame(
            outer, text="1. Participants", style="Section.TLabelframe", padding=10
        )
        people_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(people_frame, text="Participant names (comma-separated):").pack(anchor="w")
        people_row = ttk.Frame(people_frame)
        people_row.pack(fill="x", pady=(5, 0))
        self.people_var = tk.StringVar()
        self.people_entry = ttk.Entry(people_row, textvariable=self.people_var)
        self.people_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(people_row, text="Set Participants", command=self.set_participants).pack(
            side="left", padx=(8, 0)
        )
        self.people_status = ttk.Label(people_frame, text="No participants added.")
        self.people_status.pack(anchor="w", pady=(5, 0))

        entry_frame = ttk.LabelFrame(
            outer, text="2. Add an expense", style="Section.TLabelframe", padding=10
        )
        entry_frame.pack(fill="x", pady=(0, 10))

        grid = ttk.Frame(entry_frame)
        grid.pack(fill="x")
        ttk.Label(grid, text="Expense title").grid(row=0, column=0, sticky="w")
        ttk.Label(grid, text="Amount (€)").grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Label(grid, text="Description").grid(row=0, column=2, sticky="w", padx=(8, 0))
        ttk.Label(grid, text="Owner / payer").grid(row=0, column=3, sticky="w", padx=(8, 0))

        self.title_var = tk.StringVar()
        self.amount_var = tk.StringVar()
        self.description_var = tk.StringVar()
        self.owner_var = tk.StringVar()

        self.title_entry = ttk.Entry(grid, textvariable=self.title_var)
        self.amount_entry = ttk.Entry(grid, textvariable=self.amount_var, width=14)
        self.description_entry = ttk.Entry(grid, textvariable=self.description_var)
        self.owner_combo = ttk.Combobox(
            grid, textvariable=self.owner_var, state="readonly", width=20
        )

        self.title_entry.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.amount_entry.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(4, 0))
        self.description_entry.grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=(4, 0))
        self.owner_combo.grid(row=1, column=3, sticky="ew", padx=(8, 0), pady=(4, 0))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(2, weight=2)

        ttk.Button(entry_frame, text="Add Expense", command=self.add_expense).pack(
            anchor="e", pady=(8, 0)
        )

        list_frame = ttk.LabelFrame(
            outer, text="3. Expense list", style="Section.TLabelframe", padding=8
        )
        list_frame.pack(fill="both", expand=True, pady=(0, 10))

        columns = ("title", "amount", "description", "owner")
        self.expense_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        headings = {
            "title": ("Expense title", 190),
            "amount": ("Amount (€)", 110),
            "description": ("Description", 440),
            "owner": ("Paid by", 160),
        }
        for col, (heading, width) in headings.items():
            self.expense_tree.heading(col, text=heading)
            self.expense_tree.column(col, width=width, minwidth=70, anchor="w")
        self.expense_tree.column("amount", anchor="e")
        self.expense_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.expense_tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.expense_tree.configure(yscrollcommand=scrollbar.set)

        list_actions = ttk.Frame(outer)
        list_actions.pack(fill="x", pady=(0, 10))
        self.total_label = ttk.Label(list_actions, text="Total expenses: €0.00")
        self.total_label.pack(side="left")
        ttk.Button(list_actions, text="Remove Selected", command=self.remove_selected).pack(
            side="right"
        )
        ttk.Button(list_actions, text="Clear All", command=self.clear_all).pack(
            side="right", padx=8
        )

        result_frame = ttk.LabelFrame(
            outer, text="4. Settlement summary", style="Section.TLabelframe", padding=8
        )
        result_frame.pack(fill="both", expand=True, pady=(0, 10))

        result_columns = ("person", "paid", "share", "net")
        self.result_tree = ttk.Treeview(
            result_frame, columns=result_columns, show="headings", height=5
        )
        result_headings = {
            "person": ("Participant", 220),
            "paid": ("Total paid (€)", 160),
            "share": ("Individual share (€)", 180),
            "net": ("Net balance (€)", 180),
        }
        for col, (heading, width) in result_headings.items():
            self.result_tree.heading(col, text=heading)
            self.result_tree.column(col, width=width, anchor="e" if col != "person" else "w")
        self.result_tree.pack(fill="both", expand=True)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x")
        self.settlement_label = ttk.Label(
            bottom, text="Add participants and expenses, then calculate settlement."
        )
        self.settlement_label.pack(side="left", fill="x", expand=True)
        ttk.Button(
            bottom, text="Calculate Settlement", command=self.calculate_settlement
        ).pack(side="right", padx=(8, 0))
        ttk.Button(bottom, text="Export Excel", command=self.export_excel).pack(side="right")

    def load_excel(self):
        """Load expenses and participant names from a previously exported workbook."""
        path = filedialog.askopenfilename(
            title="Open existing expense workbook",
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not path:
            return

        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            if "Expenses" not in wb.sheetnames:
                raise ValueError("This workbook does not contain an 'Expenses' sheet.")

            ws = wb["Expenses"]
            rows = ws.iter_rows(values_only=True)
            header = next(rows, None)
            if not header:
                raise ValueError("The Expenses sheet is empty.")

            headers = {str(value).strip().casefold(): index
                       for index, value in enumerate(header) if value is not None}
            title_col = headers.get("expense title", headers.get("title"))
            amount_col = headers.get("amount (€)", headers.get("amount"))
            desc_col = headers.get("description")
            owner_col = headers.get("paid by", headers.get("owner"))
            if None in (title_col, amount_col, owner_col):
                raise ValueError(
                    "The Expenses sheet must have columns named Expense Title, Amount (€), and Paid By."
                )

            loaded_expenses = []
            payer_names = []
            for row in rows:
                if not row or len(row) <= max(title_col, amount_col, owner_col):
                    continue
                title = row[title_col]
                amount_value = row[amount_col]
                owner = row[owner_col]
                # Ignore blank rows and the TOTAL summary row.
                if title is None or str(title).strip() == "" or str(title).strip().upper() == "TOTAL":
                    continue
                if amount_value is None or owner is None:
                    continue
                amount = money(amount_value)
                if amount <= 0:
                    continue
                description = row[desc_col] if desc_col is not None and len(row) > desc_col else ""
                owner = str(owner).strip()
                loaded_expenses.append({
                    "title": str(title).strip(),
                    "amount": amount,
                    "description": "" if description is None else str(description),
                    "owner": owner,
                })
                if owner and owner.casefold() not in {n.casefold() for n in payer_names}:
                    payer_names.append(owner)

            loaded_participants = []
            if "Individual Balances" in wb.sheetnames:
                balance_rows = wb["Individual Balances"].iter_rows(values_only=True)
                balance_header = next(balance_rows, None)
                if balance_header:
                    first_header = str(balance_header[0] or "").strip().casefold()
                    if first_header == "participant":
                        for row in balance_rows:
                            if row and row[0] is not None:
                                name = str(row[0]).strip()
                                if name and name.casefold() not in {n.casefold() for n in loaded_participants}:
                                    loaded_participants.append(name)
            wb.close()

            if not loaded_participants:
                loaded_participants = payer_names[:]
            # Include any payer missing from the participant summary.
            for name in payer_names:
                if name.casefold() not in {p.casefold() for p in loaded_participants}:
                    loaded_participants.append(name)
            if len(loaded_participants) < 2:
                raise ValueError(
                    "Could not find at least two participants. Use a workbook exported by this app, "
                    "or ensure the Individual Balances sheet lists all participants."
                )
            if not loaded_expenses:
                raise ValueError("No valid expense rows were found in the Expenses sheet.")

        except Exception as exc:
            try:
                wb.close()
            except Exception:
                pass
            messagebox.showerror("Could not load workbook", str(exc))
            return

        # Replace the current UI data only after the workbook has been validated.
        self.expenses = loaded_expenses
        self.participants = loaded_participants
        self.people_var.set(", ".join(self.participants))
        self.owner_combo["values"] = self.participants
        self.owner_var.set(self.participants[0])
        self.people_status.config(
            text=f"{len(self.participants)} participants: {', '.join(self.participants)}"
        )
        for item in self.expense_tree.get_children():
            self.expense_tree.delete(item)
        for expense in self.expenses:
            self.expense_tree.insert(
                "", "end",
                values=(expense["title"], f"{expense['amount']:.2f}",
                        expense["description"], expense["owner"])
            )
        self.loaded_file = path
        self._update_total()
        self._clear_results()
        messagebox.showinfo(
            "Workbook loaded",
            f"Loaded {len(self.expenses)} expenses for {len(self.participants)} participants.\n\n"
            "You can now add more expenses and export the updated workbook."
        )

    def set_participants(self):
        raw = self.people_var.get().strip()
        names = [name.strip() for name in raw.split(",") if name.strip()]
        # Remove duplicates while preserving order (case-insensitive).
        unique = []
        seen = set()
        for name in names:
            key = name.casefold()
            if key not in seen:
                unique.append(name)
                seen.add(key)

        if len(unique) < 2:
            messagebox.showwarning("Participants required", "Please enter at least two participant names.")
            return

        if self.expenses:
            current = {p.casefold() for p in self.participants}
            updated = {p.casefold() for p in unique}
            if current != updated:
                ok = messagebox.askyesno(
                    "Update participants?",
                    "Changing participants affects every expense split. Continue?"
                )
                if not ok:
                    return

        self.participants = unique
        self.owner_combo["values"] = self.participants
        if self.owner_var.get() not in self.participants:
            self.owner_var.set(self.participants[0])
        self.people_status.config(
            text=f"{len(self.participants)} participants: {', '.join(self.participants)}"
        )
        self._clear_results()

    def add_expense(self):
        if len(self.participants) < 2:
            messagebox.showwarning("Participants required", "Set at least two participants first.")
            return

        title = self.title_var.get().strip()
        amount_text = self.amount_var.get().strip().replace(",", ".")
        description = self.description_var.get().strip()
        owner = self.owner_var.get().strip()

        if not title:
            messagebox.showwarning("Missing title", "Enter an expense title.")
            self.title_entry.focus_set()
            return
        try:
            amount = money(Decimal(amount_text))
            if amount <= 0:
                raise InvalidOperation
        except (InvalidOperation, ValueError):
            messagebox.showwarning("Invalid amount", "Enter a valid amount greater than zero, e.g. 42.50.")
            self.amount_entry.focus_set()
            return
        if owner not in self.participants:
            messagebox.showwarning("Invalid owner", "Choose the person who paid this expense.")
            return

        expense = {
            "title": title,
            "amount": amount,
            "description": description,
            "owner": owner,
        }
        self.expenses.append(expense)
        self.expense_tree.insert(
            "", "end",
            values=(title, f"{amount:.2f}", description, owner)
        )
        self.title_var.set("")
        self.amount_var.set("")
        self.description_var.set("")
        self.title_entry.focus_set()
        self._update_total()
        self._clear_results()

    def remove_selected(self):
        selected = self.expense_tree.selection()
        if not selected:
            messagebox.showinfo("Select an expense", "Select one or more expenses to remove.")
            return
        indices = sorted((self.expense_tree.index(item) for item in selected), reverse=True)
        for item in selected:
            self.expense_tree.delete(item)
        for index in indices:
            del self.expenses[index]
        self._update_total()
        self._clear_results()

    def clear_all(self):
        if not self.expenses:
            return
        if messagebox.askyesno("Clear expenses", "Remove all entered expenses?"):
            self.expenses.clear()
            for item in self.expense_tree.get_children():
                self.expense_tree.delete(item)
            self._update_total()
            self._clear_results()

    def _update_total(self):
        total = sum((e["amount"] for e in self.expenses), Decimal("0.00"))
        self.total_label.config(text=f"Total expenses: €{total:.2f}")

    def _clear_results(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.settlement_label.config(text="Settlement not calculated yet.")

    def _calculate(self):
        if len(self.participants) < 2:
            raise ValueError("Please set at least two participants.")
        if not self.expenses:
            raise ValueError("Please add at least one expense.")

        total = sum((e["amount"] for e in self.expenses), Decimal("0.00"))
        paid = {p: Decimal("0.00") for p in self.participants}
        for expense in self.expenses:
            paid[expense["owner"]] += expense["amount"]

        # Split each individual expense in cents, distributing remainder cents
        # deterministically so each expense's allocated shares add up exactly.
        shares = {p: Decimal("0.00") for p in self.participants}
        for expense in self.expenses:
            cents = int(expense["amount"] * 100)
            base, remainder = divmod(cents, len(self.participants))
            for index, person in enumerate(self.participants):
                allocated = base + (1 if index < remainder else 0)
                shares[person] += Decimal(allocated) / 100

        net = {p: paid[p] - shares[p] for p in self.participants}

        # Greedy transfer matching: debtors pay creditors until all balances settle.
        debtors = [[p, -v] for p, v in net.items() if v < 0]
        creditors = [[p, v] for p, v in net.items() if v > 0]
        transfers = []
        i = j = 0
        while i < len(debtors) and j < len(creditors):
            amount = min(debtors[i][1], creditors[j][1])
            amount = money(amount)
            if amount > 0:
                transfers.append((debtors[i][0], creditors[j][0], amount))
            debtors[i][1] = money(debtors[i][1] - amount)
            creditors[j][1] = money(creditors[j][1] - amount)
            if debtors[i][1] == 0:
                i += 1
            if creditors[j][1] == 0:
                j += 1

        return total, paid, shares, net, transfers

    def calculate_settlement(self):
        try:
            total, paid, shares, net, transfers = self._calculate()
        except ValueError as exc:
            messagebox.showwarning("Cannot calculate", str(exc))
            return

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        for person in self.participants:
            self.result_tree.insert(
                "", "end",
                values=(
                    person,
                    f"{paid[person]:.2f}",
                    f"{shares[person]:.2f}",
                    f"{net[person]:+.2f}",
                )
            )

        if transfers:
            transfer_text = "  |  ".join(
                f"{payer} pays {receiver} €{amount:.2f}"
                for payer, receiver, amount in transfers
            )
        else:
            transfer_text = "Everyone is settled."
        self.settlement_label.config(text=transfer_text)

    def export_excel(self):
        try:
            total, paid, shares, net, transfers = self._calculate()
        except ValueError as exc:
            messagebox.showwarning("Cannot export", str(exc))
            return

        default_name = (Path(self.loaded_file).name if self.loaded_file
                        else f"expense_settlement_{datetime.now():%Y%m%d_%H%M}.xlsx")
        initial_dir = str(Path(self.loaded_file).parent) if self.loaded_file else None
        path = filedialog.asksaveasfilename(
            title="Save updated Excel workbook",
            defaultextension=".xlsx",
            initialfile=default_name,
            initialdir=initial_dir,
            filetypes=[("Excel workbook", "*.xlsx")]
        )
        if not path:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Expenses"
        ws.append(["Expense Title", "Amount (€)", "Description", "Paid By"])
        for e in self.expenses:
            ws.append([e["title"], float(e["amount"]), e["description"], e["owner"]])
        ws.append(["", "", "", ""])
        ws.append(["TOTAL", float(total), "", ""])

        balances = wb.create_sheet("Individual Balances")
        balances.append(["Participant", "Total Paid (€)", "Individual Share (€)", "Net Balance (€)", "Status"])
        for person in self.participants:
            status = "Gets back" if net[person] > 0 else ("Owes" if net[person] < 0 else "Settled")
            balances.append([
                person, float(paid[person]), float(shares[person]), float(net[person]), status
            ])

        settlement = wb.create_sheet("Settlement")
        settlement.append(["Who Pays", "Who Receives", "Amount (€)"])
        for payer, receiver, amount in transfers:
            settlement.append([payer, receiver, float(amount)])
        if not transfers:
            settlement.append(["", "Everyone is settled", 0.0])

        notes = wb.create_sheet("Read Me")
        notes.append(["Expense Splitter - Calculation Notes"])
        notes.append(["All expenses are split equally among the participants listed in the UI."])
        notes.append(["The owner/payer is the person who initially paid each expense."])
        notes.append(["Positive net balance = participant should receive money."])
        notes.append(["Negative net balance = participant needs to pay money."])
        notes.append(["Settlement transfers are a suggested way to clear all balances."])
        notes.append(["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")])

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        thin_gray = Side(style="thin", color="D9E2F3")

        for sheet in wb.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions if sheet.max_row > 1 else "A1"
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(vertical="center")
            for row in sheet.iter_rows():
                for cell in row:
                    cell.border = Border(bottom=thin_gray)
                    if cell.column > 1 and isinstance(cell.value, (int, float)):
                        cell.number_format = '€#,##0.00;[Red]-€#,##0.00'
            for col_cells in sheet.columns:
                max_length = max(len(str(c.value or "")) for c in col_cells)
                sheet.column_dimensions[get_column_letter(col_cells[0].column)].width = min(
                    max(max_length + 3, 14), 55
                )
            sheet.row_dimensions[1].height = 23

        # Keep the Read Me notes readable and avoid currency formatting on its numeric date row.
        notes.column_dimensions["A"].width = 100

        try:
            wb.save(path)
        except OSError as exc:
            messagebox.showerror("Save failed", f"Could not save the workbook:\n{exc}")
            return

        self.loaded_file = path
        messagebox.showinfo("Excel exported", f"Excel workbook saved successfully:\n{path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExpenseSplitterApp(root)
    root.mainloop()
