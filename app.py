from pathlib import Path
import os
import tkinter as tk
from tkinter import ttk, messagebox

from employee_app.excel_store import DuplicateEmployeeError, EmployeeExcelStore
from employee_app.service import validate_employee

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "data" / "employees.xlsx"

BG = "#F3F7F6"
CARD = "#FFFFFF"
PRIMARY = "#0F766E"
PRIMARY_DARK = "#115E59"
TEXT = "#1F2937"
MUTED = "#6B7280"
BORDER = "#DCE7E4"


class EmployeeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Data Entry System")
        self.root.geometry("1220x780")
        self.root.minsize(1040, 680)
        self.root.configure(bg=BG)
        self.store = EmployeeExcelStore(EXCEL_FILE)
        self._style()
        self._build()
        self.refresh_table()

    def _style(self):
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Header.TFrame", background=PRIMARY)
        style.configure("HeaderTitle.TLabel", background=PRIMARY, foreground="white", font=("Segoe UI", 22, "bold"))
        style.configure("HeaderSub.TLabel", background=PRIMARY, foreground="#DDF3EF", font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 13, "bold"))
        style.configure("Field.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10, "bold"))
        style.configure("Muted.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 9))
        style.configure("TEntry", padding=8)
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.map("Primary.TButton",
                  background=[("!disabled", PRIMARY), ("active", PRIMARY_DARK), ("pressed", PRIMARY_DARK)],
                  foreground=[("!disabled", "white")])
        style.configure("Secondary.TButton", padding=(12, 8))
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 10), background="white", fieldbackground="white", borderwidth=0)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#E8F3F1", foreground=TEXT, padding=6)
        style.map("Treeview", background=[("selected", "#CCFBF1")], foreground=[("selected", TEXT)])

    def _build(self):
        header = ttk.Frame(self.root, style="Header.TFrame", padding=(24, 18))
        header.pack(fill="x")
        ttk.Label(header, text="Employee Data Entry System", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(header, text="Professional employee record form with direct Excel save", style="HeaderSub.TLabel").pack(anchor="w", pady=(4, 0))

        body = ttk.Frame(self.root, style="TFrame", padding=18)
        body.pack(fill="both", expand=True)

        form_card = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        form_card.pack(fill="x", pady=(0, 14))
        form_inner = ttk.Frame(form_card, style="Card.TFrame", padding=18)
        form_inner.pack(fill="x")
        ttk.Label(form_inner, text="Employee Details", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))
        for c in (1, 3):
            form_inner.columnconfigure(c, weight=1)

        self.vars = {k: tk.StringVar() for k in ["name", "employee_id", "mobile", "aadhaar", "address", "email"]}
        self._field(form_inner, "Employee Name *", "name", 1, 0)
        self._field(form_inner, "Employee ID *", "employee_id", 1, 2)
        self._field(form_inner, "Mobile No *", "mobile", 2, 0)
        self._field(form_inner, "Aadhaar No *", "aadhaar", 2, 2)
        self._field(form_inner, "Email ID *", "email", 3, 0)

        ttk.Label(form_inner, text="Address *", style="Field.TLabel").grid(row=3, column=2, sticky="nw", padx=8, pady=8)
        self.address_text = tk.Text(form_inner, height=3, width=30, font=("Segoe UI", 10), wrap="word",
                                    relief="solid", borderwidth=1, highlightthickness=1,
                                    highlightbackground=BORDER, highlightcolor=PRIMARY, padx=8, pady=8)
        self.address_text.grid(row=3, column=3, sticky="ew", padx=8, pady=8)

        actions = ttk.Frame(form_inner, style="Card.TFrame")
        actions.grid(row=4, column=0, columnspan=4, sticky="e", pady=(12, 0))
        ttk.Button(actions, text="Save to Excel", style="Primary.TButton", command=self.save).pack(side="left", padx=4)
        ttk.Button(actions, text="Clear Form", style="Secondary.TButton", command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(actions, text="Open Excel Folder", style="Secondary.TButton", command=self.open_folder).pack(side="left", padx=4)

        table_card = tk.Frame(body, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        table_card.pack(fill="both", expand=True)
        table_inner = ttk.Frame(table_card, style="Card.TFrame", padding=14)
        table_inner.pack(fill="both", expand=True)

        top = ttk.Frame(table_inner, style="Card.TFrame")
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Saved Employee Records", style="CardTitle.TLabel").pack(side="left")
        search_wrap = ttk.Frame(top, style="Card.TFrame")
        search_wrap.pack(side="right")
        ttk.Label(search_wrap, text="Search:", style="Field.TLabel").pack(side="left", padx=(0, 6))
        self.search_var = tk.StringVar()
        search = ttk.Entry(search_wrap, textvariable=self.search_var, width=32)
        search.pack(side="left")
        search.bind("<KeyRelease>", lambda _e: self.refresh_table())
        ttk.Button(search_wrap, text="Clear", style="Secondary.TButton", command=self.clear_search).pack(side="left", padx=6)

        cols = ("name", "employee_id", "mobile", "aadhaar", "address", "email", "created_at")
        self.tree = ttk.Treeview(table_inner, columns=cols, show="headings")
        headings = {
            "name": "Employee Name", "employee_id": "Employee ID", "mobile": "Mobile No",
            "aadhaar": "Aadhaar No", "address": "Address", "email": "Email ID", "created_at": "Saved At"
        }
        widths = {"name": 170, "employee_id": 110, "mobile": 120, "aadhaar": 140, "address": 260, "email": 210, "created_at": 140}
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor="w" if c in ("name", "address", "email") else "center")
        y = ttk.Scrollbar(table_inner, orient="vertical", command=self.tree.yview)
        x = ttk.Scrollbar(table_inner, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y.set, xscrollcommand=x.set)
        self.tree.pack(side="left", fill="both", expand=True)
        y.pack(side="right", fill="y")

        self.status = ttk.Label(self.root, text=f"Excel File: {EXCEL_FILE}", style="Muted.TLabel", padding=(18, 0, 18, 10))
        self.status.pack(fill="x", side="bottom")

    def _field(self, parent, label, key, row, col):
        ttk.Label(parent, text=label, style="Field.TLabel").grid(row=row, column=col, sticky="w", padx=8, pady=8)
        entry = ttk.Entry(parent, textvariable=self.vars[key])
        entry.grid(row=row, column=col + 1, sticky="ew", padx=8, pady=8)

    def save(self):
        try:
            employee = validate_employee(
                self.vars["name"].get(),
                self.vars["employee_id"].get(),
                self.vars["mobile"].get(),
                self.vars["aadhaar"].get(),
                self.address_text.get("1.0", "end").strip(),
                self.vars["email"].get(),
            )
            self.store.add_employee(employee)
        except DuplicateEmployeeError as exc:
            messagebox.showwarning("Duplicate Employee ID", str(exc), parent=self.root)
            return
        except (ValueError, PermissionError) as exc:
            messagebox.showerror("Check Details", str(exc), parent=self.root)
            return
        except Exception as exc:
            messagebox.showerror("Save Error", f"Data save nahi hua:\n{exc}", parent=self.root)
            return
        self.refresh_table()
        self.clear_form()
        messagebox.showinfo("Saved", "Employee data Excel mein save ho gaya.", parent=self.root)

    def refresh_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        try:
            records = self.store.get_employees(self.search_var.get() if hasattr(self, "search_var") else "")
            for r in records:
                self.tree.insert("", "end", values=r)
            self.status.configure(text=f"Showing {len(records)} record(s) | Excel: {EXCEL_FILE}")
        except Exception as exc:
            messagebox.showerror("Read Error", str(exc), parent=self.root)

    def clear_form(self):
        for v in self.vars.values():
            v.set("")
        self.address_text.delete("1.0", "end")

    def clear_search(self):
        self.search_var.set("")
        self.refresh_table()

    def open_folder(self):
        EXCEL_FILE.parent.mkdir(parents=True, exist_ok=True)
        os.startfile(str(EXCEL_FILE.parent))


def main():
    root = tk.Tk()
    try:
        EmployeeApp(root)
    except Exception as exc:
        messagebox.showerror("Startup Error", str(exc))
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
