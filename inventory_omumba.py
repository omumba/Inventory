#!/usr/bin/env python3
import os
import shutil
import hashlib
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd

# ---- Config: data files under data/ ----
DATA_DIR = os.path.join(os.path.abspath(os.path.dirname(__file__)), "data")
os.makedirs(DATA_DIR, exist_ok=True)
DATA_FILE = os.path.join(DATA_DIR, "Nkhokwe.xlsx")
AUTH_FILE = os.path.join(DATA_DIR, "Nkhokwe_Auth.xlsx")

INVENTORY_SHEET = "Inventory"
ISSUED_SHEET = "Issued"
HISTORY_SHEET = "History"
STAFF_SHEET = "Staff"
USERS_SHEET = "Users"
SETTINGS_SHEET = "Settings"

INVENTORY_COLUMNS = [
    "Product Name", "Category", "Serial Number", "Barcode",
    "Available", "Location", "Quantity", "Unit Price", "Low Stock Threshold"
]
ISSUED_COLUMNS = [
    "Product Name", "Category", "Serial Number", "Barcode",
    "Issued To", "Department", "Position", "Issue Date"
]
HISTORY_COLUMNS = ["Timestamp", "User", "Action", "Details"]
STAFF_COLUMNS = ["Username", "Department", "Position"]
USERS_COLUMNS = ["Username", "PasswordHash", "Role", "Active"]
SETTINGS_COLUMNS = ["Key", "Value"]

SHEETS_TO_FILE = {
    INVENTORY_SHEET: DATA_FILE,
    ISSUED_SHEET: DATA_FILE,
    HISTORY_SHEET: DATA_FILE,
    STAFF_SHEET: DATA_FILE,
    USERS_SHEET: AUTH_FILE,
    SETTINGS_SHEET: AUTH_FILE,
}
DATA_SPEC = {
    INVENTORY_SHEET: INVENTORY_COLUMNS,
    ISSUED_SHEET: ISSUED_COLUMNS,
    HISTORY_SHEET: HISTORY_COLUMNS,
    STAFF_SHEET: STAFF_COLUMNS,
}
AUTH_SPEC = {
    USERS_SHEET: USERS_COLUMNS,
    SETTINGS_SHEET: SETTINGS_COLUMNS,
}

ROLES = ("Admin", "Clerk", "Viewer")

# ---- Security helpers ----


def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()


def check_pw(pw: str, pw_hash: str) -> bool:
    return hash_pw(pw) == (str(pw_hash) if pw_hash is not None else "")

# ---- Workbook helpers ----


def _create_workbook(file_path: str, spec: dict, *, seed_auth=False):
    with pd.ExcelWriter(file_path, engine="openpyxl") as wr:
        for sheet, cols in spec.items():
            pd.DataFrame(columns=cols).to_excel(wr, sheet, index=False)
        if seed_auth:
            # add default admin user and some default settings
            pd.DataFrame([{
                "Username": "admin",
                "PasswordHash": hash_pw("Admin@2028"),
                "Role": "Admin",
                "Active": True
            }], columns=USERS_COLUMNS).to_excel(wr, USERS_SHEET, index=False)
            pd.DataFrame([
                {"Key": "camera_index", "Value": "0"},
                {"Key": "backup_enabled", "Value": "1"},
                {"Key": "ad_server", "Value": ""},
                {"Key": "ad_base_dn", "Value": ""},
                {"Key": "ad_user", "Value": ""},
            ], columns=SETTINGS_COLUMNS).to_excel(wr, SETTINGS_SHEET, index=False)


def _add_missing_sheets(file_path: str, spec: dict, seed_auth=False):
    try:
        existing = set(pd.ExcelFile(file_path).sheet_names)
    except Exception:
        # corrupt or unreadable -> recreate
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        bad = f"{os.path.splitext(file_path)[0]}_corrupt_{ts}.xlsx"
        try:
            shutil.move(file_path, bad)
        except Exception:
            pass
        _create_workbook(file_path, spec, seed_auth=seed_auth)
        return
    to_create = [s for s in spec.keys() if s not in existing]
    if to_create:
        with pd.ExcelWriter(file_path, engine="openpyxl", mode="a") as wr:
            for s in to_create:
                pd.DataFrame(columns=spec[s]).to_excel(wr, s, index=False)


def ensure_workbooks_exist():
    if not os.path.exists(DATA_FILE):
        _create_workbook(DATA_FILE, DATA_SPEC, seed_auth=False)
    else:
        _add_missing_sheets(DATA_FILE, DATA_SPEC, seed_auth=False)
    if not os.path.exists(AUTH_FILE):
        _create_workbook(AUTH_FILE, AUTH_SPEC, seed_auth=True)
    else:
        _add_missing_sheets(AUTH_FILE, AUTH_SPEC, seed_auth=True)


def get_file_for_sheet(sheet: str) -> str:
    return SHEETS_TO_FILE.get(sheet, DATA_FILE)


def settings_to_dict(df):
    s = {}
    if df is None or df.empty:
        return s
    for _, r in df.iterrows():
        k = str(r.get("Key", "")).strip()
        v = r.get("Value", "")
        s[k] = "" if pd.isna(v) else str(v)
    return s


def _backup_file_if_enabled(target_file: str):
    try:
        if os.path.exists(AUTH_FILE):
            sdf = pd.read_excel(AUTH_FILE, sheet_name=SETTINGS_SHEET)
            if "Key" in sdf.columns and "Value" in sdf.columns:
                if settings_to_dict(sdf).get("backup_enabled", "1") != "1":
                    return
    except Exception:
        pass
    if os.path.exists(target_file):
        try:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            root, ext = os.path.splitext(target_file)
            shutil.copy2(target_file, f"{root}_backup_{ts}{ext}")
        except Exception:
            pass

# ---- IO helpers ----


def load_sheet(sheet: str, required_cols):
    ensure_workbooks_exist()
    file_path = get_file_for_sheet(sheet)
    try:
        df = pd.read_excel(file_path, sheet_name=sheet)
    except Exception:
        df = pd.DataFrame(columns=required_cols)
    # ensure required cols present with defaults
    for col in required_cols:
        if col not in df.columns:
            if col == "Quantity":
                df[col] = 0
            elif col in ("Unit Price",):
                df[col] = 0.0
            elif col == "Available":
                df[col] = "Yes"
            elif col == "Active":
                df[col] = True
            else:
                df[col] = ""
    # normalize types
    if "Quantity" in df.columns:
        try:
            df["Quantity"] = pd.to_numeric(
                df["Quantity"], errors="coerce").fillna(0).astype(int)
        except Exception:
            df["Quantity"] = 0
    if "Unit Price" in df.columns:
        try:
            df["Unit Price"] = pd.to_numeric(
                df["Unit Price"], errors="coerce").fillna(0.0).astype(float)
        except Exception:
            df["Unit Price"] = 0.0
    if "Low Stock Threshold" in df.columns:
        try:
            df["Low Stock Threshold"] = pd.to_numeric(
                df["Low Stock Threshold"], errors="coerce").fillna(0).astype(int)
        except Exception:
            df["Low Stock Threshold"] = 0
    # Available normalization
    if "Available" in df.columns:
        try:
            df["Available"] = df["Available"].apply(lambda v: "Yes" if str(
                v).strip().lower() in ("true", "1", "yes", "y", "t") else "No")
        except Exception:
            df["Available"] = df["Available"].astype(str).fillna("Yes")
    # cleanup textual
    for c in ("Serial Number", "Barcode", "Product Name", "Category", "Location", "Username"):
        if c in df.columns:
            df[c] = df[c].astype(str).fillna("")
    if sheet == USERS_SHEET and "Active" in df.columns:
        df["Active"] = df["Active"].apply(lambda v: True if str(
            v).strip().lower() in ("true", "1", "yes", "y", "t") else False)
    return df


def save_sheet(df: pd.DataFrame, sheet: str):
    ensure_workbooks_exist()
    file_path = get_file_for_sheet(sheet)
    _backup_file_if_enabled(file_path)
    with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as wr:
        df.to_excel(wr, sheet_name=sheet, index=False)


def append_history(user, action, details):
    hist = load_sheet(HISTORY_SHEET, HISTORY_COLUMNS)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    hist = pd.concat([hist, pd.DataFrame(
        [{"Timestamp": now, "User": user or "", "Action": action, "Details": details}])], ignore_index=True)
    save_sheet(hist, HISTORY_SHEET)


def ensure_admin_user():
    users = load_sheet(USERS_SHEET, USERS_COLUMNS)
    mask = users["Username"].astype(str).str.lower() == "admin"
    default_hash = hash_pw("Admin@2028")
    if mask.any():
        i = users.index[mask][0]
        users.at[i, "Role"] = "Admin"
        users.at[i, "Active"] = True
        if not str(users.at[i, "PasswordHash"]).strip():
            users.at[i, "PasswordHash"] = default_hash
    else:
        users = pd.concat([users, pd.DataFrame(
            [{"Username": "admin", "PasswordHash": default_hash, "Role": "Admin", "Active": True}])], ignore_index=True)
        append_history("system", "SEED_ADMIN", "Built-in admin created")
    save_sheet(users, USERS_SHEET)


# ensure files + admin on start
ensure_workbooks_exist()
ensure_admin_user()

# ---- App ----


class InventoryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Omumba Tech Solution – Inventory System v1.0")
        # Start maximized (Windows supports 'zoomed'; fallback to fullscreen flag)
        try:
            self.state('zoomed')
        except Exception:
            try:
                self.attributes("-fullscreen", True)
            except Exception:
                pass

        # data frames
        self.inventory_df = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
        self.issued_df = load_sheet(ISSUED_SHEET, ISSUED_COLUMNS)
        self.staff_df = load_sheet(STAFF_SHEET, STAFF_COLUMNS)
        self.users_df = load_sheet(USERS_SHEET, USERS_COLUMNS)
        self.settings_df = load_sheet(SETTINGS_SHEET, SETTINGS_COLUMNS)
        self.current_user = None
        self.current_role = "Viewer"

        # UI state
        self.search_var = tk.StringVar()
        self.inv_row_map = {}
        self.iss_row_map = {}
        self.categories = sorted(self.inventory_df["Category"].replace(
            "", pd.NA).dropna().astype(str).unique().tolist())

        # build login first
        self._build_login()

    # ---- login ----
    def _build_login(self):
        self.login_frame = ttk.Frame(self, padding=20)
        self.login_frame.pack(expand=True)
        ttk.Label(self.login_frame, text="Welcome to Omumba Tech Solution Inventory System", font=(
            "Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 12))
        ttk.Label(self.login_frame, text="Username").grid(
            row=1, column=0, sticky="w", pady=6)
        ttk.Label(self.login_frame, text="Password").grid(
            row=2, column=0, sticky="w", pady=6)
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()
        ttk.Entry(self.login_frame, textvariable=self.username_var).grid(
            row=1, column=1, pady=6)
        ttk.Entry(self.login_frame, textvariable=self.password_var,
                  show="*").grid(row=2, column=1, pady=6)
        ttk.Button(self.login_frame, text="Login", command=self._login).grid(
            row=3, column=0, columnspan=2, pady=12)

    def _login(self):
        u = self.username_var.get().strip()
        p = self.password_var.get().strip()
        self.users_df = load_sheet(USERS_SHEET, USERS_COLUMNS)
        sel = self.users_df[(self.users_df["Username"] == u)
                            & (self.users_df["Active"] == True)]
        if not sel.empty and check_pw(p, sel.iloc[0]["PasswordHash"]):
            self.current_user = u
            self.current_role = sel.iloc[0]["Role"] if sel.iloc[0]["Role"] in ROLES else "Viewer"
            self.login_frame.destroy()
            append_history(self.current_user, "LOGIN",
                           f"Role={self.current_role}")
            self._build_main_ui()
        else:
            messagebox.showerror(
                "Login failed", "Invalid username or password.")

    # ---- permissions ----
    def is_admin(self): return self.current_role == "Admin"
    def is_clerk(self): return self.current_role in ("Admin", "Clerk")
    def is_viewer(self): return self.current_role == "Viewer"

    # ---- main UI ----
    def _build_main_ui(self):
        # left menu
        self.menu_frame = ttk.Frame(self, width=220)
        self.menu_frame.pack(side="left", fill="y")
        ttk.Button(self.menu_frame, text="Dashboard",
                   command=self.show_dashboard).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Inventory",
                   command=self.show_inventory).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Issued Items",
                   command=self.show_issued).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Import CSV", command=self.import_csv, state=(
            "normal" if self.is_clerk() else "disabled")).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Export CSV", command=self.export_csv, state=(
            "normal" if self.is_admin() else "disabled")).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Categories", command=self.manage_categories, state=(
            "normal" if self.is_admin() else "disabled")).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Staff", command=self.manage_staff, state=(
            "normal" if self.is_clerk() else "disabled")).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text="Settings", command=self.open_settings, state=(
            "normal" if self.is_admin() else "disabled")).pack(fill="x", padx=6, pady=4)
        ttk.Button(self.menu_frame, text=f"Logout ({self.current_user})", command=self.logout).pack(
            side="bottom", fill="x", padx=6, pady=8)

        # right area
        self.right_frame = ttk.Frame(self)
        self.right_frame.pack(side="left", fill="both", expand=True)

        top_row = ttk.Frame(self.right_frame)
        top_row.pack(fill="x", padx=6, pady=6)
        ttk.Label(top_row, text="Search:").pack(side="left")
        e = ttk.Entry(top_row, textvariable=self.search_var)
        e.pack(side="left", fill="x", expand=True, padx=6)
        e.bind("<KeyRelease>", lambda _e: self._refresh_inventory_table())

        self.table_frame = ttk.Frame(self.right_frame)
        self.table_frame.pack(fill="both", expand=True)
        self.status_var = tk.StringVar(
            value=f"Data: {DATA_FILE} | Auth: {AUTH_FILE}")
        ttk.Label(self.right_frame, textvariable=self.status_var,
                  anchor="w").pack(side="bottom", fill="x")

        self.show_dashboard()

    def _clear_content(self):
        for w in self.table_frame.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass

    def _open_inline(self, builder):
        self._clear_content()
        page = builder(self.table_frame)
        page.pack(fill="both", expand=True)
        return page

    # ---- Dashboard ----
    def show_dashboard(self):
        def build(parent):
            frm = ttk.Frame(parent, padding=12)
            inv = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
            total_items = len(inv)
            total_qty = int(inv["Quantity"].sum()) if not inv.empty else 0
            total_value = float(
                (inv["Quantity"] * inv["Unit Price"]).sum()) if not inv.empty else 0.0
            low_stock = int(
                (inv["Quantity"] <= inv["Low Stock Threshold"]).sum()) if not inv.empty else 0
            available_count = int(
                (inv["Available"] == "Yes").sum()) if not inv.empty else 0

            grid = ttk.Frame(frm)
            grid.pack(anchor="nw")

            def stat(r, label, value):
                ttk.Label(grid, text=label, font=("Segoe UI", 11, "bold")).grid(
                    row=r, column=0, sticky="w", padx=8, pady=6)
                ttk.Label(grid, text=value, font=("Segoe UI", 12)).grid(
                    row=r, column=1, sticky="w", padx=8, pady=6)
            stat(0, "Total SKUs", total_items)
            stat(1, "Total Quantity", total_qty)
            stat(2, "Total Inventory Value", f"{total_value:,.2f}")
            stat(3, "Low-Stock Items", low_stock)
            stat(4, "Available Items", available_count)

            actions = ttk.Frame(frm)
            actions.pack(pady=10)
            ttk.Button(actions, text="Go to Inventory",
                       command=self.show_inventory).pack(side="left", padx=6)
            ttk.Button(actions, text="View Unavailable Report",
                       command=self.export_unavailable).pack(side="left", padx=6)
            return frm
        self._open_inline(build)

    # ---- Inventory view ----
    def show_inventory(self):
        def build(parent):
            wrap = ttk.Frame(parent)
            cols = INVENTORY_COLUMNS + ["Stock Value"]
            self.inv_tree = ttk.Treeview(wrap, columns=cols, show="headings")
            for c in cols:
                self.inv_tree.heading(c, text=c)
                self.inv_tree.column(c, width=120 if c not in (
                    "Product Name", "Location") else 180)
            self.inv_tree.pack(fill="both", expand=True)
            btns = ttk.Frame(wrap)
            btns.pack(fill="x", pady=8)
            ttk.Button(btns, text="Add Item", command=self.add_item_form, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=4)
            ttk.Button(btns, text="Edit Item", command=self.edit_item_form, state=(
                "normal" if self.is_admin() else "disabled")).pack(side="left", padx=4)
            ttk.Button(btns, text="Delete Item", command=self.delete_item, state=(
                "normal" if self.is_admin() else "disabled")).pack(side="left", padx=4)
            ttk.Button(btns, text="Issue Item", command=self.issue_item_form, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=8)
            ttk.Button(btns, text="Batch Issue", command=self.batch_issue, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=4)
            ttk.Button(btns, text="Save Label", command=self.save_label, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=8)
            self._refresh_inventory_table()
            return wrap
        self._open_inline(build)

    def _filter_inventory(self, df):
        q = self.search_var.get().strip().lower()
        if not q:
            return df
        mask = (
            df["Product Name"].astype(str).str.lower().str.contains(q, na=False) |
            df["Category"].astype(str).str.lower().str.contains(q, na=False) |
            df["Serial Number"].astype(str).str.lower().str.contains(q, na=False) |
            df["Barcode"].astype(str).str.lower().str.contains(q, na=False) |
            df["Location"].astype(str).str.lower().str.contains(q, na=False)
        )
        return df[mask]

    def _refresh_inventory_table(self):
        if not hasattr(self, "inv_tree"):
            return
        for r in self.inv_tree.get_children():
            self.inv_tree.delete(r)
        self.inv_row_map.clear()
        self.inventory_df = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
        view = self._filter_inventory(self.inventory_df).copy()
        view["Stock Value"] = (view["Quantity"] * view["Unit Price"]).round(2)
        for idx, r in view.iterrows():
            vals = [r.get(c, "") for c in INVENTORY_COLUMNS] + \
                [r.get("Stock Value", 0.0)]
            tags = ()
            if str(r.get("Available", "Yes")).strip().lower() != "yes":
                tags = ("notavail",)
            iid = self.inv_tree.insert("", "end", values=vals, tags=tags)
            self.inv_row_map[iid] = idx
        self.inv_tree.tag_configure("notavail", foreground="#888888")
        self.status_var.set(
            f"Items: {len(view)} | Data: {DATA_FILE} | Auth: {AUTH_FILE}")

    def _get_selected_inventory_index(self):
        if not hasattr(self, "inv_tree"):
            return None
        sel = self.inv_tree.selection()
        if not sel:
            return None
        return self.inv_row_map.get(sel[0])

    # ---- Add / Edit / Delete items ----
    def add_item_form(self): self._item_form("add")

    def edit_item_form(self):
        idx = self._get_selected_inventory_index()
        if idx is None:
            messagebox.showinfo("Select", "Select an item to edit.")
            return
        self._item_form("edit", idx)

    def delete_item(self):
        idx = self._get_selected_inventory_index()
        if idx is None:
            messagebox.showinfo("Select", "Select an item to delete.")
            return
        row = self.inventory_df.loc[idx]
        sn = row.get("Serial Number", "")
        if not messagebox.askyesno("Confirm Delete", f"Delete item with Serial Number: {sn}?"):
            return
        self.inventory_df = self.inventory_df.drop(
            index=idx).reset_index(drop=True)
        save_sheet(self.inventory_df, INVENTORY_SHEET)
        append_history(self.current_user, "DELETE", f"SN:{sn}")
        self._refresh_inventory_table()
        messagebox.showinfo(
            "Deleted", f"Item with Serial Number {sn} deleted.")

    def _get_category_choices_list(self):
        try:
            self.inventory_df = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
            cats = sorted(self.inventory_df["Category"].replace(
                "", pd.NA).dropna().astype(str).unique().tolist())
            self.categories = cats
        except Exception:
            self.categories = []
        return self.categories

    def _item_form(self, mode="add", df_index=None):
        def build(parent):
            form = ttk.Frame(parent)
            header = ttk.Frame(form)
            header.pack(fill="x", pady=(0, 8))
            ttk.Button(header, text="← Back",
                       command=self.show_inventory).pack(side="left")
            ttk.Label(header, text=("Add Item" if mode == "add" else "Edit Item"), font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)

            fields = INVENTORY_COLUMNS
            vars_ = {c: tk.StringVar() for c in fields}
            if mode == "edit" and df_index is not None:
                row = self.inventory_df.loc[df_index]
                for c in fields:
                    vars_[c].set(str(row.get(c, "")))

            grid = ttk.Frame(form)
            grid.pack(fill="x", padx=8, pady=6)
            for i, c in enumerate(fields):
                ttk.Label(grid, text=c).grid(
                    row=i, column=0, sticky="w", padx=6, pady=4)
                if c == "Category":
                    cat_values = self._get_category_choices_list()
                    cat_cb = ttk.Combobox(grid, textvariable=vars_[
                                          "Category"], values=cat_values, state="readonly")
                    cat_cb.grid(row=i, column=1, sticky="we", padx=6, pady=4)

                    def add_category_inline():
                        new_cat = simpledialog.askstring(
                            "New Category", "Enter new category name:", parent=self)
                        if new_cat:
                            new_cat = new_cat.strip()
                            if new_cat:
                                # update inventory categories immediately (no row modification)
                                self.categories = sorted(
                                    set(cat_values + [new_cat]))
                                cat_cb["values"] = self.categories
                                vars_["Category"].set(new_cat)
                    ttk.Button(
                        grid, text="+", width=3, command=add_category_inline).grid(row=i, column=2, padx=4)
                else:
                    ent = ttk.Entry(grid, textvariable=vars_[c])
                    ent.grid(row=i, column=1, sticky="we", padx=6, pady=4)
            grid.columnconfigure(1, weight=1)

            footer = ttk.Frame(form)
            footer.pack(fill="x", pady=10)

            def save_item():
                try:
                    new_row = {c: vars_[c].get().strip() for c in fields}
                    # types/defaults
                    try:
                        new_row["Quantity"] = int(
                            float(new_row.get("Quantity") or 0))
                    except Exception:
                        new_row["Quantity"] = 0
                    try:
                        new_row["Unit Price"] = float(
                            new_row.get("Unit Price") or 0.0)
                    except Exception:
                        new_row["Unit Price"] = 0.0
                    try:
                        new_row["Low Stock Threshold"] = int(
                            float(new_row.get("Low Stock Threshold") or 0))
                    except Exception:
                        new_row["Low Stock Threshold"] = 0
                    if new_row.get("Available", "") == "":
                        new_row["Available"] = "Yes"
                    if not new_row["Serial Number"] and not new_row["Barcode"]:
                        messagebox.showerror(
                            "Error", "Enter Serial Number or Barcode.")
                        return
                    # merge or add
                    ix = self._find_item_index(
                        new_row["Serial Number"], new_row["Barcode"])
                    if mode == "add":
                        if ix is not None:
                            # merge: update non-empty fields and add qty
                            self.inventory_df.at[ix, "Quantity"] = int(
                                self.inventory_df.at[ix, "Quantity"]) + new_row["Quantity"]
                            for col in ["Product Name", "Category", "Location", "Unit Price", "Low Stock Threshold", "Barcode", "Serial Number", "Available"]:
                                v = new_row.get(col, "")
                                if v != "":
                                    self.inventory_df.at[ix, col] = v
                            append_history(
                                self.current_user, "EDIT", f"Merge add SN:{new_row['Serial Number']} BC:{new_row['Barcode']}")
                        else:
                            # ensure all columns present
                            for c in INVENTORY_COLUMNS:
                                if c not in new_row:
                                    new_row[c] = "" if c != "Quantity" else 0
                            self.inventory_df = pd.concat(
                                [self.inventory_df, pd.DataFrame([new_row])], ignore_index=True)
                            append_history(
                                self.current_user, "ADD", f"SN:{new_row['Serial Number']} BC:{new_row['Barcode']}")
                    else:
                        # edit
                        for k, v in new_row.items():
                            if k in self.inventory_df.columns:
                                self.inventory_df.at[df_index, k] = v
                        append_history(self.current_user,
                                       "EDIT", f"idx={df_index}")
                    save_sheet(self.inventory_df, INVENTORY_SHEET)
                    # ensure categories updated
                    self.categories = sorted(self.inventory_df["Category"].replace(
                        "", pd.NA).dropna().astype(str).unique().tolist())
                    self._refresh_inventory_table()
                    self.show_inventory()
                except Exception as e:
                    messagebox.showerror("Error", str(e))
            ttk.Button(footer, text="Save",
                       command=save_item).pack(side="right")
            return form
        self._open_inline(build)

    def _find_item_index(self, serial: str, barcode: str):
        if serial:
            m = self.inventory_df.index[self.inventory_df["Serial Number"] == str(
                serial)]
            if len(m) > 0:
                return int(m[0])
        if barcode:
            m = self.inventory_df.index[self.inventory_df["Barcode"] == str(
                barcode)]
            if len(m) > 0:
                return int(m[0])
        return None

    # ---- Issue / Return ----
    def issue_item_form(self):
        def build(parent):
            frm = ttk.Frame(parent, padding=8)
            header = ttk.Frame(frm)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_inventory).pack(side="left")
            ttk.Label(header, text="Issue Item", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)

            row1 = ttk.Frame(frm)
            row1.pack(fill="x", pady=6)
            ttk.Label(row1, text="Serial/Barcode").grid(row=0,
                                                        column=0, sticky="w")
            code_var = tk.StringVar()
            ttk.Entry(row1, textvariable=code_var).grid(
                row=0, column=1, sticky="we", padx=6)
            row1.columnconfigure(1, weight=1)

            idx = self._get_selected_inventory_index()
            if idx is not None:
                r = self.inventory_df.loc[idx]
                code_var.set(r.get("Barcode") or r.get("Serial Number"))

            staff_choices = self._get_staff_choices()
            staff_var = tk.StringVar()
            dept_var = tk.StringVar()
            pos_var = tk.StringVar()
            row2 = ttk.Frame(frm)
            row2.pack(fill="x", pady=6)
            ttk.Label(row2, text="Staff").grid(row=0, column=0, sticky="w")
            cb = ttk.Combobox(row2, textvariable=staff_var,
                              values=staff_choices, state="readonly")
            cb.grid(row=0, column=1, sticky="we", padx=6)
            ttk.Label(row2, text="Department").grid(
                row=1, column=0, sticky="w")
            ttk.Entry(row2, textvariable=dept_var).grid(
                row=1, column=1, sticky="we", padx=6)
            ttk.Label(row2, text="Position").grid(row=2, column=0, sticky="w")
            ttk.Entry(row2, textvariable=pos_var).grid(
                row=2, column=1, sticky="we", padx=6)
            row2.columnconfigure(1, weight=1)

            def on_staff(_e=None):
                s = staff_var.get()
                local = self.staff_df[self.staff_df["Username"] == s]
                if not local.empty:
                    dept_var.set(str(local.iloc[0].get("Department", "")))
                    pos_var.set(str(local.iloc[0].get("Position", "")))
            cb.bind("<<ComboboxSelected>>", on_staff)

            def do_issue():
                code = code_var.get().strip()
                ix = self._find_item_index(serial=code, barcode=code)
                if ix is None:
                    messagebox.showerror(
                        "Not found", f"No inventory item with code '{code}'.")
                    return
                if not self.is_clerk():
                    messagebox.showerror(
                        "Permission", "Insufficient permissions.")
                    return
                r = self.inventory_df.loc[ix]
                available = str(r.get("Available", "Yes")).strip(
                ).lower() in ("yes", "true", "1", "y", "t")
                if not available:
                    messagebox.showerror("Stock", "Item already issued.")
                    return
                staff = staff_var.get().strip()
                if not staff:
                    messagebox.showerror("Missing", "Select staff.")
                    return
                # mark unavailable
                self.inventory_df.at[ix, "Available"] = "No"
                # create issued record
                self.issued_df = pd.concat([self.issued_df, pd.DataFrame([{
                    "Product Name": r.get("Product Name", ""),
                    "Category": r.get("Category", ""),
                    "Serial Number": r.get("Serial Number", ""),
                    "Barcode": r.get("Barcode", ""),
                    "Issued To": staff,
                    "Department": dept_var.get(),
                    "Position": pos_var.get(),
                    "Issue Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }])], ignore_index=True)
                save_sheet(self.inventory_df, INVENTORY_SHEET)
                save_sheet(self.issued_df, ISSUED_SHEET)
                append_history(self.current_user, "ISSUE",
                               f"{r.get('Serial Number', '') or r.get('Barcode', '')} -> {staff}")
                self._refresh_inventory_table()
                self.show_inventory()
                messagebox.showinfo(
                    "Issued", f"Issued {r.get('Serial Number', '') or r.get('Barcode', '')} to {staff}")
            ttk.Button(frm, text="Issue", command=do_issue, state=(
                "normal" if self.is_clerk() else "disabled")).pack(pady=8)
            return frm
        self._open_inline(build)

    def batch_issue(self):
        def build(parent):
            win = ttk.Frame(parent, padding=8)
            header = ttk.Frame(win)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_inventory).pack(side="left")
            ttk.Label(header, text="Batch Issue", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            ttk.Label(win, text="Enter codes (one per line)").pack(
                anchor="w", pady=(6, 0))
            txt = tk.Text(win, height=12)
            txt.pack(fill="both", expand=True, pady=6)
            staff_choices = self._get_staff_choices()
            staff_var = tk.StringVar(
                value=staff_choices[0] if staff_choices else "")
            dept_var = tk.StringVar()
            pos_var = tk.StringVar()
            frm = ttk.Frame(win)
            frm.pack(fill="x", pady=6)
            ttk.Label(frm, text="Staff").grid(row=0, column=0, sticky="w")
            ttk.Combobox(frm, textvariable=staff_var, values=staff_choices,
                         state="readonly").grid(row=0, column=1, sticky="we", padx=6)
            ttk.Label(frm, text="Dept").grid(row=1, column=0, sticky="w")
            ttk.Entry(frm, textvariable=dept_var).grid(
                row=1, column=1, sticky="we", padx=6)
            ttk.Label(frm, text="Position").grid(row=2, column=0, sticky="w")
            ttk.Entry(frm, textvariable=pos_var).grid(
                row=2, column=1, sticky="we", padx=6)
            frm.columnconfigure(1, weight=1)

            def process():
                if not staff_var.get().strip():
                    messagebox.showerror("Missing", "Select staff.")
                    return
                codes = [c.strip() for c in txt.get(
                    "1.0", "end").splitlines() if c.strip()]
                count = 0
                for code in codes:
                    ix = self._find_item_index(serial=code, barcode=code)
                    if ix is None:
                        continue
                    r = self.inventory_df.loc[ix]
                    if str(r.get("Available", "Yes")).strip().lower() != "yes":
                        continue
                    self.inventory_df.at[ix, "Available"] = "No"
                    self.issued_df = pd.concat([self.issued_df, pd.DataFrame([{
                        "Product Name": r.get("Product Name", ""),
                        "Category": r.get("Category", ""),
                        "Serial Number": r.get("Serial Number", ""),
                        "Barcode": r.get("Barcode", ""),
                        "Issued To": staff_var.get(),
                        "Department": dept_var.get(),
                        "Position": pos_var.get(),
                        "Issue Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }])], ignore_index=True)
                    count += 1
                if count:
                    save_sheet(self.inventory_df, INVENTORY_SHEET)
                    save_sheet(self.issued_df, ISSUED_SHEET)
                    append_history(self.current_user, "ISSUE_BATCH",
                                   f"{count} items -> {staff_var.get()}")
                self._refresh_inventory_table()
                self.show_inventory()
                messagebox.showinfo("Batch Issue", f"Issued {count} items.")
            ttk.Button(win, text="Done", command=process).pack(pady=8)
            return win
        self._open_inline(build)

    def show_issued(self):
        def build(parent):
            wrap = ttk.Frame(parent)
            cols = ISSUED_COLUMNS
            self.iss_tree = ttk.Treeview(wrap, columns=cols, show="headings")
            for c in cols:
                self.iss_tree.heading(c, text=c)
                self.iss_tree.column(c, width=140 if c !=
                                     "Product Name" else 180)
            self.iss_tree.pack(fill="both", expand=True)
            self.iss_row_map.clear()
            self.issued_df = load_sheet(ISSUED_SHEET, ISSUED_COLUMNS)
            for idx, r in self.issued_df.iterrows():
                iid = self.iss_tree.insert(
                    "", "end", values=[r.get(c, "") for c in cols])
                self.iss_row_map[iid] = idx
            btn = ttk.Frame(wrap)
            btn.pack(fill="x", pady=8)
            ttk.Button(btn, text="Return Item", command=self.return_item, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=6)
            ttk.Button(btn, text="Batch Return", command=self.batch_return, state=(
                "normal" if self.is_clerk() else "disabled")).pack(side="left", padx=6)
            self.status_var.set(
                f"Issued: {len(self.issued_df)} | Data: {DATA_FILE} | Auth: {AUTH_FILE}")
            return wrap
        self._open_inline(build)

    def _get_selected_issued_index(self):
        sel = getattr(self, "iss_tree", None)
        if sel is None:
            return None
        s = sel.selection()
        if not s:
            return None
        return self.iss_row_map.get(s[0])

    def return_item(self):
        def build(parent):
            frm = ttk.Frame(parent, padding=8)
            header = ttk.Frame(frm)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_issued).pack(side="left")
            ttk.Label(header, text="Return Item", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            row1 = ttk.Frame(frm)
            row1.pack(fill="x", pady=6)
            ttk.Label(row1, text="Serial/Barcode").grid(row=0,
                                                        column=0, sticky="w")
            code_var = tk.StringVar()
            ttk.Entry(row1, textvariable=code_var).grid(
                row=0, column=1, sticky="we", padx=6)
            row1.columnconfigure(1, weight=1)

            def do_return():
                code = code_var.get().strip()
                # find issued index by code
                ix = None
                for i, r in self.issued_df.iterrows():
                    if str(r.get("Serial Number", "")) == code or str(r.get("Barcode", "")) == code:
                        ix = i
                        break
                if ix is None:
                    messagebox.showerror(
                        "Not found", f"No issued record with code '{code}'.")
                    return
                row = self.issued_df.loc[ix]
                sn, bc = row.get("Serial Number", ""), row.get("Barcode", "")
                inv_ix = self._find_item_index(sn, bc)
                if inv_ix is None:
                    # create new inventory row on return
                    new_inv = {"Product Name": row.get("Product Name", ""), "Category": row.get("Category", ""),
                               "Serial Number": sn, "Barcode": bc, "Available": "Yes", "Location": "", "Quantity": 1, "Unit Price": 0.0, "Low Stock Threshold": 0}
                    self.inventory_df = pd.concat(
                        [self.inventory_df, pd.DataFrame([new_inv])], ignore_index=True)
                else:
                    self.inventory_df.at[inv_ix, "Available"] = "Yes"
                # remove issued record
                self.issued_df = self.issued_df.drop(
                    index=ix).reset_index(drop=True)
                save_sheet(self.inventory_df, INVENTORY_SHEET)
                save_sheet(self.issued_df, ISSUED_SHEET)
                append_history(self.current_user, "RETURN", f"{sn or bc}")
                self.show_issued()
                messagebox.showinfo("Returned", f"Returned {sn or bc}")
            ttk.Button(frm, text="Return", command=do_return, state=(
                "normal" if self.is_clerk() else "disabled")).pack(pady=8)
            return frm
        self._open_inline(build)

    def batch_return(self):
        def build(parent):
            win = ttk.Frame(parent, padding=8)
            header = ttk.Frame(win)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_issued).pack(side="left")
            ttk.Label(header, text="Batch Return", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            ttk.Label(win, text="Enter codes (one per line)").pack(anchor="w")
            txt = tk.Text(win, height=10)
            txt.pack(fill="both", expand=True, pady=6)

            def process():
                codes = [c.strip() for c in txt.get(
                    "1.0", "end").splitlines() if c.strip()]
                count = 0
                for code in codes:
                    ix = None
                    for i, r in self.issued_df.iterrows():
                        if str(r.get("Serial Number", "")) == code or str(r.get("Barcode", "")) == code:
                            ix = i
                            break
                    if ix is None:
                        continue
                    row = self.issued_df.loc[ix]
                    sn, bc = row.get("Serial Number", ""), row.get(
                        "Barcode", "")
                    inv_ix = self._find_item_index(sn, bc)
                    if inv_ix is None:
                        new_inv = {"Product Name": row.get("Product Name", ""), "Category": row.get("Category", ""),
                                   "Serial Number": sn, "Barcode": bc, "Available": "Yes", "Location": "", "Quantity": 1, "Unit Price": 0.0, "Low Stock Threshold": 0}
                        self.inventory_df = pd.concat(
                            [self.inventory_df, pd.DataFrame([new_inv])], ignore_index=True)
                    else:
                        self.inventory_df.at[inv_ix, "Available"] = "Yes"
                    self.issued_df = self.issued_df.drop(
                        index=ix).reset_index(drop=True)
                    count += 1
                if count:
                    save_sheet(self.inventory_df, INVENTORY_SHEET)
                    save_sheet(self.issued_df, ISSUED_SHEET)
                    append_history(self.current_user,
                                   "RETURN_BATCH", f"{count} items")
                self.show_issued()
                messagebox.showinfo("Batch Return", f"Returned {count} items.")
            ttk.Button(win, text="Done", command=process).pack(pady=8)
            return win
        self._open_inline(build)

    def _get_staff_choices(self):
        self.staff_df = load_sheet(STAFF_SHEET, STAFF_COLUMNS)
        return self.staff_df["Username"].dropna().astype(str).tolist()

    # ---- CSV import/export ----
    def import_csv(self):
        file = filedialog.askopenfilename(
            title="Select CSV", filetypes=[("CSV Files", "*.csv")])
        if not file:
            return
        try:
            df = pd.read_csv(file)
            for col in INVENTORY_COLUMNS:
                if col not in df.columns:
                    if col == "Quantity":
                        df[col] = 0
                    elif col == "Unit Price":
                        df[col] = 0.0
                    elif col == "Available":
                        df[col] = "Yes"
                    else:
                        df[col] = ""
            df["Quantity"] = pd.to_numeric(
                df["Quantity"], errors="coerce").fillna(0).astype(int)
            df["Unit Price"] = pd.to_numeric(
                df["Unit Price"], errors="coerce").fillna(0.0).astype(float)
            inv = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
            for _, row in df.iterrows():
                ix = self._find_item_index(
                    row.get("Serial Number", ""), row.get("Barcode", ""))
                if ix is not None:
                    inv.at[ix, "Quantity"] = int(
                        inv.at[ix, "Quantity"]) + int(row["Quantity"])
                    for c in ["Product Name", "Category", "Location", "Unit Price", "Low Stock Threshold", "Serial Number", "Barcode", "Available"]:
                        val = row.get(c, "")
                        if pd.notna(val) and val != "":
                            inv.at[ix, c] = val
                else:
                    inv = pd.concat([inv, pd.DataFrame([row])],
                                    ignore_index=True)
            self.inventory_df = inv
            save_sheet(self.inventory_df, INVENTORY_SHEET)
            append_history(self.current_user, "IMPORT", os.path.basename(file))
            self._refresh_inventory_table()
            messagebox.showinfo("Import", f"Imported/merged {len(df)} rows.")
        except Exception as e:
            messagebox.showerror("Import error", str(e))

    def export_csv(self):
        file = filedialog.asksaveasfilename(
            title="Export Inventory", defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not file:
            return
        try:
            self.inventory_df.to_csv(file, index=False)
            messagebox.showinfo("Export", f"Exported to {file}")
        except Exception as e:
            messagebox.showerror("Export error", str(e))

    def export_unavailable(self):
        low = self.inventory_df[self.inventory_df["Available"].apply(
            lambda v: str(v).strip().lower()) != "yes"]
        if low.empty:
            messagebox.showinfo("Unavailable Items", "No unavailable items.")
            return
        file = filedialog.asksaveasfilename(
            title="Save Unavailable Items CSV", defaultextension=".csv", filetypes=[("CSV", ".csv")])
        if not file:
            return
        low.to_csv(file, index=False)
        messagebox.showinfo(
            "Saved", f"Unavailable items report saved to {file}")

    # ---- Categories management ----
    def manage_categories(self):
        def build(parent):
            win = ttk.Frame(parent, padding=8)
            header = ttk.Frame(win)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_inventory).pack(side="left")
            ttk.Label(header, text="Categories", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            lst = tk.Listbox(win)
            lst.pack(fill="both", expand=True, padx=8, pady=8)
            self.categories = sorted(load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)[
                                     "Category"].replace("", pd.NA).dropna().astype(str).unique().tolist())
            for c in self.categories:
                lst.insert("end", c)
            row = ttk.Frame(win)
            row.pack(fill="x", padx=8, pady=6)
            var = tk.StringVar()
            ttk.Entry(row, textvariable=var).pack(
                side="left", fill="x", expand=True)

            def add():
                v = var.get().strip()
                if v and v not in self.categories:
                    self.categories.append(v)
                    self.categories.sort()
                    lst.insert("end", v)
                    var.set("")
                    # no inventory rows changed on add (category list is just convenience)

            def rename():
                sel = lst.curselection()
                if not sel:
                    return
                old = lst.get(sel[0])
                new = simpledialog.askstring(
                    "Rename Category", f"Rename '{old}' to:", parent=self)
                if not new:
                    return
                new = new.strip()
                if not new:
                    return
                # update inventory rows where category == old
                inv = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
                inv.loc[inv["Category"] == old, "Category"] = new
                save_sheet(inv, INVENTORY_SHEET)
                append_history(self.current_user,
                               "CATEGORY_RENAME", f"{old} -> {new}")
                # refresh listbox
                self.categories = sorted(inv["Category"].replace(
                    "", pd.NA).dropna().astype(str).unique().tolist())
                lst.delete(0, "end")
                for c in self.categories:
                    lst.insert("end", c)

            def remove():
                sel = lst.curselection()
                if not sel:
                    return
                v = lst.get(sel[0])
                if not messagebox.askyesno("Confirm", f"Remove category '{v}'? This will clear the Category field from any items using it."):
                    return
                inv = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
                inv.loc[inv["Category"] == v, "Category"] = ""
                save_sheet(inv, INVENTORY_SHEET)
                append_history(self.current_user, "CATEGORY_REMOVE", v)
                self.categories = sorted(inv["Category"].replace(
                    "", pd.NA).dropna().astype(str).unique().tolist())
                lst.delete(0, "end")
                for c in self.categories:
                    lst.insert("end", c)
            ttk.Button(row, text="Add", command=add).pack(side="left", padx=4)
            ttk.Button(row, text="Rename", command=rename).pack(
                side="left", padx=4)
            ttk.Button(row, text="Remove", command=remove).pack(
                side="left", padx=4)
            return win
        self._open_inline(build)

    # ---- Staff ----
    def manage_staff(self):
        def build(parent):
            win = ttk.Frame(parent, padding=8)
            header = ttk.Frame(win)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_dashboard).pack(side="left")
            ttk.Label(header, text="Staff", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            cols = STAFF_COLUMNS
            tree = ttk.Treeview(win, columns=cols, show="headings")
            for c in cols:
                tree.heading(c, text=c)
                tree.column(c, width=180)
            tree.pack(fill="both", expand=True, padx=6, pady=6)

            def refresh():
                for r in tree.get_children():
                    tree.delete(r)
                self.staff_df = load_sheet(STAFF_SHEET, STAFF_COLUMNS)
                for _, row in self.staff_df.iterrows():
                    tree.insert("", "end", values=[row[c] for c in cols])
            frm = ttk.Frame(win)
            frm.pack(fill="x", padx=6, pady=6)
            u = tk.StringVar()
            d = tk.StringVar()
            p = tk.StringVar()
            ttk.Label(frm, text="Username").grid(row=0, column=0, sticky="w")
            ttk.Entry(frm, textvariable=u).grid(row=0, column=1, sticky="we")
            ttk.Label(frm, text="Department").grid(row=1, column=0, sticky="w")
            ttk.Entry(frm, textvariable=d).grid(row=1, column=1, sticky="we")
            ttk.Label(frm, text="Position").grid(row=2, column=0, sticky="w")
            ttk.Entry(frm, textvariable=p).grid(row=2, column=1, sticky="we")
            frm.columnconfigure(1, weight=1)

            def add_staff():
                if not u.get().strip():
                    messagebox.showerror("Error", "Username required")
                    return
                new_row = {"Username": u.get().strip(), "Department": d.get(
                ).strip(), "Position": p.get().strip()}
                self.staff_df = pd.concat(
                    [self.staff_df, pd.DataFrame([new_row])], ignore_index=True)
                save_sheet(self.staff_df, STAFF_SHEET)
                refresh()
            ttk.Button(win, text="Add/Save", command=add_staff).pack(pady=4)
            refresh()
            return win
        self._open_inline(build)

    # ---- Settings & Users ----
    def open_settings(self):
        if not self.is_admin():
            return

        def build(parent):
            win = ttk.Frame(parent, padding=8)
            header = ttk.Frame(win)
            header.pack(fill="x", pady=(0, 6))
            ttk.Button(header, text="← Back",
                       command=self.show_dashboard).pack(side="left")
            ttk.Label(header, text="Settings", font=(
                "Segoe UI", 12, "bold")).pack(side="left", padx=8)
            nb = ttk.Notebook(win)
            nb.pack(fill="both", expand=True)

            # General
            gen = ttk.Frame(nb)
            nb.add(gen, text="General")
            self.settings_df = load_sheet(SETTINGS_SHEET, SETTINGS_COLUMNS)
            settings = settings_to_dict(self.settings_df)
            cam_var = tk.StringVar(value=settings.get("camera_index", "0"))
            bkp_var = tk.BooleanVar(
                value=(settings.get("backup_enabled", "1") == "1"))
            ttk.Label(gen, text="Camera Index").grid(
                row=0, column=0, sticky="w", padx=8, pady=6)
            ttk.Entry(gen, textvariable=cam_var).grid(
                row=0, column=1, sticky="we", padx=8, pady=6)
            ttk.Checkbutton(gen, text="Backup before save", variable=bkp_var).grid(
                row=1, column=0, columnspan=2, sticky="w", padx=8, pady=6)
            gen.columnconfigure(1, weight=1)

            def save_general():
                self._set_setting("camera_index", cam_var.get().strip() or "0")
                self._set_setting("backup_enabled",
                                  "1" if bkp_var.get() else "0")
                messagebox.showinfo("Saved", "General settings saved.")
            ttk.Button(gen, text="Save", command=save_general).grid(
                row=99, column=0, columnspan=2, pady=8)

            # AD
            ad = ttk.Frame(nb)
            nb.add(ad, text="Active Directory")
            ad_server = tk.StringVar(value=settings.get("ad_server", ""))
            ad_base = tk.StringVar(value=settings.get("ad_base_dn", ""))
            ad_user = tk.StringVar(value=settings.get("ad_user", ""))
            ad_pass = tk.StringVar(value="")
            ttk.Label(ad, text="LDAP Server").grid(
                row=0, column=0, sticky="w", padx=8, pady=4)
            ttk.Entry(ad, textvariable=ad_server).grid(
                row=0, column=1, sticky="we", padx=8, pady=4)
            ttk.Label(ad, text="Base DN").grid(
                row=1, column=0, sticky="w", padx=8, pady=4)
            ttk.Entry(ad, textvariable=ad_base).grid(
                row=1, column=1, sticky="we", padx=8, pady=4)
            ttk.Label(ad, text="AD User").grid(
                row=2, column=0, sticky="w", padx=8, pady=4)
            ttk.Entry(ad, textvariable=ad_user).grid(
                row=2, column=1, sticky="we", padx=8, pady=4)
            ttk.Label(ad, text="Password").grid(
                row=3, column=0, sticky="w", padx=8, pady=4)
            ttk.Entry(ad, textvariable=ad_pass, show="*").grid(row=3,
                                                               column=1, sticky="we", padx=8, pady=4)
            ad.columnconfigure(1, weight=1)

            def save_ad():
                self._set_setting("ad_server", ad_server.get().strip())
                self._set_setting("ad_base_dn", ad_base.get().strip())
                self._set_setting("ad_user", ad_user.get().strip())
                # not storing password plaintext; if you want secure store, integrate encryption
                messagebox.showinfo("Saved", "AD settings saved.")
            row = ttk.Frame(ad)
            row.grid(row=99, column=0, columnspan=2, pady=8)
            ttk.Button(row, text="Save AD Settings",
                       command=save_ad).pack(side="left", padx=4)

            # Users & Roles
            usr = ttk.Frame(nb)
            nb.add(usr, text="Users & Roles")
            cols = ["Username", "Role", "Active"]
            tree = ttk.Treeview(usr, columns=cols, show="headings", height=8)
            for c in cols:
                tree.heading(c, text=c)
                tree.column(c, width=160)
            tree.pack(fill="both", expand=True, padx=8, pady=8)

            def refresh_users():
                for r in tree.get_children():
                    tree.delete(r)
                self.users_df = load_sheet(USERS_SHEET, USERS_COLUMNS)
                for _, row_ in self.users_df.iterrows():
                    tree.insert("", "end", values=[row_["Username"], row_[
                                "Role"], bool(row_["Active"])])
            refresh_users()
            addf = ttk.LabelFrame(usr, text="Add / Update User")
            addf.pack(fill="x", padx=8, pady=8)
            un = tk.StringVar()
            ro = tk.StringVar(value="Viewer")
            ac = tk.BooleanVar(value=True)
            pw = tk.StringVar()
            ttk.Label(addf, text="Username").grid(row=0, column=0, sticky="w")
            ttk.Entry(addf, textvariable=un).grid(
                row=0, column=1, sticky="we", padx=6)
            ttk.Label(addf, text="Role").grid(row=1, column=0, sticky="w")
            ttk.Combobox(addf, textvariable=ro, values=list(ROLES), state="readonly").grid(
                row=1, column=1, sticky="we", padx=6)
            ttk.Label(addf, text="Active").grid(row=2, column=0, sticky="w")
            ttk.Checkbutton(addf, variable=ac).grid(
                row=2, column=1, sticky="w", padx=6)
            ttk.Label(addf, text="Password (leave blank to keep)").grid(
                row=3, column=0, sticky="w")
            ttk.Entry(addf, textvariable=pw, show="*").grid(row=3,
                                                            column=1, sticky="we", padx=6)
            addf.columnconfigure(1, weight=1)

            def save_user():
                if not un.get().strip():
                    messagebox.showerror("User", "Username required")
                    return
                users = load_sheet(USERS_SHEET, USERS_COLUMNS)
                mask = users["Username"] == un.get().strip()
                if mask.any():
                    i = users.index[mask][0]
                    users.at[i, "Role"] = ro.get()
                    users.at[i, "Active"] = bool(ac.get())
                    if pw.get().strip():
                        users.at[i, "PasswordHash"] = hash_pw(pw.get().strip())
                else:
                    users = pd.concat([users, pd.DataFrame([{
                        "Username": un.get().strip(),
                        "PasswordHash": hash_pw(pw.get().strip() or "ChangeMe123!"),
                        "Role": ro.get(),
                        "Active": bool(ac.get())
                    }])], ignore_index=True)
                save_sheet(users, USERS_SHEET)
                ensure_admin_user()
                refresh_users()
                messagebox.showinfo("Users", "Saved.")
            ttk.Button(addf, text="Save User", command=save_user).grid(
                row=4, column=0, columnspan=2, pady=8)

            return win
        self._open_inline(build)

    def _set_setting(self, key, value):
        df = load_sheet(SETTINGS_SHEET, SETTINGS_COLUMNS)
        mask = df["Key"] == key
        if mask.any():
            df.loc[mask, "Value"] = str(value)
        else:
            df = pd.concat(
                [df, pd.DataFrame([{"Key": key, "Value": str(value)}])], ignore_index=True)
        save_sheet(df, SETTINGS_SHEET)
        self.settings_df = df

    # ---- Labels (simple) ----
    def save_label(self):
        idx = self._get_selected_inventory_index()
        if idx is None:
            messagebox.showinfo("Select", "Select an item to create label.")
            return
        row = self.inventory_df.loc[idx]
        code = row.get("Barcode") or row.get("Serial Number")
        if not code:
            messagebox.showerror("Missing", "No Serial/Barcode on this item.")
            return
        messagebox.showinfo(
            "Label", f"Label would be generated for code: {code}\n(Install Pillow + barcode libs to enable PNG labels)")

    # ---- logout ----
    def logout(self):
        append_history(self.current_user, "LOGOUT", "")
        self.current_user = None
        self.current_role = "Viewer"
        for w in self.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass
        # reload data and show login
        self.inventory_df = load_sheet(INVENTORY_SHEET, INVENTORY_COLUMNS)
        self.issued_df = load_sheet(ISSUED_SHEET, ISSUED_COLUMNS)
        self.staff_df = load_sheet(STAFF_SHEET, STAFF_COLUMNS)
        self.users_df = load_sheet(USERS_SHEET, USERS_COLUMNS)
        self.settings_df = load_sheet(SETTINGS_SHEET, SETTINGS_COLUMNS)
        ensure_admin_user()
        self._build_login()


# ---- run ----
if __name__ == "__main__":
    app = InventoryApp()
    app.mainloop()
