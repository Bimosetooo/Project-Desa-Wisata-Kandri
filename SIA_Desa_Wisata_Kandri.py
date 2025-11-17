import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import sqlite3
from decimal import Decimal, InvalidOperation

DB_FILE = "akuntansi.db"

# --------------------------
# Initial COA (user list)
# --------------------------
INITIAL_ACCOUNTS = [
    ("101", "Kas"),
    ("111", "Persediaan Caping"),
    ("112", "Persediaan Cat"),
    ("113", "Persediaan Kuas"),
    ("201", "Utang Freelance"),
    ("202", "Utang Operasional"),
    ("301", "Ekuitas Pokdarwis"),
    ("401", "Pendapatan Paket Wisata"),
    ("402", "Pendapatan Melukis Caping"),
    ("403", "Pendapatan Eksplorasi Singkong"),
    ("404", "Pendapatan Edukasi Pertanian"),
    ("405", "Pendapatan Konsumsi / Makan"),
    ("406", "Pendapatan Aktivitas Ikan & Sendang"),
    ("407", "Pendapatan Pemandu"),
    ("501", "Beban Melukis Caping"),
    ("502", "Beban Snack / Break"),
    ("503", "Beban Sendang"),
    ("504", "Beban Pakan Kambing / Sapi"),
    ("505", "Beban Kolam Ikan"),
    ("506", "Beban Menanam Padi"),
    ("507", "Beban Tangkap Ikan"),
    ("508", "Beban Makan Siang"),
    ("509", "Beban Pemandu Wisata"),
    ("510", "Beban Pokdarwis"),
    ("511", "Beban Marketing"),
    ("512", "Beban Juru Kunci"),
    ("520", "Beban Tenaga Kerja Freelance"),
    ("530", "Beban Listrik, Internet, kebersihan"),
]


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SIA Desa Wisata Kandri")
        self.root.geometry("980x620")

        # DB
        self.conn = sqlite3.connect(DB_FILE)
        self.cur = self.conn.cursor()
        self._setup_db()
        self.accounts = self.load_accounts()

        # Styles
        self.style = ttk.Style(root)
        try:
            self.style.theme_use("clam")
        except:
            pass

        # ===== UI Layout =====
        self.setup_top_frame()
        self.setup_notebook()
        self.refresh_all()

    # ----------------- Database Setup -----------------
    def _setup_db(self):
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT UNIQUE,
                name TEXT
            )
        """)
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                deskripsi TEXT,
                account_code TEXT,
                account_name TEXT,
                debit REAL,
                kredit REAL
            )
        """)
        self.conn.commit()
        # seed accounts if empty
        self.cur.execute("SELECT COUNT(*) FROM accounts")
        if self.cur.fetchone()[0] == 0:
            self.cur.executemany("INSERT INTO accounts (code, name) VALUES (?, ?)", INITIAL_ACCOUNTS)
            self.conn.commit()

    # ----------------- Load / Account CRUD -----------------
    def load_accounts(self):
        self.cur.execute("SELECT code, name FROM accounts ORDER BY code")
        return self.cur.fetchall()

    def add_account_db(self, code, name):
        try:
            self.cur.execute("INSERT INTO accounts (code, name) VALUES (?,?)", (code, name))
            self.conn.commit()
            self.accounts = self.load_accounts()
            return True, "Akun ditambahkan"
        except sqlite3.IntegrityError:
            return False, "Kode akun sudah ada"

    def delete_account_db(self, code):
        # simple delete (no cascade check)
        self.cur.execute("DELETE FROM accounts WHERE code=?", (code,))
        self.conn.commit()
        self.accounts = self.load_accounts()

    # ----------------- Top Input Frame -----------------
    def setup_top_frame(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(side=tk.TOP, fill=tk.X)

        ttk.Label(top, text="Tanggal (YYYY-MM-DD):").grid(row=0, column=0, sticky="w")
        self.entry_date = ttk.Entry(top, width=14)
        self.entry_date.grid(row=0, column=1, padx=6)
        self.entry_date.insert(0, datetime.today().strftime("%Y-%m-%d"))

        ttk.Label(top, text="Deskripsi:").grid(row=0, column=2, sticky="w")
        self.entry_desc = ttk.Entry(top, width=40)
        self.entry_desc.grid(row=0, column=3, padx=6)

        ttk.Label(top, text="Akun Debit:").grid(row=1, column=0, sticky="w", pady=6)
        self.combo_debit = ttk.Combobox(top, values=self._acct_display_list(), width=36, state="readonly")
        self.combo_debit.grid(row=1, column=1, padx=6)
        if self.combo_debit['values']:
            self.combo_debit.current(0)

        ttk.Label(top, text="Akun Kredit:").grid(row=1, column=2, sticky="w")
        self.combo_credit = ttk.Combobox(top, values=self._acct_display_list(), width=36, state="readonly")
        self.combo_credit.grid(row=1, column=3, padx=6)
        if self.combo_credit['values']:
            self.combo_credit.current(1 if len(self.combo_credit['values']) > 1 else 0)

        ttk.Label(top, text="Nominal:").grid(row=2, column=0, sticky="w", pady=6)
        self.entry_amount = ttk.Entry(top, width=20)
        self.entry_amount.grid(row=2, column=1, padx=6, sticky="w")

        btn_save = ttk.Button(top, text="Simpan Transaksi (Double Entry)", command=self.on_save_transaction)
        btn_save.grid(row=2, column=3, sticky="e", padx=6)

        # small controls: manage accounts
        ttk.Button(top, text="Daftar Akun", command=self.show_account_manager).grid(row=2, column=2, sticky="w", padx=6)

    def _acct_display_list(self):
        return [f"{code} - {name}" for code, name in self.accounts]

    # ----------------- Notebook (tabs) -----------------
    def setup_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # Tab 1: Jurnal Umum
        self.tab_journal = ttk.Frame(self.nb)
        self.nb.add(self.tab_journal, text="Jurnal Umum")
        self.tree_journal = ttk.Treeview(self.tab_journal, columns=("Tanggal", "Keterangan", "NamaAkun", "NoAkun", "Debit", "Kredit"), show="headings")
        for col, w in [("Tanggal",100), ("Keterangan",280), ("NamaAkun",200), ("NoAkun",80), ("Debit",100), ("Kredit",100)]:
            self.tree_journal.heading(col, text=col)
            self.tree_journal.column(col, width=w, anchor="w")
        self.tree_journal.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll_j = ttk.Scrollbar(self.tab_journal, orient=tk.VERTICAL, command=self.tree_journal.yview)
        self.tree_journal.configure(yscroll=scroll_j.set)
        scroll_j.pack(side=tk.RIGHT, fill=tk.Y)

        # Tab 2: Buku Besar (select account)
        self.tab_ledger = ttk.Frame(self.nb)
        self.nb.add(self.tab_ledger, text="Buku Besar (T-account)")
        top2 = ttk.Frame(self.tab_ledger, padding=6)
        top2.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(top2, text="Pilih Akun:").pack(side=tk.LEFT)
        self.ledger_combo = ttk.Combobox(top2, values=self._acct_display_list(), state="readonly", width=40)
        self.ledger_combo.pack(side=tk.LEFT, padx=6)
        ttk.Button(top2, text="Tampilkan", command=self.show_ledger_for_account).pack(side=tk.LEFT, padx=6)
        self.tree_ledger = ttk.Treeview(self.tab_ledger, columns=("Tanggal", "Deskripsi", "Debit", "Kredit", "Saldo"), show="headings")
        for col, w in [("Tanggal",120), ("Deskripsi",260), ("Debit",100), ("Kredit",100), ("Saldo",120)]:
            self.tree_ledger.heading(col, text=col)
            self.tree_ledger.column(col, width=w, anchor="w")
        self.tree_ledger.pack(fill=tk.BOTH, expand=True)

        # Tab 3: Neraca Saldo
        self.tab_tb = ttk.Frame(self.nb)
        self.nb.add(self.tab_tb, text="Neraca Saldo")
        self.tree_tb = ttk.Treeview(self.tab_tb, columns=("NoAkun", "NamaAkun", "Debit", "Kredit"), show="headings")
        for col, w in [("NoAkun",100), ("NamaAkun",400), ("Debit",140), ("Kredit",140)]:
            self.tree_tb.heading(col, text=col)
            self.tree_tb.column(col, width=w, anchor="w")
        self.tree_tb.pack(fill=tk.BOTH, expand=True)
        # total label
        self.tb_total_label = ttk.Label(self.tab_tb, text="")
        self.tb_total_label.pack(pady=6)

    # ----------------- Transactions -----------------
    def on_save_transaction(self):
        tanggal = self.entry_date.get().strip()
        deskripsi = self.entry_desc.get().strip()
        debit_sel = self.combo_debit.get()
        credit_sel = self.combo_credit.get()
        amount_s = self.entry_amount.get().strip()
        # validations
        try:
            datetime.strptime(tanggal, "%Y-%m-%d")
        except:
            messagebox.showerror("Error", "Format tanggal harus YYYY-MM-DD")
            return
        if not deskripsi:
            messagebox.showerror("Error", "Isi deskripsi")
            return
        if not debit_sel or not credit_sel:
            messagebox.showerror("Error", "Pilih akun debit & kredit")
            return
        debit_code = debit_sel.split(" - ")[0].strip()
        credit_code = credit_sel.split(" - ")[0].strip()
        if debit_code == credit_code:
            messagebox.showerror("Error", "Akun debit dan kredit tidak boleh sama")
            return
        try:
            amt = Decimal(amount_s.replace(",", ""))
            if amt <= 0:
                raise InvalidOperation
        except Exception:
            messagebox.showerror("Error", "Nominal harus angka > 0")
            return

        # Double entry: two rows
        self.cur.execute("INSERT INTO journal (tanggal, deskripsi, account_code, account_name, debit, kredit) VALUES (?, ?, ?, ?, ?, ?)",
                         (tanggal, deskripsi, debit_code, self._get_account_name(debit_code), float(amt), 0.0))
        self.cur.execute("INSERT INTO journal (tanggal, deskripsi, account_code, account_name, debit, kredit) VALUES (?, ?, ?, ?, ?, ?)",
                         (tanggal, deskripsi, credit_code, self._get_account_name(credit_code), 0.0, float(amt)))
        self.conn.commit()
        messagebox.showinfo("Sukses", "Transaksi tersimpan (Jurnal Umum).")
        # clear nominal & desc
        self.entry_amount.delete(0, tk.END)
        self.entry_desc.delete(0, tk.END)
        # refresh displays
        self.refresh_all()

    def _get_account_name(self, code):
        self.cur.execute("SELECT name FROM accounts WHERE code=?", (code,))
        r = self.cur.fetchone()
        return r[0] if r else ""

    # ----------------- Displays / Reports -----------------
    def refresh_journal(self):
        for r in self.tree_journal.get_children():
            self.tree_journal.delete(r)
        self.cur.execute("SELECT tanggal, deskripsi, account_name, account_code, debit, kredit FROM journal ORDER BY tanggal, id")
        for row in self.cur.fetchall():
            # display like: date | description | account_name | account_code | debit | credit
            self.tree_journal.insert("", tk.END, values=(row[0], row[1], row[2], row[3], f"{row[4]:,.2f}" if row[4] else "", f"{row[5]:,.2f}" if row[5] else ""))

    def show_ledger_for_account(self):
        sel = self.ledger_combo.get()
        if not sel:
            messagebox.showerror("Error", "Pilih akun terlebih dahulu")
            return
        code = sel.split(" - ")[0]
        # clear
        for r in self.tree_ledger.get_children():
            self.tree_ledger.delete(r)
        # fetch
        self.cur.execute("SELECT tanggal, deskripsi, debit, kredit FROM journal WHERE account_code=? ORDER BY tanggal, id", (code,))
        saldo = Decimal("0.00")
        for tgl, desc, deb, kre in self.cur.fetchall():
            deb = Decimal(str(deb))
            kre = Decimal(str(kre))
            saldo += deb - kre
            self.tree_ledger.insert("", tk.END, values=(tgl, desc, f"{deb:,.2f}" if deb else "", f"{kre:,.2f}" if kre else "", f"{saldo:,.2f}"))

    def refresh_trial_balance(self):
        for r in self.tree_tb.get_children():
            self.tree_tb.delete(r)
        self.cur.execute("SELECT code, name FROM accounts ORDER BY code")
        total_d = Decimal("0.00")
        total_k = Decimal("0.00")
        for code, name in self.cur.fetchall():
            self.cur.execute("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(kredit),0) FROM journal WHERE account_code=?", (code,))
            d, k = self.cur.fetchone()
            d = Decimal(str(d))
            k = Decimal(str(k))
            total_d += d
            total_k += k
            self.tree_tb.insert("", tk.END, values=(code, name, f"{d:,.2f}" if d else "", f"{k:,.2f}" if k else ""))
        self.tb_total_label.config(text=f"TOTAL DEBIT: {total_d:,.2f}    |    TOTAL KREDIT: {total_k:,.2f}")

    def refresh_accounts_in_ui(self):
        self.accounts = self.load_accounts()
        # update combobox lists
        vals = self._acct_display_list()
        self.combo_debit['values'] = vals
        self.combo_credit['values'] = vals
        self.ledger_combo['values'] = vals
        # set default selections if empty
        if vals:
            self.combo_debit.current(0)
            self.combo_credit.current(1 if len(vals) > 1 else 0)

    def refresh_all(self):
        self.refresh_accounts_in_ui()
        self.refresh_journal()
        self.refresh_trial_balance()

    # ----------------- Account Manager Window -----------------
    def show_account_manager(self):
        win = tk.Toplevel(self.root)
        win.title("Daftar Akun & Manajemen")
        win.geometry("580x420")
        # tree
        cols = ("Kode", "Nama")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for c, w in [("Kode",120), ("Nama",420)]:
            tree.heading(c, text=c); tree.column(c, width=w)
        tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        # insert
        self.cur.execute("SELECT code, name FROM accounts ORDER BY code")
        for code, name in self.cur.fetchall():
            tree.insert("", tk.END, values=(code, name))
        # add frame
        f = ttk.Frame(win, padding=6)
        f.pack(fill=tk.X)
        ttk.Label(f, text="Kode").grid(row=0, column=0); e_code = ttk.Entry(f, width=12); e_code.grid(row=0, column=1, padx=4)
        ttk.Label(f, text="Nama").grid(row=0, column=2); e_name = ttk.Entry(f, width=40); e_name.grid(row=0, column=3, padx=4)
        def add_acc():
            code = e_code.get().strip(); name = e_name.get().strip()
            if not code or not name:
                messagebox.showerror("Error", "Isi kode & nama akun")
                return
            ok, msg = self.add_account_db(code, name)
            if not ok:
                messagebox.showerror("Error", msg)
            else:
                tree.insert("", tk.END, values=(code, name))
                self.refresh_all()
                e_code.delete(0, tk.END); e_name.delete(0, tk.END)
        ttk.Button(f, text="Tambah", command=add_acc).grid(row=1, column=1, pady=6)
        def del_acc():
            sel = tree.selection()
            if not sel:
                return
            code = tree.item(sel[0])['values'][0]
            if messagebox.askyesno("Confirm", f"Hapus akun {code}?"):
                self.delete_account_db(code)
                tree.delete(sel[0])
                self.refresh_all()
        ttk.Button(f, text="Hapus Terpilih", command=del_acc).grid(row=1, column=3, pady=6)

# ----------------- Run -----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
