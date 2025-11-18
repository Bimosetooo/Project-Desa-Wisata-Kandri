# SIA_Desa_Wisata_Kandri.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import sqlite3
from decimal import Decimal
import os
import csv

# Optional libs (Excel/PDF). We'll handle missing libs gracefully.
try:
    from openpyxl import Workbook
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.pdfgen import canvas
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False

DB_FILE = "akuntansi.db"

INITIAL_ACCOUNTS = [
    ("101", "Kas"),
    ("102", "Bank"),
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
    ("530", "Beban Listrik, internet, kebersihan"),
]

def to_decimal(x):
    try:
        return Decimal(str(x))
    except Exception:
        return Decimal("0.00")

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("SIA Desa Wisata Kandri")
        self.root.geometry("1000x700")

        # DB
        self.conn = sqlite3.connect(DB_FILE)
        self.cur = self.conn.cursor()
        self._setup_db()
        self.load_accounts_cache()

        # inventory method default
        self.inventory_method = tk.StringVar(value="AVG")  # "FIFO" or "AVG"

        # UI
        self.setup_menu()
        self.setup_top_input()
        self.setup_notebook()
        self.refresh_all()

    # ---------------- DB setup ----------------
    def _setup_db(self):
    # Accounts
        self.cur.execute("""
        CREATE TABLE IF NOT EXISTS accounts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT
        )
    """)
    
    # Users (login)
        self.cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE,
            password TEXT
        )
    """)
    # Seed default user jika kosong
        self.cur.execute("SELECT COUNT(*) FROM users")
        if self.cur.fetchone()[0] == 0:
            self.cur.execute("INSERT INTO users (email, password) VALUES (?, ?)", ("admin@example.com", "admin123"))

    # main journal
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
    # ... semua table lainnya tetap seperti sebelumnya ...
    
        self.conn.commit()
    
    # seed accounts jika kosong
        self.cur.execute("SELECT COUNT(*) FROM accounts")
        if self.cur.fetchone()[0] == 0:
            self.cur.executemany("INSERT INTO accounts (code, name) VALUES (?, ?)", INITIAL_ACCOUNTS)
            self.conn.commit()


    def _setup_db(self):
        # accounts
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS accounts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT UNIQUE,
                name TEXT
            )
        """)
        # main journal: each row = single side of double entry
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
        # adjustment journal (separate)
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS adjustment_journal (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                deskripsi TEXT,
                account_code TEXT,
                account_name TEXT,
                debit REAL,
                kredit REAL
            )
        """)
        # purchases (IN)
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                doc_no TEXT,
                account_code TEXT,
                account_name TEXT,
                qty REAL,
                unit_price REAL,
                total REAL
            )
        """)
        # usage_out (OUT)
        self.cur.execute("""
            CREATE TABLE IF NOT EXISTS usage_out (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal TEXT,
                doc_no TEXT,
                account_code TEXT,
                account_name TEXT,
                qty REAL,
                cost_per_unit REAL,
                total_cost REAL,
                expense_account_code TEXT,
                expense_account_name TEXT,
                keterangan TEXT
            )
        """)
        self.conn.commit()

        # seed accounts if empty
        self.cur.execute("SELECT COUNT(*) FROM accounts")
        if self.cur.fetchone()[0] == 0:
            self.cur.executemany("INSERT INTO accounts (code, name) VALUES (?, ?)", INITIAL_ACCOUNTS)
            self.conn.commit()

    def load_accounts_cache(self):
        self.cur.execute("SELECT code, name FROM accounts ORDER BY code")
        self.accounts = self.cur.fetchall()

    # ---------------- Menu & UI ----------------
    def show_login(self):
        win = tk.Toplevel(self.root)
        win.title("Login")
        win.geometry("400x220")
        win.grab_set()

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Email:").grid(row=0, column=0, sticky="w", pady=4)
        e_email = ttk.Entry(frm, width=40)
        e_email.grid(row=0, column=1, pady=4)

        ttk.Label(frm, text="Password:").grid(row=1, column=0, sticky="w", pady=4)
        e_pass = ttk.Entry(frm, width=40, show="*")
        e_pass.grid(row=1, column=1, pady=4)

        def do_login():
            email = e_email.get().strip()
            pw = e_pass.get().strip()
            self.cur.execute("SELECT id FROM users WHERE email=? AND password=?", (email, pw))
            if self.cur.fetchone():
                messagebox.showinfo("Sukses", "Login berhasil")
                win.destroy()
            else:
                messagebox.showerror("Error", "Email atau password salah")

        ttk.Button(frm, text="Login", command=do_login).grid(row=2, column=1, pady=12, sticky="e")


    def setup_menu(self):
        menubar = tk.Menu(self.root)

        master_menu = tk.Menu(menubar, tearoff=0)
        master_menu.add_command(label="Daftar Akun / Manajemen", command=self.show_account_manager)
        master_menu.add_command(label="Daftar Persediaan (Stock Master)", command=self.show_stock_master)
        menubar.add_cascade(label="Master", menu=master_menu)

        trans_menu = tk.Menu(menubar, tearoff=0)
        trans_menu.add_command(label="Input Transaksi (Jurnal Umum)", command=self.open_journal_input)
        trans_menu.add_command(label="Pembelian Persediaan", command=self.open_purchase_input)
        trans_menu.add_command(label="Pemakaian Persediaan", command=self.open_usage_input)
        trans_menu.add_command(label="Jurnal Penyesuaian", command=self.open_adjustment_input)
        menubar.add_cascade(label="Transaksi", menu=trans_menu)

        laporan_menu = tk.Menu(menubar, tearoff=0)
        laporan_menu.add_command(label="Laporan Keuangan (Menu C)", command=self.open_reports_window)
        laporan_menu.add_command(label="Kartu Persediaan", command=self.open_inventory_card)
        menubar.add_cascade(label="Laporan", menu=laporan_menu)

        self.root.config(menu=menubar)

    def setup_top_input(self):
        frame = ttk.Frame(self.root, padding=6)
        frame.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(frame, text="Quick Input:").pack(side=tk.LEFT, padx=6)
        ttk.Button(frame, text="Input Jurnal", command=self.open_journal_input).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame, text="Pembelian", command=self.open_purchase_input).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame, text="Pemakaian", command=self.open_usage_input).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame, text="Jurnal Penyesuaian", command=self.open_adjustment_input).pack(side=tk.LEFT, padx=4)
        ttk.Button(frame, text="Buku Besar (Ledger)", width=30, command=self.open_ledger_window).pack(pady=5)


    def setup_notebook(self):
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Journal quick view
        tab1 = ttk.Frame(self.nb)
        self.nb.add(tab1, text="Jurnal Umum (Quick View)")
        self.tree_journal = ttk.Treeview(tab1, columns=("Tanggal","Deskripsi","Akun","Kode","Debit","Kredit"), show="headings")
        for col,w in [("Tanggal",100),("Deskripsi",300),("Akun",220),("Kode",80),("Debit",100),("Kredit",100)]:
            self.tree_journal.heading(col, text=col); self.tree_journal.column(col, width=w, anchor="w")
        self.tree_journal.pack(fill=tk.BOTH, expand=True)

        # Trial balance quick
        tab2 = ttk.Frame(self.nb)
        self.nb.add(tab2, text="Neraca Saldo (Quick)")
        self.tree_tb = ttk.Treeview(tab2, columns=("Kode","Akun","Debit","Kredit"), show="headings")
        for col,w in [("Kode",100),("Akun",420),("Debit",160),("Kredit",160)]:
            self.tree_tb.heading(col, text=col); self.tree_tb.column(col, width=w, anchor="w")
        self.tree_tb.pack(fill=tk.BOTH, expand=True)
        self.tb_label = ttk.Label(tab2, text="")
        self.tb_label.pack(pady=6)

    # ---------------- Account manager & stock master ----------------
    def show_account_manager(self):
        win = tk.Toplevel(self.root); win.title("Daftar Akun & Manajemen"); win.geometry("700x460")
        tree = ttk.Treeview(win, columns=("Kode","Nama"), show="headings")
        tree.heading("Kode", text="Kode"); tree.column("Kode", width=120)
        tree.heading("Nama", text="Nama"); tree.column("Nama", width=520)
        tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.cur.execute("SELECT code,name FROM accounts ORDER BY code")
        for code,name in self.cur.fetchall():
            tree.insert("", tk.END, values=(code, name))

        f = ttk.Frame(win, padding=6); f.pack(fill=tk.X)
        ttk.Label(f, text="Kode").grid(row=0,column=0); e_code = ttk.Entry(f, width=10); e_code.grid(row=0,column=1,padx=4)
        ttk.Label(f, text="Nama").grid(row=0,column=2); e_name = ttk.Entry(f, width=50); e_name.grid(row=0,column=3,padx=4)

        def add_acc():
            code = e_code.get().strip(); name = e_name.get().strip()
            if not code or not name:
                messagebox.showerror("Error","Isi kode & nama"); return
            try:
                self.cur.execute("INSERT INTO accounts (code,name) VALUES (?,?)",(code,name))
                self.conn.commit()
                tree.insert("",tk.END, values=(code,name))
                e_code.delete(0,tk.END); e_name.delete(0,tk.END)
                self.load_accounts_cache(); self.refresh_all()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error","Kode akun sudah ada")
        ttk.Button(f, text="Tambah Akun", command=add_acc).grid(row=1,column=1,pady=6)

        def del_acc():
            sel = tree.selection()
            if not sel: return
            code = tree.item(sel[0])['values'][0]
            if messagebox.askyesno("Confirm", f"Hapus akun {code}? (Pastikan tidak ada transaksi terkait)"):
                self.cur.execute("DELETE FROM accounts WHERE code=?", (code,))
                self.conn.commit()
                tree.delete(sel[0])
                self.load_accounts_cache(); self.refresh_all()
        ttk.Button(f, text="Hapus Terpilih", command=del_acc).grid(row=1,column=3,pady=6)

    def show_stock_master(self):
        win = tk.Toplevel(self.root); win.title("Daftar Persediaan (Stock Master)"); win.geometry("600x400")
        tree = ttk.Treeview(win, columns=("Kode","Nama"), show="headings")
        tree.heading("Kode", text="Kode"); tree.column("Kode", width=120)
        tree.heading("Nama", text="Nama"); tree.column("Nama", width=440)
        tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.cur.execute("SELECT code,name FROM accounts WHERE code LIKE '11%' ORDER BY code")
        for code,name in self.cur.fetchall():
            tree.insert("", tk.END, values=(code,name))
        ttk.Label(win, text="Tambah akun persediaan (kode mulai 11xx)").pack()
        f = ttk.Frame(win, padding=6); f.pack()
        e_code = ttk.Entry(f, width=10); e_code.grid(row=0,column=0,padx=4)
        e_name = ttk.Entry(f, width=40); e_name.grid(row=0,column=1,padx=4)
        def add_stock():
            code = e_code.get().strip(); name = e_name.get().strip()
            if not code.startswith("11"):
                messagebox.showerror("Error","Kode persediaan harus mulai dengan '11'"); return
            try:
                self.cur.execute("INSERT INTO accounts (code,name) VALUES (?,?)",(code,name))
                self.conn.commit()
                tree.insert("", tk.END, values=(code,name))
                e_code.delete(0,tk.END); e_name.delete(0,tk.END)
                self.load_accounts_cache(); self.refresh_all()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error","Kode sudah ada")
        ttk.Button(f, text="Tambah Persediaan", command=add_stock).grid(row=1,column=0,pady=6)

    # ---------------- Journal input (double entry) ----------------
    def open_journal_input(self):
        win = tk.Toplevel(self.root)
        win.title("Input Jurnal Umum (Double Entry)")
        win.geometry("860x260")
        frm = ttk.Frame(win, padding=8); frm.pack(fill=tk.BOTH, expand=True)
        for i in range(4): frm.columnconfigure(i, weight=1)

        ttk.Label(frm, text="Tanggal (YYYY-MM-DD)").grid(row=0, column=0, sticky="w")
        e_date = ttk.Entry(frm); e_date.grid(row=0, column=1, padx=6, sticky="we")
        e_date.insert(0, datetime.today().strftime("%Y-%m-%d"))

        ttk.Label(frm, text="Deskripsi").grid(row=0, column=2, sticky="w")
        e_desc = ttk.Entry(frm); e_desc.grid(row=0, column=3, padx=6, sticky="we")

        self.load_accounts_cache()
        vals = [f"{c} - {n}" for c,n in self.accounts]

        ttk.Label(frm, text="Akun Debit").grid(row=1, column=0, sticky="w", pady=8)
        cb_debit = ttk.Combobox(frm, values=vals); cb_debit.grid(row=1, column=1, padx=6, sticky="we")
        ttk.Label(frm, text="Akun Kredit").grid(row=1, column=2, sticky="w")
        cb_credit = ttk.Combobox(frm, values=vals); cb_credit.grid(row=1, column=3, padx=6, sticky="we")

        ttk.Label(frm, text="Nominal").grid(row=2, column=0, sticky="w", pady=8)
        e_amount = ttk.Entry(frm); e_amount.grid(row=2, column=1, padx=6, sticky="we")

        def do_save():
            tanggal = e_date.get().strip(); desc = e_desc.get().strip()
            if not tanggal or not desc or not cb_debit.get() or not cb_credit.get() or not e_amount.get().strip():
                messagebox.showerror("Error","Isi semua field"); return
            try:
                datetime.strptime(tanggal, "%Y-%m-%d")
            except:
                messagebox.showerror("Error","Format tanggal YYYY-MM-DD"); return
            deb_code = cb_debit.get().split(" - ")[0].strip()
            cred_code = cb_credit.get().split(" - ")[0].strip()
            if deb_code == cred_code:
                messagebox.showerror("Error","Akun debit & kredit tidak boleh sama"); return
            try:
                amt = float(e_amount.get().replace(",",""))
                if amt <= 0: raise ValueError
            except:
                messagebox.showerror("Error","Nominal harus angka > 0"); return

            deb_name = self._get_account_name(deb_code); cred_name = self._get_account_name(cred_code)
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, desc, deb_code, deb_name, amt, 0.0))
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, desc, cred_code, cred_name, 0.0, amt))
            self.conn.commit()
            messagebox.showinfo("Sukses","Jurnal tersimpan")
            win.destroy(); self.refresh_all()

        ttk.Button(frm, text="Simpan Jurnal", command=do_save).grid(row=3, column=3, pady=10, sticky="e")

    def _get_account_name(self, code):
        self.cur.execute("SELECT name FROM accounts WHERE code=?", (code,))
        r = self.cur.fetchone()
        return r[0] if r else ""
    
    # ---------------- Ledger (Buku Besar) ----------------
        # Method untuk ambil semua akun
    def get_all_accounts(self):
        try:
            self.cur.execute("SELECT code || ' - ' || name FROM accounts ORDER BY code")
            rows = self.cur.fetchall()
            return [r[0] for r in rows]
        except sqlite3.OperationalError as e:
            messagebox.showerror("Database Error", f"Gagal mengambil akun:\n{e}")
            return []

    # Fungsi T-Account / Buku Besar
    def open_ledger_window(self):
        win = tk.Toplevel(self.root)
        win.title("Buku Besar (T-Account)")
        win.geometry("850x550")

        ttk.Label(win, text="Pilih Akun", font=("Segoe UI", 11)).pack(pady=5)

        # Ambil semua akun
        akun_list = self.get_all_accounts()
        akun_combo = ttk.Combobox(win, values=akun_list, width=50, state="readonly")
        akun_combo.pack(pady=5)

        frame_table = ttk.Frame(win)
        frame_table.pack(fill="both", expand=True, padx=10, pady=10)

        tree = ttk.Treeview(
            frame_table,
            columns=("tanggal", "deskripsi", "debit", "kredit", "saldo"),
            show="headings",
            height=20
        )
        tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        tree.heading("tanggal", text="Tanggal")
        tree.heading("deskripsi", text="Deskripsi")
        tree.heading("debit", text="Debit")
        tree.heading("kredit", text="Kredit")
        tree.heading("saldo", text="Saldo")

        tree.column("tanggal", width=100)
        tree.column("deskripsi", width=300)
        tree.column("debit", width=100, anchor="e")
        tree.column("kredit", width=100, anchor="e")
        tree.column("saldo", width=120, anchor="e")

        def format_currency(value):
            return f"{value:,.2f}"

        def load_ledger():
            selected = akun_combo.get()
            if not selected:
                messagebox.showwarning("Peringatan", "Pilih akun terlebih dahulu!")
                return

            kode_akun = selected.split(" - ")[0].strip()

            # Tentukan tipe akun berdasarkan kode akun
            if kode_akun.startswith("1"):
                tipe_akun = "Aset"
            elif kode_akun.startswith("2"):
                tipe_akun = "Liabilitas"
            elif kode_akun.startswith("3"):
                tipe_akun = "Ekuitas"
            elif kode_akun.startswith("4"):
                tipe_akun = "Pendapatan"
            else:
                tipe_akun = "Beban"

            conn = sqlite3.connect("akuntansi.db")
            cur = conn.cursor()
            try:
                cur.execute("""
                    SELECT tanggal, deskripsi, debit, kredit
                    FROM journal
                    WHERE account_code=?
                    ORDER BY tanggal, id
                """, (kode_akun,))
                rows = cur.fetchall()
            finally:
                conn.close()

            tree.delete(*tree.get_children())
            saldo = 0
            current_month = None

            for r in rows:
                tanggal_str, desk, debit, kredit = r
                debit = float(debit) if debit else 0
                kredit = float(kredit) if kredit else 0

                # Hitung saldo berdasarkan normal balance
                if tipe_akun in ["Aset", "Beban"]:
                    saldo += debit - kredit
                else:  # Liabilitas, Ekuitas, Pendapatan
                    saldo += kredit - debit

                # Subtotal per bulan
                tanggal_obj = datetime.strptime(tanggal_str, "%Y-%m-%d")
                bulan = tanggal_obj.strftime("%Y-%m")
                if current_month and bulan != current_month:
                    tree.insert(
                        "",
                        tk.END,
                        values=("", f"Subtotal {current_month}", "", "", f"{saldo:,.2f}"),
                        tags=("subtotal",)
                    )
                current_month = bulan

                tree.insert(
                    "",
                    tk.END,
                    values=(
                        tanggal_str,
                        desk,
                        f"{debit:,.2f}" if debit != 0 else "",
                        f"{kredit:,.2f}" if kredit != 0 else "",
                        f"{saldo:,.2f}"
                    ),
                    tags=("negative",) if saldo < 0 else ()
                )

            tree.tag_configure("negative", foreground="red")
            tree.tag_configure("subtotal", background="#e0e0e0", font=("Segoe UI", 10, "bold"))

        ttk.Button(win, text="Tampilkan Buku Besar", command=load_ledger).pack(pady=10)


    # ---------------- Purchases (Pembelian Persediaan) ----------------
    def open_purchase_input(self):
        win = tk.Toplevel(self.root); win.title("Pembelian Persediaan"); win.geometry("740x320")
        frm = ttk.Frame(win, padding=8); frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Tanggal").grid(row=0,column=0); e_date = ttk.Entry(frm); e_date.grid(row=0,column=1); e_date.insert(0, datetime.today().strftime("%Y-%m-%d"))
        ttk.Label(frm, text="No Dokumen").grid(row=0,column=2); e_doc = ttk.Entry(frm); e_doc.grid(row=0,column=3)

        self.cur.execute("SELECT code,name FROM accounts WHERE code LIKE '11%' ORDER BY code")
        stock_vals = [f"{c} - {n}" for c,n in self.cur.fetchall()]
        ttk.Label(frm, text="Barang (Persediaan)").grid(row=1,column=0)
        cb_stock = ttk.Combobox(frm, values=stock_vals, width=40); cb_stock.grid(row=1,column=1)
        ttk.Label(frm, text="Qty").grid(row=1,column=2); e_qty = ttk.Entry(frm); e_qty.grid(row=1,column=3)
        ttk.Label(frm, text="Harga/unit").grid(row=2,column=0); e_price = ttk.Entry(frm); e_price.grid(row=2,column=1)
        ttk.Label(frm, text="Akun Lawan (Kas/Bank/Utang)").grid(row=3,column=0)
        self.load_accounts_cache()
        all_vals = [f"{c} - {n}" for c,n in self.accounts]
        cb_lawan = ttk.Combobox(frm, values=all_vals, width=40); cb_lawan.grid(row=3,column=1)

        def do_purchase():
            tanggal = e_date.get().strip(); doc = e_doc.get().strip()
            if not cb_stock.get() or not e_qty.get().strip() or not e_price.get().strip() or not cb_lawan.get():
                messagebox.showerror("Error","Isi semua field"); return
            stock_code = cb_stock.get().split(" - ")[0].strip(); stock_name = cb_stock.get().split(" - ")[1].strip()
            try:
                qty = to_decimal(e_qty.get().replace(",",""))
                price = to_decimal(e_price.get().replace(",",""))
                total = qty * price
            except:
                messagebox.showerror("Error","Qty/Harga tidak valid"); return
            self.cur.execute("""INSERT INTO purchases (tanggal, doc_no, account_code, account_name, qty, unit_price, total)
                                VALUES (?,?,?,?,?,?,?)""", (tanggal, doc, stock_code, stock_name, float(qty), float(price), float(total)))
            lawan_code = cb_lawan.get().split(" - ")[0].strip(); lawan_name = cb_lawan.get().split(" - ")[1].strip()
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, f"Pembelian {stock_name} ({doc})", stock_code, stock_name, float(total), 0.0))
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, f"Pembelian {stock_name} ({doc})", lawan_code, lawan_name, 0.0, float(total)))
            self.conn.commit()
            messagebox.showinfo("Sukses","Pembelian tersimpan dan diposting ke Jurnal")
            win.destroy(); self.refresh_all()

        ttk.Button(frm, text="Simpan Pembelian", command=do_purchase).grid(row=4,column=3,pady=8)

    # ---------------- Usage (Pemakaian Persediaan) ----------------
    def open_usage_input(self):
        win = tk.Toplevel(self.root); win.title("Pemakaian Persediaan (Usage Out)"); win.geometry("760x380")
        frm = ttk.Frame(win, padding=8); frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="Tanggal").grid(row=0,column=0); e_date = ttk.Entry(frm); e_date.grid(row=0,column=1); e_date.insert(0, datetime.today().strftime("%Y-%m-%d"))
        ttk.Label(frm, text="No Dokumen").grid(row=0,column=2); e_doc = ttk.Entry(frm); e_doc.grid(row=0,column=3)

        self.cur.execute("SELECT code,name FROM accounts WHERE code LIKE '11%' ORDER BY code")
        stock_vals = [f"{c} - {n}" for c,n in self.cur.fetchall()]
        ttk.Label(frm, text="Barang (Persediaan)").grid(row=1,column=0); cb_stock = ttk.Combobox(frm, values=stock_vals, width=40); cb_stock.grid(row=1,column=1)
        ttk.Label(frm, text="Qty keluar").grid(row=1,column=2); e_qty = ttk.Entry(frm); e_qty.grid(row=1,column=3)
        ttk.Label(frm, text="Akun Beban (contoh: 501)").grid(row=2,column=0)
        self.load_accounts_cache()
        all_vals = [f"{c} - {n}" for c,n in self.accounts]
        cb_beban = ttk.Combobox(frm, values=all_vals, width=40); cb_beban.grid(row=2,column=1)
        ttk.Label(frm, text="Metode HPP").grid(row=3,column=0)
        cb_method = ttk.Combobox(frm, values=["FIFO","AVG"], width=10, state="readonly", textvariable=self.inventory_method); cb_method.grid(row=3,column=1)
        cb_method.set(self.inventory_method.get())
        lbl_info = ttk.Label(frm, text="Cost per unit will be computed based on selected method."); lbl_info.grid(row=4,column=0, columnspan=4, pady=6)

        def do_usage():
            tanggal = e_date.get().strip(); doc = e_doc.get().strip()
            if not cb_stock.get() or not e_qty.get().strip() or not cb_beban.get():
                messagebox.showerror("Error","Isi semua field"); return
            stock_code = cb_stock.get().split(" - ")[0].strip(); stock_name = cb_stock.get().split(" - ")[1].strip()
            try:
                qty = to_decimal(e_qty.get().replace(",",""))
                if qty <= 0: raise ValueError
            except:
                messagebox.showerror("Error","Qty tidak valid"); return
            beban_code = cb_beban.get().split(" - ")[0].strip(); beban_name = cb_beban.get().split(" - ")[1].strip()
            method = cb_method.get()
            if method == "AVG":
                cpu = self._compute_moving_average_cost(stock_code)
                total_cost = cpu * qty
            else:
                cpu, total_cost = self._compute_fifo_cost_and_consume(stock_code, qty)
            # save usage_out
            self.cur.execute("""INSERT INTO usage_out
                (tanggal, doc_no, account_code, account_name, qty, cost_per_unit, total_cost, expense_account_code, expense_account_name, keterangan)
                VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (tanggal, doc, stock_code, stock_name, float(qty), float(cpu), float(total_cost), beban_code, beban_name, f"Pemakaian {stock_name}"))
            # journals: Dr Beban, Cr Persediaan
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, f"Pemakaian {stock_name} ({doc})", beban_code, beban_name, float(total_cost), 0.0))
            self.cur.execute("""INSERT INTO journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)""",
                             (tanggal, f"Pemakaian {stock_name} ({doc})", stock_code, stock_name, 0.0, float(total_cost)))
            self.conn.commit()
            messagebox.showinfo("Sukses", f"Pemakaian tersimpan. Total cost = {total_cost:,.2f} (cpu {cpu:,.2f})")
            win.destroy(); self.refresh_all()

        ttk.Button(frm, text="Simpan Pemakaian", command=do_usage).grid(row=5,column=3,pady=8)

    # ---------------- Inventory calculations ----------------
    def _compute_moving_average_cost(self, stock_code):
        self.cur.execute("SELECT COALESCE(SUM(qty*unit_price),0), COALESCE(SUM(qty),0) FROM purchases WHERE account_code=?", (stock_code,))
        total_cost, total_qty = self.cur.fetchone()
        total_cost = to_decimal(total_cost); total_qty = to_decimal(total_qty)
        self.cur.execute("SELECT COALESCE(SUM(qty),0), COALESCE(SUM(total_cost),0) FROM usage_out WHERE account_code=?", (stock_code,))
        used_qty, used_cost = self.cur.fetchone()
        used_qty = to_decimal(used_qty); used_cost = to_decimal(used_cost)
        remaining_qty = total_qty - used_qty
        remaining_cost = total_cost - used_cost
        if remaining_qty <= 0:
            return Decimal("0.00")
        return (remaining_cost / remaining_qty).quantize(Decimal("0.01"))

    def _compute_fifo_cost_and_consume(self, stock_code, qty_needed):
        qty_needed = to_decimal(qty_needed)
        self.cur.execute("SELECT id, qty, unit_price FROM purchases WHERE account_code=? ORDER BY id", (stock_code,))
        batches = [(r[0], to_decimal(r[1]), to_decimal(r[2])) for r in self.cur.fetchall()]
        self.cur.execute("SELECT COALESCE(SUM(qty),0) FROM usage_out WHERE account_code=?", (stock_code,))
        total_used = to_decimal(self.cur.fetchone()[0])
        # Build available batches after previous consumption
        avail_batches = []
        consumed = total_used
        for pid, bqty, bprice in batches:
            if consumed >= bqty:
                consumed -= bqty
                continue
            else:
                avail = bqty - consumed
                avail_batches.append([avail, bprice])
                consumed = Decimal("0.00")
        remaining = qty_needed
        total_cost = Decimal("0.00")
        for i in range(len(avail_batches)):
            if remaining <= 0:
                break
            avail, price = avail_batches[i]
            take = min(avail, remaining)
            total_cost += take * price
            avail_batches[i][0] -= take
            remaining -= take
        if remaining > 0:
            messagebox.showwarning("Warning", "Qty pemakaian melebihi stok yang tersedia (FIFO). Sistem memproses sisa dengan cost 0.")
        cpu = (total_cost / qty_needed) if qty_needed > 0 else Decimal("0.00")
        return (cpu.quantize(Decimal("0.01")), float(total_cost))

    # ---------------- Adjustment Journal ----------------
    def open_adjustment_input(self):
        win = tk.Toplevel(self.root); win.title("Jurnal Penyesuaian"); win.geometry("720x260")
        frm = ttk.Frame(win, padding=8); frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="Tanggal").grid(row=0,column=0); e_date = ttk.Entry(frm); e_date.grid(row=0,column=1); e_date.insert(0, datetime.today().strftime("%Y-%m-%d"))
        ttk.Label(frm, text="Deskripsi").grid(row=0,column=2); e_desc = ttk.Entry(frm, width=40); e_desc.grid(row=0,column=3)
        self.load_accounts_cache(); vals = [f"{c} - {n}" for c,n in self.accounts]
        ttk.Label(frm, text="Akun").grid(row=1,column=0); cb_acc = ttk.Combobox(frm, values=vals, width=40); cb_acc.grid(row=1,column=1)
        ttk.Label(frm, text="Debit").grid(row=2,column=0); e_debit = ttk.Entry(frm); e_debit.grid(row=2,column=1)
        ttk.Label(frm, text="Kredit").grid(row=2,column=2); e_credit = ttk.Entry(frm); e_credit.grid(row=2,column=3)

        def save_adj():
            t = e_date.get().strip(); d = e_desc.get().strip()
            if not cb_acc.get() or (not e_debit.get().strip() and not e_credit.get().strip()):
                messagebox.showerror("Error","Isi semua field"); return
            code = cb_acc.get().split(" - ")[0].strip(); name = cb_acc.get().split(" - ")[1].strip()
            try:
                deb = float(e_debit.get().replace(",","")) if e_debit.get().strip() else 0.0
                kre = float(e_credit.get().replace(",","")) if e_credit.get().strip() else 0.0
            except:
                messagebox.showerror("Error","Nominal tidak valid"); return
            self.cur.execute("INSERT INTO adjustment_journal (tanggal,deskripsi,account_code,account_name,debit,kredit) VALUES (?,?,?,?,?,?)",
                             (t,d,code,name,deb,kre))
            self.conn.commit()
            messagebox.showinfo("Sukses","Jurnal penyesuaian disimpan")
            win.destroy(); self.refresh_all()
        ttk.Button(frm, text="Simpan Penyesuaian", command=save_adj).grid(row=3,column=3,pady=8)

    # ---------------- Inventory Card (Stock Card) ----------------
    def open_inventory_card(self):
        win = tk.Toplevel(self.root); win.title("Kartu Persediaan (Stock Card)"); win.geometry("900x600")
        top = ttk.Frame(win); top.pack(fill=tk.X, pady=6)
        ttk.Label(top, text="Pilih Barang (akun 11xx)").pack(side=tk.LEFT, padx=6)
        self.cur.execute("SELECT code,name FROM accounts WHERE code LIKE '11%' ORDER BY code")
        stock_vals = [f"{c} - {n}" for c,n in self.cur.fetchall()]
        cb_stock = ttk.Combobox(top, values=stock_vals, width=40); cb_stock.pack(side=tk.LEFT, padx=6)
        ttk.Label(top, text="Metode").pack(side=tk.LEFT, padx=6)
        cb_method = ttk.Combobox(top, values=["FIFO","AVG"], width=8, state="readonly", textvariable=self.inventory_method); cb_method.pack(side=tk.LEFT)
        cb_method.set(self.inventory_method.get())

        tree = ttk.Treeview(win, columns=("Tanggal","Doc","In_Qty","In_Unit","In_Amount","Out_Qty","Out_Unit","Out_Amount","Bal_Qty","Bal_Unit","Bal_Amount"), show="headings")
        headers = [("Tanggal",100),("Doc",80),("In_Qty",70),("In_Unit",80),("In_Amount",100),("Out_Qty",70),("Out_Unit",80),("Out_Amount",100),("Bal_Qty",70),("Bal_Unit",80),("Bal_Amount",110)]
        for h,w in headers:
            tree.heading(h, text=h); tree.column(h, width=w, anchor="w")
        tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        def load_card():
            for r in tree.get_children(): tree.delete(r)
            sel = cb_stock.get()
            if not sel: return
            code = sel.split(" - ")[0]
            method = cb_method.get()
            rows = []
            self.cur.execute("SELECT tanggal, doc_no, qty, unit_price, total FROM purchases WHERE account_code=? ORDER BY tanggal, id", (code,))
            for t,doc,qty,unit,total in self.cur.fetchall():
                rows.append(("IN", t, doc, to_decimal(qty), to_decimal(unit), to_decimal(total)))
            self.cur.execute("SELECT tanggal, doc_no, qty, cost_per_unit, total_cost FROM usage_out WHERE account_code=? ORDER BY tanggal, id", (code,))
            for t,doc,qty,cpu,total in self.cur.fetchall():
                rows.append(("OUT", t, doc, to_decimal(qty), to_decimal(cpu), to_decimal(total)))
            rows.sort(key=lambda x: x[1])
            bal_qty = Decimal("0.00"); bal_amount = Decimal("0.00")
            for typ, t, doc, q, unit, amt in rows:
                if typ == "IN":
                    bal_qty += q; bal_amount += amt
                    bal_unit = (bal_amount / bal_qty) if bal_qty>0 else Decimal("0.00")
                    tree.insert("", tk.END, values=(t, doc, f"{q}", f"{unit:,.2f}", f"{amt:,.2f}", "", "", "", f"{bal_qty}", f"{bal_unit:,.2f}", f"{bal_amount:,.2f}"))
                else:
                    if method == "AVG":
                        cpu = (bal_amount / bal_qty) if bal_qty>0 else Decimal("0.00")
                        out_amount = (cpu * q)
                        bal_qty -= q
                        bal_amount -= out_amount
                        bal_unit = (bal_amount / bal_qty) if bal_qty>0 else Decimal("0.00")
                        tree.insert("", tk.END, values=(t, doc, "", "", "", f"{q}", f"{cpu:,.2f}", f"{out_amount:,.2f}", f"{bal_qty}", f"{bal_unit:,.2f}", f"{bal_amount:,.2f}"))
                    else:
                        lots = []
                        self.cur.execute("SELECT qty, unit_price FROM purchases WHERE account_code=? ORDER BY id", (code,))
                        for bqty,bprice in self.cur.fetchall():
                            lots.append([to_decimal(bqty), to_decimal(bprice)])
                        # calculate used before this out
                        used_before = Decimal("0.00")
                        self.cur.execute("SELECT qty FROM usage_out WHERE account_code=? AND tanggal<=? ORDER BY id", (code, t))
                        used_before = sum(to_decimal(r[0]) for r in self.cur.fetchall())
                        # subtract used_before from lots
                        ub = used_before
                        new_lots = []
                        for l in lots:
                            if ub >= l[0]:
                                ub -= l[0]; continue
                            else:
                                new_lots.append([l[0]-ub, l[1]]); ub = Decimal("0.00")
                        remaining = q
                        out_amount = Decimal("0.00")
                        for ln in new_lots:
                            if remaining<=0: break
                            take = min(ln[0], remaining)
                            out_amount += take * ln[1]
                            ln[0] -= take
                            remaining -= take
                        total_qty = sum(l[0] for l in new_lots)
                        total_amount = sum(l[0]*l[1] for l in new_lots)
                        bal_qty = total_qty - q
                        bal_amount = (total_amount - out_amount)
                        bal_unit = (bal_amount / bal_qty) if bal_qty>0 else Decimal("0.00")
                        tree.insert("", tk.END, values=(t, doc, "", "", "", f"{q}", "", f"{out_amount:,.2f}", f"{bal_qty}", f"{bal_unit:,.2f}", f"{bal_amount:,.2f}"))
            # finished

        ttk.Button(win, text="Tampilkan Kartu", command=load_card).pack(pady=6)

    # ---------------- Reports (Menu C) ----------------
    def open_reports_window(self):
        win = tk.Toplevel(self.root); win.title("Laporan Keuangan"); win.geometry("1000x700")
        nb = ttk.Notebook(win); nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        tab_lr = ttk.Frame(nb); nb.add(tab_lr, text="Laba Rugi")
        tree_lr = ttk.Treeview(tab_lr, columns=("Akun","Nama","Jumlah"), show="headings")
        tree_lr.heading("Akun", text="Kode"); tree_lr.heading("Nama", text="Nama"); tree_lr.heading("Jumlah", text="Jumlah"); tree_lr.pack(fill=tk.BOTH, expand=True)
        tab_pm = ttk.Frame(nb); nb.add(tab_pm, text="Perubahan Modal")
        tree_pm = ttk.Treeview(tab_pm, columns=("Keterangan","Jumlah"), show="headings"); tree_pm.heading("Keterangan", text="Keterangan"); tree_pm.heading("Jumlah", text="Jumlah"); tree_pm.pack(fill=tk.BOTH, expand=True)
        tab_neraca = ttk.Frame(nb); nb.add(tab_neraca, text="Neraca")
        tree_neraca = ttk.Treeview(tab_neraca, columns=("Akun","Nama","Jumlah"), show="headings")
        for col in ("Akun","Nama","Jumlah"): tree_neraca.heading(col, text=col)
        tree_neraca.pack(fill=tk.BOTH, expand=True)

        frm = ttk.Frame(win); frm.pack(fill=tk.X)
        ttk.Button(frm, text="Export Laba Rugi CSV", command=self.export_lr_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(frm, text="Export Neraca CSV", command=self.export_neraca_csv).pack(side=tk.LEFT, padx=6)
        ttk.Button(frm, text="Export Laba Rugi Excel", command=self.export_lr_excel).pack(side=tk.LEFT, padx=6)
        ttk.Button(frm, text="Export Neraca Excel", command=self.export_neraca_excel).pack(side=tk.LEFT, padx=6)
        ttk.Button(frm, text="Export Laba Rugi PDF", command=self.export_lr_pdf).pack(side=tk.LEFT, padx=6)
        ttk.Button(frm, text="Export Neraca PDF", command=self.export_neraca_pdf).pack(side=tk.LEFT, padx=6)

        def fill_all():
            tree_lr.delete(*tree_lr.get_children())
            total_pend = Decimal("0.00"); total_beban = Decimal("0.00")
            self.cur.execute("""SELECT account_code, account_name, COALESCE(SUM(kredit),0) as kred, COALESCE(SUM(debit),0) as deb FROM journal GROUP BY account_code,account_name""")
            rows = self.cur.fetchall()
            for code,name,kred,deb in rows:
                if code.startswith("4"):
                    total_pend += to_decimal(kred) - to_decimal(deb)
                elif code.startswith("5"):
                    total_beban += to_decimal(deb) - to_decimal(kred)
            for code,name,kred,deb in rows:
                if code.startswith("4"):
                    val = to_decimal(kred) - to_decimal(deb); tree_lr.insert("", tk.END, values=(code, name, f"{val:,.2f}"))
                elif code.startswith("5"):
                    val = to_decimal(deb) - to_decimal(kred); tree_lr.insert("", tk.END, values=(code, name, f"{val:,.2f}"))
            tree_lr.insert("", tk.END, values=("", "TOTAL PENDAPATAN", f"{total_pend:,.2f}"))
            tree_lr.insert("", tk.END, values=("", "TOTAL BEBAN", f"{total_beban:,.2f}"))
            laba = total_pend - total_beban
            tree_lr.insert("", tk.END, values=("", "LABA / RUGI BERSIH", f"{laba:,.2f}"))

            tree_pm.delete(*tree_pm.get_children())
            self.cur.execute("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(kredit),0) FROM journal WHERE account_code='301'")
            deb301,k301 = self.cur.fetchone()
            modal_awal = to_decimal(k301) - to_decimal(deb301)
            tree_pm.insert("", tk.END, values=("Modal Awal (301)", f"{modal_awal:,.2f}"))
            tree_pm.insert("", tk.END, values=("Laba Bersih", f"{laba:,.2f}"))
            modal_akhir = modal_awal + laba
            tree_pm.insert("", tk.END, values=("Modal Akhir", f"{modal_akhir:,.2f}"))

            tree_neraca.delete(*tree_neraca.get_children())
            total_assets = Decimal("0.00"); total_liab = Decimal("0.00"); total_equity = Decimal("0.00")
            for code,name,kred,deb in rows:
                bal = to_decimal(deb) - to_decimal(kred)
                if code.startswith("1"):
                    total_assets += bal; tree_neraca.insert("", tk.END, values=(code,name,f"{bal:,.2f}"))
                elif code.startswith("2"):
                    total_liab += -bal; tree_neraca.insert("", tk.END, values=(code,name,f"{-bal:,.2f}"))
                elif code.startswith("3"):
                    total_equity += -bal; tree_neraca.insert("", tk.END, values=(code,name,f"{-bal:,.2f}"))
            tree_neraca.insert("", tk.END, values=("", "TOTAL ASET", f"{total_assets:,.2f}"))
            tree_neraca.insert("", tk.END, values=("", "TOTAL LIABILITAS", f"{total_liab:,.2f}"))
            tree_neraca.insert("", tk.END, values=("", "TOTAL EKUITAS", f"{total_equity:,.2f}"))

        fill_all()

    # ---------------- Export CSV / Excel / PDF ----------------
    def export_lr_csv(self):
        fn = f"laba_rugi_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(fn, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f); w.writerow(["Kode","Nama","Jumlah"])
            self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(kredit),0) as kred, COALESCE(SUM(debit),0) as deb FROM journal GROUP BY account_code,account_name ORDER BY account_code")
            for code,name,kred,deb in self.cur.fetchall():
                if code.startswith("4"):
                    val = to_decimal(kred) - to_decimal(deb); w.writerow([code,name,f"{val:.2f}"])
                elif code.startswith("5"):
                    val = to_decimal(deb) - to_decimal(kred); w.writerow([code,name,f"{val:.2f}"])
        messagebox.showinfo("Export", f"Laporan Laba Rugi diexport ke {fn}")

    def export_neraca_csv(self):
        fn = f"neraca_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        with open(fn, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f); w.writerow(["Kode","Nama","Saldo"])
            self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(debit),0) as deb, COALESCE(SUM(kredit),0) as kred FROM journal GROUP BY account_code,account_name ORDER BY account_code")
            for code,name,deb,kred in self.cur.fetchall():
                bal = to_decimal(deb) - to_decimal(kred)
                if code.startswith("1"):
                    w.writerow([code,name,f"{bal:.2f}"])
                elif code.startswith("2"):
                    w.writerow([code,name,f"{-bal:.2f}"])
                elif code.startswith("3"):
                    w.writerow([code,name,f"{-bal:.2f}"])
        messagebox.showinfo("Export", f"Neraca diexport ke {fn}")

    def export_lr_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing", "openpyxl belum terinstall. Jalankan: pip install openpyxl")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"laba_rugi_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not fn: return
        wb = Workbook(); ws = wb.active; ws.title = "Laba Rugi"
        ws.append(["Kode", "Nama", "Jumlah"])
        self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(kredit),0) as kred, COALESCE(SUM(debit),0) as deb FROM journal GROUP BY account_code,account_name ORDER BY account_code")
        for code,name,kred,deb in self.cur.fetchall():
            if code.startswith("4"):
                val = to_decimal(kred) - to_decimal(deb); ws.append([code,name,float(val)])
            elif code.startswith("5"):
                val = to_decimal(deb) - to_decimal(kred); ws.append([code,name,float(val)])
        wb.save(fn); messagebox.showinfo("Export", f"Laba Rugi diexport ke {fn}")

    def export_neraca_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing", "openpyxl belum terinstall. Jalankan: pip install openpyxl")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"neraca_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not fn: return
        wb = Workbook(); ws = wb.active; ws.title = "Neraca"
        ws.append(["Kode", "Nama", "Saldo"])
        self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(debit),0) as deb, COALESCE(SUM(kredit),0) as kred FROM journal GROUP BY account_code,account_name ORDER BY account_code")
        for code,name,deb,kred in self.cur.fetchall():
            bal = to_decimal(deb) - to_decimal(kred)
            if code.startswith("1"):
                ws.append([code,name,float(bal)])
            elif code.startswith("2"):
                ws.append([code,name,float(-bal)])
            elif code.startswith("3"):
                ws.append([code,name,float(-bal)])
        wb.save(fn); messagebox.showinfo("Export", f"Neraca diexport ke {fn}")

    def export_lr_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Missing", "reportlab belum terinstall. Jalankan: pip install reportlab")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"laba_rugi_{datetime.now().strftime('%Y%m%d')}.pdf")
        if not fn: return
        c = canvas.Canvas(fn)
        c.setFont("Helvetica-Bold", 14); c.drawString(50, 800, "Laporan Laba Rugi - Desa Wisata Kandri")
        y = 770; c.setFont("Helvetica", 10)
        self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(kredit),0) as kred, COALESCE(SUM(debit),0) as deb FROM journal GROUP BY account_code,account_name ORDER BY account_code")
        total_pend = Decimal("0.00"); total_beban = Decimal("0.00")
        for code,name,kred,deb in self.cur.fetchall():
            if code.startswith("4"):
                val = to_decimal(kred) - to_decimal(deb); total_pend += val; c.drawString(50,y,f"{code} {name} .. {val:,.2f}"); y -= 16
            elif code.startswith("5"):
                val = to_decimal(deb) - to_decimal(kred); total_beban += val; c.drawString(50,y,f"{code} {name} .. {val:,.2f}"); y -= 16
            if y < 80: c.showPage(); y = 800
        c.drawString(50, y-10, f"Total Pendapatan: {total_pend:,.2f}"); c.drawString(50, y-30, f"Total Beban: {total_beban:,.2f}")
        c.save(); messagebox.showinfo("Export", f"Laba Rugi PDF dibuat: {fn}")

    def export_neraca_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Missing", "reportlab belum terinstall. Jalankan: pip install reportlab")
            return
        fn = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"neraca_{datetime.now().strftime('%Y%m%d')}.pdf")
        if not fn: return
        c = canvas.Canvas(fn); c.setFont("Helvetica-Bold", 14); c.drawString(50, 800, "Neraca - Desa Wisata Kandri")
        y = 770; c.setFont("Helvetica", 10)
        self.cur.execute("SELECT account_code, account_name, COALESCE(SUM(debit),0) as deb, COALESCE(SUM(kredit),0) as kred FROM journal GROUP BY account_code,account_name ORDER BY account_code")
        total_assets = Decimal("0.00"); total_liab = Decimal("0.00"); total_equity = Decimal("0.00")
        for code,name,deb,kred in self.cur.fetchall():
            bal = to_decimal(deb) - to_decimal(kred)
            if code.startswith("1"):
                total_assets += bal; c.drawString(50,y,f"{code} {name} .. {bal:,.2f}"); y -= 14
            elif code.startswith("2"):
                total_liab += -bal; c.drawString(50,y,f"{code} {name} .. {-bal:,.2f}"); y -= 14
            elif code.startswith("3"):
                total_equity += -bal; c.drawString(50,y,f"{code} {name} .. {-bal:,.2f}"); y -= 14
            if y < 80: c.showPage(); y = 800
        c.drawString(50, y-10, f"Total Aset: {total_assets:,.2f}"); c.drawString(50, y-30, f"Total Liabilitas: {total_liab:,.2f}"); c.drawString(50, y-50, f"Total Ekuitas: {total_equity:,.2f}")
        c.save(); messagebox.showinfo("Export", f"Neraca PDF dibuat: {fn}")

    # ---------------- Refresh views ----------------
    def refresh_journal_quick(self):
        for r in self.tree_journal.get_children(): self.tree_journal.delete(r)
        self.cur.execute("SELECT tanggal, deskripsi, account_name, account_code, debit, kredit FROM journal ORDER BY tanggal, id")
        for row in self.cur.fetchall():
            self.tree_journal.insert("", tk.END, values=(row[0], row[1], row[2], row[3], f"{row[4]:,.2f}" if row[4] else "", f"{row[5]:,.2f}" if row[5] else ""))

    def refresh_tb_quick(self):
        for r in self.tree_tb.get_children(): self.tree_tb.delete(r)
        total_d = Decimal("0.00"); total_k = Decimal("0.00")
        self.cur.execute("SELECT code,name FROM accounts ORDER BY code")
        for code,name in self.cur.fetchall():
            self.cur.execute("SELECT COALESCE(SUM(debit),0), COALESCE(SUM(kredit),0) FROM journal WHERE account_code=?", (code,))
            d,k = self.cur.fetchone(); d = to_decimal(d); k = to_decimal(k)
            total_d += d; total_k += k
            self.tree_tb.insert("", tk.END, values=(code,name, f"{d:,.2f}", f"{k:,.2f}"))
        self.tb_label.config(text=f"TOTAL DEBIT: {total_d:,.2f}    |    TOTAL KREDIT: {total_k:,.2f}")

    def refresh_all(self):
        self.load_accounts_cache()
        self.refresh_journal_quick()
        self.refresh_tb_quick()

# ---------------- Run ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
