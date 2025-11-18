import sqlite3

conn = sqlite3.connect("akuntansi.db")
cur = conn.cursor()

print("=== DAFTAR TABEL ===")
cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
print(cur.fetchall())

print("\n=== STRUKTUR TABEL journal ===")
try:
    cur.execute("PRAGMA table_info(journal)")
    print(cur.fetchall())
except:
    print("Tabel journal tidak ditemukan.")

print("\n=== SAMPLE ISI jurnal ===")
try:
    cur.execute("SELECT * FROM journal LIMIT 5")
    print(cur.fetchall())
except:
    print("Tabel journal tidak bisa dibaca.")

conn.close()
