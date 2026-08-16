import sqlite3, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
conn = sqlite3.connect('gestion_sistema.db')
cur = conn.cursor()
cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
rows = cur.fetchall()
for r in rows:
    print(r[0])
conn.close()
