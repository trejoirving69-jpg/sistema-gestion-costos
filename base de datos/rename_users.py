from conexion import obtener_conexion

conn = obtener_conexion()
cur = conn.cursor()
cur.execute("UPDATE usuarios SET username = REPLACE(username,' ','_') WHERE username LIKE '% %'")
conn.commit()
cur.execute('SELECT username, rol FROM usuarios ORDER BY username')
rows = cur.fetchall()
for r in rows:
    print(r)
conn.close()
