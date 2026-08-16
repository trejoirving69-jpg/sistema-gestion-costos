"""
Script sencillo para añadir/actualizar usuarios en la base de datos `gestion_sistema.db`.
"""

from conexion import obtener_conexion, hash_password

# Cambia/añade aquí los usuarios que quieras. 
USUARIOS = [

    ("carlos_macias", "macias1975", "Usuario"),
    ("isabel_napolitano", "napolitano1975", "Usuario"),
    
]

# Hashing para mayor seguridad
HASH_PASSWORDS = True


def main():
    conn = obtener_conexion()
    cur = conn.cursor()

    for username, password, rol in USUARIOS:
        # Forzar que no sean administradores
        if str(rol).strip().lower() in ("administrador", "admin"):
            print(f"Aviso: rol '{rol}' no permitido para usuario '{username}'; se usará 'Usuario' en su lugar.")
            rol = "Usuario"

        pw_to_store = hash_password(password) if HASH_PASSWORDS else password

        cur.execute("SELECT id FROM usuarios WHERE username = ?", (username,))
        row = cur.fetchone()
        if row:
            cur.execute(
                "UPDATE usuarios SET password_hash = ?, rol = ? WHERE username = ?",
                (pw_to_store, rol, username)
            )
            print(f"Actualizado: {username} -> rol={rol}")
        else:
            cur.execute(
                "INSERT INTO usuarios (username, password_hash, rol) VALUES (?, ?, ?)",
                (username, pw_to_store, rol)
            )
            print(f"Insertado: {username} -> rol={rol}")

    conn.commit()
    conn.close()


if __name__ == '__main__':
    main()
