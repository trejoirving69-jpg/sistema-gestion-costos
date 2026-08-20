import os
import sqlite3
import uuid
import secrets
from pathlib import Path
from functools import wraps

from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename

# ============================================================
# CONFIGURACIÓN
# ============================================================

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent

# IMPORTANTE:
# Usamos la MISMA base de datos que utiliza el sistema principal.
DB_PATH = BASE_DIR / "gestion_sistema.db"

# Carpeta para archivos enviados por el chat.
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "archivos"

DATA_DIR.mkdir(parents=True, exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Clave interna para proteger la API.
# En producción la configuraremos mediante variable de entorno.
API_KEY = os.getenv(
    "CHAT_API_KEY",
    "CAMBIAR_ESTA_CLAVE_EN_PRODUCCION"
)

# Sesiones temporales del servidor.
# Para producción posteriormente las pasaremos a almacenamiento persistente.
SESIONES = {}

# ============================================================
# BASE DE DATOS
# ============================================================

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_password(password):
    """
    Usa exactamente el mismo algoritmo que tu sistema:
    SHA-256 hexadecimal.
    """
    import hashlib
    return hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()


def init_db():
    """
    Crea solamente las tablas necesarias para el chat.
    NO reemplaza ni modifica las tablas existentes del sistema.
    """

    conn = db()

    conn.executescript("""
        CREATE TABLE IF NOT EXISTS conversaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            creado_en TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS participantes_chat (
            conversacion_id INTEGER NOT NULL,
            usuario_id INTEGER NOT NULL,

            PRIMARY KEY (
                conversacion_id,
                usuario_id
            ),

            FOREIGN KEY (conversacion_id)
                REFERENCES conversaciones(id)
                ON DELETE CASCADE,

            FOREIGN KEY (usuario_id)
                REFERENCES usuarios(id)
                ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS mensajes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            conversacion_id INTEGER NOT NULL,
            usuario_id INTEGER NOT NULL,

            contenido TEXT,

            enviado_en TEXT NOT NULL
                DEFAULT CURRENT_TIMESTAMP,

            tipo TEXT NOT NULL
                DEFAULT 'texto',

            archivo_nombre TEXT,
            archivo_ruta TEXT,
            archivo_mime TEXT,

            leido INTEGER NOT NULL
                DEFAULT 0,

            FOREIGN KEY (conversacion_id)
                REFERENCES conversaciones(id)
                ON DELETE CASCADE,

            FOREIGN KEY (usuario_id)
                REFERENCES usuarios(id)
                ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS
            idx_mensajes_conv_id
        ON mensajes(conversacion_id, id);

        CREATE INDEX IF NOT EXISTS
            idx_participantes_usuario
        ON participantes_chat(usuario_id, conversacion_id);

        CREATE TABLE IF NOT EXISTS solicitudes_web (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_solicitud TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            cliente_potencial TEXT NOT NULL,
            servicio_interes TEXT,
            descripcion TEXT,
            correo TEXT,
            telefono TEXT,
            estado TEXT NOT NULL DEFAULT 'Pendiente',
            actualizado_en TEXT,
            procesado_por TEXT
        );

        CREATE INDEX IF NOT EXISTS idx_solicitudes_web_estado
        ON solicitudes_web(estado, id);
    """)

    # Migración para instalaciones que ya tenían solicitudes_web.
    columnas_sol = {
        row["name"]
        for row in conn.execute("PRAGMA table_info(solicitudes_web)").fetchall()
    }

    if "actualizado_en" not in columnas_sol:
        conn.execute(
            "ALTER TABLE solicitudes_web ADD COLUMN actualizado_en TEXT"
        )

    if "procesado_por" not in columnas_sol:
        conn.execute(
            "ALTER TABLE solicitudes_web ADD COLUMN procesado_por TEXT"
        )

    conn.commit()
    conn.close()


# ============================================================
# AUTENTICACIÓN
# ============================================================

def obtener_usuario(username, password):
    """
    Comprueba el usuario contra la tabla REAL usuarios
    de gestion_sistema.db.
    """

    username = (username or "").strip()
    password = password or ""

    if not username or not password:
        return None

    password_hash = hash_password(password)

    conn = db()

    usuario = conn.execute("""
        SELECT
            id,
            username,
            rol
        FROM usuarios
        WHERE username = ?
          AND password_hash = ?
        LIMIT 1
    """, (
        username,
        password_hash
    )).fetchone()

    conn.close()

    return usuario


def crear_sesion(usuario):
    token = secrets.token_urlsafe(32)

    SESIONES[token] = {
        "user_id": int(usuario["id"]),
        "username": usuario["username"],
        "rol": usuario["rol"]
    }

    return token


def obtener_sesion():
    """
    Obtiene el token enviado por Authorization:
    Bearer TOKEN
    """

    auth = request.headers.get("Authorization", "")

    if not auth.startswith("Bearer "):
        return None

    token = auth[7:].strip()

    if not token:
        return None

    return SESIONES.get(token)


def require_auth(fn):

    @wraps(fn)
    def wrapper(*args, **kwargs):

        # Permite la clave interna para pruebas
        # y comunicación servidor-servidor.
        chat_key = request.headers.get("X-Chat-Key", "")

        if API_KEY and chat_key == API_KEY:
            return fn(*args, **kwargs)

        sesion = obtener_sesion()

        if not sesion:
            return jsonify({
                "ok": False,
                "error": "Sesión no autorizada"
            }), 401

        request.chat_user = sesion

        return fn(*args, **kwargs)

    return wrapper


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return jsonify({
        "ok": True,
        "service": "chat-sgc"
    })


# ============================================================
# LOGIN DEL CHAT
# ============================================================

@app.post("/api/session")
def session():

    data = request.get_json(silent=True) or {}

    username = data.get("username", "")
    password = data.get("password", "")

    usuario = obtener_usuario(
        username,
        password
    )

    if not usuario:

        return jsonify({
            "ok": False,
            "error": "Usuario o contraseña incorrectos"
        }), 401

    token = crear_sesion(usuario)

    return jsonify({
        "ok": True,
        "token": token,
        "user_id": usuario["id"],
        "username": usuario["username"],
        "rol": usuario["rol"]
    })


# ============================================================
# USUARIOS
# ============================================================

@app.get("/api/users")
@require_auth
def users():

    sesion = obtener_sesion()

    if sesion:
        me_id = sesion["user_id"]
    else:
        me_id = int(
            request.args.get(
                "user_id",
                0
            )
        )

    search = request.args.get(
        "search",
        ""
    ).strip()

    conn = db()

    pattern = f"%{search}%"

    rows = conn.execute("""
        SELECT
            u.id,
            u.username,
            u.rol,

            (
                SELECT
                    CASE

                        WHEN COALESCE(
                            m.tipo,
                            'texto'
                        ) = 'texto'

                        THEN COALESCE(
                            m.contenido,
                            ''
                        )

                        WHEN m.tipo = 'imagen'

                        THEN '📷 ' ||
                             COALESCE(
                                m.archivo_nombre,
                                'Imagen'
                             )

                        WHEN m.tipo = 'video'

                        THEN '🎥 ' ||
                             COALESCE(
                                m.archivo_nombre,
                                'Video'
                             )

                        ELSE '📎 ' ||
                             COALESCE(
                                m.archivo_nombre,
                                'Archivo'
                             )

                    END

                FROM mensajes m

                JOIN participantes_chat p1
                    ON p1.conversacion_id =
                       m.conversacion_id

                JOIN participantes_chat p2
                    ON p2.conversacion_id =
                       m.conversacion_id

                WHERE p1.usuario_id = ?
                  AND p2.usuario_id = u.id

                ORDER BY m.id DESC

                LIMIT 1

            ) AS ultimo_mensaje,

            (
                SELECT COUNT(*)

                FROM mensajes m2

                JOIN participantes_chat p3
                    ON p3.conversacion_id =
                       m2.conversacion_id

                JOIN participantes_chat p4
                    ON p4.conversacion_id =
                       m2.conversacion_id

                WHERE p3.usuario_id = ?
                  AND p4.usuario_id = u.id

                  AND m2.usuario_id != ?

                  AND m2.leido = 0

            ) AS no_leidos

        FROM usuarios u

        WHERE u.id != ?

          AND u.username LIKE ?

        ORDER BY
            u.username COLLATE NOCASE

    """, (
        me_id,
        me_id,
        me_id,
        me_id,
        pattern
    )).fetchall()

    conn.close()

    return jsonify({
        "ok": True,
        "usuarios": [
            dict(row)
            for row in rows
        ]
    })


# ============================================================
# CONVERSACIONES
# ============================================================

@app.post("/api/conversations")
@require_auth
def conversations():

    data = request.get_json(
        silent=True
    ) or {}

    sesion = obtener_sesion()

    if sesion:
        me = sesion["user_id"]
    else:
        me = int(
            data.get(
                "user_id",
                0
            )
        )

    other = int(
        data.get(
            "target_user_id",
            0
        )
    )

    if (
        not me
        or not other
        or me == other
    ):

        return jsonify({
            "ok": False,
            "error": "Participantes inválidos"
        }), 400

    conn = db()

    row = conn.execute("""
        SELECT
            p1.conversacion_id

        FROM participantes_chat p1

        JOIN participantes_chat p2
            ON p1.conversacion_id =
               p2.conversacion_id

        WHERE p1.usuario_id = ?
          AND p2.usuario_id = ?

        LIMIT 1

    """, (
        me,
        other
    )).fetchone()

    if row:

        conv_id = row["conversacion_id"]

    else:

        cur = conn.execute(
            "INSERT INTO conversaciones DEFAULT VALUES"
        )

        conv_id = cur.lastrowid

        conn.executemany("""
            INSERT INTO participantes_chat(
                conversacion_id,
                usuario_id
            )
            VALUES (?, ?)
        """, [
            (
                conv_id,
                me
            ),
            (
                conv_id,
                other
            )
        ])

        conn.commit()

    conn.close()

    return jsonify({
        "ok": True,
        "conversacion_id": conv_id
    })


# ============================================================
# MENSAJES
# ============================================================

@app.get(
    "/api/conversations/<int:conv_id>/messages"
)
@require_auth
def messages(conv_id):

    sesion = obtener_sesion()

    if sesion:
        user_id = sesion["user_id"]
    else:
        user_id = int(
            request.args.get(
                "user_id",
                0
            )
        )

    conn = db()

    allowed = conn.execute("""
        SELECT 1

        FROM participantes_chat

        WHERE conversacion_id = ?
          AND usuario_id = ?

    """, (
        conv_id,
        user_id
    )).fetchone()

    if not allowed:

        conn.close()

        return jsonify({
            "ok": False,
            "error": "Sin acceso"
        }), 403

    rows = conn.execute("""
        SELECT
            m.id,
            m.usuario_id,
            u.username,
            m.contenido,
            m.enviado_en,
            m.tipo,
            m.archivo_nombre,
            m.archivo_ruta,
            m.archivo_mime

        FROM mensajes m

        JOIN usuarios u
            ON u.id = m.usuario_id

        WHERE m.conversacion_id = ?

        ORDER BY m.id ASC

    """, (
        conv_id,
    )).fetchall()

    conn.close()

    return jsonify({
        "ok": True,
        "mensajes": [
            dict(row)
            for row in rows
        ]
    })


@app.post("/api/messages")
@require_auth
def create_message():

    data = request.get_json(
        silent=True
    ) or {}

    sesion = obtener_sesion()

    if sesion:
        user_id = sesion["user_id"]
    else:
        user_id = int(
            data.get(
                "user_id",
                0
            )
        )

    conv_id = int(
        data.get(
            "conversacion_id",
            0
        )
    )

    content = (
        data.get(
            "contenido",
            ""
        )
        or ""
    )

    msg_type = (
        data.get(
            "tipo",
            "texto"
        )
        or "texto"
    )

    file_name = data.get(
        "archivo_nombre"
    )

    file_path = data.get(
        "archivo_ruta"
    )

    file_mime = data.get(
        "archivo_mime"
    )

    conn = db()

    allowed = conn.execute("""
        SELECT 1

        FROM participantes_chat

        WHERE conversacion_id = ?
          AND usuario_id = ?

    """, (
        conv_id,
        user_id
    )).fetchone()

    if not allowed:

        conn.close()

        return jsonify({
            "ok": False,
            "error": "Sin acceso"
        }), 403

    cur = conn.execute("""
        INSERT INTO mensajes(
            conversacion_id,
            usuario_id,
            contenido,
            tipo,
            archivo_nombre,
            archivo_ruta,
            archivo_mime,
            leido
        )

        VALUES (
            ?, ?, ?, ?, ?, ?, ?, 0
        )

    """, (
        conv_id,
        user_id,
        content,
        msg_type,
        file_name,
        file_path,
        file_mime
    ))

    conn.commit()

    message_id = cur.lastrowid

    conn.close()

    return jsonify({
        "ok": True,
        "message_id": message_id
    })


# ============================================================
# MARCAR MENSAJES COMO LEÍDOS
# ============================================================

@app.post("/api/messages/read")
@require_auth
def mark_read():

    data = request.get_json(
        silent=True
    ) or {}

    sesion = obtener_sesion()

    if sesion:
        user_id = sesion["user_id"]
    else:
        user_id = int(
            data.get(
                "user_id",
                0
            )
        )

    conv_id = int(
        data.get(
            "conversacion_id",
            0
        )
    )

    conn = db()

    allowed = conn.execute("""
        SELECT 1

        FROM participantes_chat

        WHERE conversacion_id = ?
          AND usuario_id = ?

    """, (
        conv_id,
        user_id
    )).fetchone()

    if not allowed:

        conn.close()

        return jsonify({
            "ok": False,
            "error": "Sin acceso"
        }), 403

    conn.execute("""
        UPDATE mensajes

        SET leido = 1

        WHERE conversacion_id = ?

          AND usuario_id != ?

    """, (
        conv_id,
        user_id
    ))

    conn.commit()

    conn.close()

    return jsonify({
        "ok": True
    })


# ============================================================
# NOTIFICACIONES
# ============================================================

@app.get("/api/notifications")
@require_auth
def notifications():

    sesion = obtener_sesion()

    if sesion:
        user_id = sesion["user_id"]
    else:
        user_id = int(
            request.args.get(
                "user_id",
                0
            )
        )

    after_id = int(
        request.args.get(
            "after_id",
            0
        )
    )

    conn = db()

    rows = conn.execute("""
        SELECT
            m.id,
            m.contenido,
            u.username,
            m.conversacion_id,
            m.tipo,
            m.archivo_nombre

        FROM mensajes m

        JOIN usuarios u
            ON u.id = m.usuario_id

        JOIN participantes_chat p
            ON p.conversacion_id =
               m.conversacion_id

        WHERE p.usuario_id = ?

          AND m.usuario_id != ?

          AND m.id > ?

          AND m.leido = 0

        ORDER BY m.id ASC

    """, (
        user_id,
        user_id,
        after_id
    )).fetchall()

    conn.close()

    return jsonify({
        "ok": True,
        "mensajes": [
            dict(row)
            for row in rows
        ]
    })


# ============================================================
# SOLICITUDES WEB (Landing Page -> Sistema de escritorio)
# ============================================================

def _cors_solicitudes(response):
    """CORS temporal para la landing durante la fase de pruebas."""
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PATCH, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Chat-Key"
    return response


@app.route("/api/solicitudes", methods=["OPTIONS"])
def solicitudes_options():
    return _cors_solicitudes(jsonify({"ok": True}))


@app.post("/api/solicitudes")
def crear_solicitud_web():
    """Endpoint público usado por el formulario del sitio web."""
    data = request.get_json(silent=True) or request.form.to_dict() or {}

    nombre = (data.get("nombre") or "").strip()
    apellido = (data.get("apellido") or "").strip()
    cliente = (data.get("cliente") or data.get("cliente_potencial") or f"{nombre} {apellido}").strip()
    servicio = (data.get("servicio") or data.get("servicio_interes") or "").strip()
    custom_servicio = (data.get("custom_servicio") or "").strip()
    if servicio.lower() == "otro":
        servicio = custom_servicio

    descripcion = (data.get("descripcion_negocio") or data.get("descripcion") or data.get("mensaje") or "").strip()
    correo = (data.get("correo") or data.get("email") or "").strip()
    telefono = (data.get("telefono") or data.get("telefono_contacto") or "").strip()

    if not cliente or not servicio or not descripcion or not correo or not telefono:
        return _cors_solicitudes(jsonify({
            "ok": False,
            "status": "error",
            "error": "Faltan datos obligatorios",
            "message": "Completa nombre, servicio, descripción, correo y teléfono."
        })), 400

    conn = db()
    cur = conn.execute("""
        INSERT INTO solicitudes_web
        (cliente_potencial, servicio_interes, descripcion, correo, telefono, estado)
        VALUES (?, ?, ?, ?, ?, 'Pendiente')
    """, (cliente, servicio, descripcion, correo, telefono))
    conn.commit()
    solicitud_id = cur.lastrowid
    conn.close()

    return _cors_solicitudes(jsonify({
        "ok": True,
        "status": "success",
        "id": solicitud_id,
        "message": "Solicitud recibida correctamente."
    })), 201


@app.get("/api/solicitudes")
@require_auth
def listar_solicitudes_web():
    estado = (request.args.get("estado") or "").strip()
    conn = db()
    if estado:
        rows = conn.execute("""
            SELECT id, fecha_solicitud, cliente_potencial, servicio_interes,
                   descripcion, correo, telefono, estado,
                   actualizado_en, procesado_por
            FROM solicitudes_web
            WHERE estado = ?
            ORDER BY id DESC
        """, (estado,)).fetchall()
    else:
        rows = conn.execute("""
            SELECT id, fecha_solicitud, cliente_potencial, servicio_interes,
                   descripcion, correo, telefono, estado,
                   actualizado_en, procesado_por
            FROM solicitudes_web
            ORDER BY id DESC
        """).fetchall()
    conn.close()
    return jsonify({"ok": True, "items": [dict(r) for r in rows]})


@app.patch("/api/solicitudes/<int:solicitud_id>")
@require_auth
def actualizar_solicitud_web(solicitud_id):
    data = request.get_json(silent=True) or {}
    estado = (data.get("estado") or "").strip()

    permitidos = {"Pendiente", "Contactado", "Aprobado", "Rechazado"}
    if estado not in permitidos:
        return jsonify({"ok": False, "error": "Estado inválido"}), 400

    sesion = getattr(request, "chat_user", None)
    usuario = "Sistema"
    if sesion:
        usuario = sesion.get("username") or "Sistema"

    conn = db()

    actual = conn.execute(
        "SELECT estado FROM solicitudes_web WHERE id=?",
        (solicitud_id,)
    ).fetchone()

    if not actual:
        conn.close()
        return jsonify({"ok": False, "error": "Solicitud no encontrada"}), 404

    estado_anterior = actual["estado"] or "Pendiente"

    # Una solicitud finalizada no puede volver a procesarse.
    if estado_anterior in {"Aprobado", "Rechazado"}:
        conn.close()
        return jsonify({
            "ok": False,
            "error": f"La solicitud ya fue {estado_anterior.lower()}"
        }), 409

    cur = conn.execute(
        """
        UPDATE solicitudes_web
        SET estado=?,
            actualizado_en=CURRENT_TIMESTAMP,
            procesado_por=?
        WHERE id=?
        """,
        (estado, usuario, solicitud_id)
    )

    conn.commit()
    actualizado = cur.rowcount

    fila = conn.execute(
        """
        SELECT id, estado, actualizado_en, procesado_por
        FROM solicitudes_web
        WHERE id=?
        """,
        (solicitud_id,)
    ).fetchone()

    conn.close()

    return jsonify({
        "ok": True,
        "id": solicitud_id,
        "estado": estado,
        "actualizado_en": fila["actualizado_en"] if fila else None,
        "procesado_por": fila["procesado_por"] if fila else usuario
    })


# ============================================================
# ARCHIVOS
# ============================================================

@app.post("/api/files")
@require_auth
def upload_file():

    file = request.files.get(
        "file"
    )

    if not file:

        return jsonify({
            "ok": False,
            "error": "Falta el archivo"
        }), 400

    original = secure_filename(
        file.filename or "archivo"
    )

    if not original:

        return jsonify({
            "ok": False,
            "error": "Nombre de archivo inválido"
        }), 400

    ext = Path(
        original
    ).suffix.lower()

    stored = (
        f"{uuid.uuid4().hex}"
        f"{ext}"
    )

    file.save(
        UPLOAD_DIR / stored
    )

    return jsonify({
        "ok": True,
        "archivo_nombre": original,
        "archivo_ruta": stored,
        "archivo_mime": (
            file.mimetype
            or "application/octet-stream"
        )
    })


@app.get("/api/files/<path:name>")
@require_auth
def download_file(name):

    return send_from_directory(
        UPLOAD_DIR,
        name,
        as_attachment=False
    )


# ============================================================
# INICIALIZACIÓN
# ============================================================

init_db()


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":

    port = int(
        os.getenv(
            "PORT",
            "5000"
        )
    )

    app.run(
        host="0.0.0.0",
        port=port,
        debug=False
    )