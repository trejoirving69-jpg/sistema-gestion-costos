import sqlite3
import os
import sys
import hashlib

def generar_id_cliente(tipo_industria, conexion_db):
    # Mapeo simplificado a prefijos según tipo de industria (1,2,3)
    industrias = {
        "manufacturera": "1",
        "manufactura": "1",
        "comercial": "2",
        "retail": "2",
        "servicios": "3",
        "servicio": "3",
    }

    tipo_lower = tipo_industria.lower() if tipo_industria else ""
    prefix = "0"
    # Buscar por subcadena para permitir valores como 'Industria manufacturera'
    for key, val in industrias.items():
        if key in tipo_lower:
            prefix = val
            break

    # Consultar cuántos clientes existen con ese mismo prefijo para el correlativo
    cursor = conexion_db.cursor()
    cursor.execute("SELECT COUNT(*) FROM clientes WHERE codigo LIKE ?", (f"{prefix}-%",))
    contador = cursor.fetchone()[0] + 1

    # Formato final: 1-001, 2-005, 3-012
    nuevo_id = f"{prefix}-{contador:03d}"
    return nuevo_id

def obtener_app_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def obtener_ruta_recurso(nombre_archivo):
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nombre_archivo)

def obtener_conexion():
    ruta_bd = os.path.join(obtener_app_path(), "gestion_sistema.db")

    conexion = sqlite3.connect(ruta_bd)
    conexion.execute("PRAGMA foreign_keys = ON")
    return conexion

def hash_password(pw: str) -> str:
    """Devuelve el hash SHA-256 hex de la contraseña proporcionada."""
    return hashlib.sha256(pw.encode('utf-8')).hexdigest()

def inicializar_base_datos():
    """Inicializa la base de datos local creando todas las tablas necesarias."""
    conexion = obtener_conexion()
    cursor = conexion.cursor()
    
    # 1. Tabla Usuarios para el login
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        rol TEXT NOT NULL
    )
    """)

    # 2. Tabla Clientes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL,
        correo TEXT,
        telefono TEXT,
        fecha_registro DATE,
        codigo TEXT,
        industria TEXT,
        servicio TEXT
    )
    """)

    # Migraciones: asegurarse de que columnas usadas por la UI existan
    cursor.execute("PRAGMA table_info(clientes)")
    existing_cols = [row[1] for row in cursor.fetchall()]
    if 'nombre_contacto' not in existing_cols:
        cursor.execute("ALTER TABLE clientes ADD COLUMN nombre_contacto TEXT")
    if 'estado' not in existing_cols:
        cursor.execute("ALTER TABLE clientes ADD COLUMN estado TEXT DEFAULT 'Activo'")
    # Consolidar y eliminar columna duplicada 'persona_contacto' si existe
    if 'persona_contacto' in existing_cols:
        # Copiar valores no vacíos a 'nombre_contacto' cuando falte
        try:
            cursor.execute("UPDATE clientes SET nombre_contacto = persona_contacto WHERE (nombre_contacto IS NULL OR nombre_contacto = '') AND (persona_contacto IS NOT NULL AND persona_contacto != '')")
            # Recreate table without persona_contacto (SQLite no soporta DROP COLUMN directo)
            # Definir nueva estructura explícita (sin persona_contacto)
            cursor.execute("CREATE TABLE IF NOT EXISTS clientes_new (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT NOT NULL, correo TEXT, telefono TEXT, fecha_registro DATE, codigo TEXT, industria TEXT, servicio TEXT, nombre_contacto TEXT, estado TEXT DEFAULT 'Activo')")
            # Copiar datos seleccionando las columnas relevantes
            cursor.execute("INSERT INTO clientes_new (id, nombre, correo, telefono, fecha_registro, codigo, industria, servicio, nombre_contacto, estado) SELECT id, nombre, correo, telefono, fecha_registro, codigo, industria, servicio, nombre_contacto, COALESCE(estado, 'Activo') FROM clientes")
            cursor.execute("DROP TABLE clientes")
            cursor.execute("ALTER TABLE clientes_new RENAME TO clientes")
        except Exception:
            # Si algo falla, no interrumpir la inicialización; dejar ambas columnas
            pass

    # 2b. Tabla Asesores (vinculada a un usuario de login)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS asesores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario_id INTEGER UNIQUE,
        nombre TEXT NOT NULL,
        correo TEXT,
        telefono TEXT,
        especialidad TEXT,
        fecha_contratacion DATE,
        salario REAL DEFAULT 0.0,
        activo TEXT DEFAULT 'Activo',
        FOREIGN KEY(usuario_id) REFERENCES usuarios(id)
    )
    """)

    cursor.execute("PRAGMA table_info(asesores)")
    asesores_cols = [row[1] for row in cursor.fetchall()]
    if 'fecha_contratacion' not in asesores_cols:
        cursor.execute("ALTER TABLE asesores ADD COLUMN fecha_contratacion DATE")
    if 'salario' not in asesores_cols:
        cursor.execute("ALTER TABLE asesores ADD COLUMN salario REAL DEFAULT 0.0")
    
    # 3. Tabla Servicios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS servicios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre_servicio TEXT NOT NULL,
        descripcion TEXT,
        costo_base REAL NOT NULL
    )
    """)
    
    # Asegurar que los servicios predeterminados existan (insertar solo los faltantes)
    servicios_iniciales = [
        (
            "Planificación Estratégica Empresarial",
            "Diseño de mapas de ruta corporativos, proyecciones financieras y modelos de negocio orientados al crecimiento sostenible y a la resiliencia en mercados dinámicos.",
            150.0
        ),
        (
            "Gestión Financiera, Costos y Tesorería",
            "Reestructuración y control de costos, auditorías de inventario, ingeniería de flujos de caja, presupuestos y optimización del capital de trabajo (cuentas por cobrar/pagar).",
            180.0
        ),
        (
            "Optimización de Procesos Administrativos y Operacionales",
            "Diagnóstico, rediseño e implementación de manuales de procedimientos para maximizar la eficiencia interna y reducir las brechas operativas.",
            140.0
        ),
        (
            "Transformación Digital (Sistema Homologado)",
            "Acompañamiento en la migración tecnológica e integración de sistemas de información alineados a las exigencias regulatorias y operativas actuales.",
            220.0
        ),
        (
            "Outsourcing Contable y Gestión de Nómina",
            "Externalización profesional de la contabilidad general, estados financieros confiables y gestión técnica de recursos humanos.",
            200.0
        ),
        (
            "Cumplimiento Fiscal, Tributario y Parafiscal",
            "Asesoría experta y ejecución de declaraciones fiscales y parafiscales ante los organismos correspondientes (SENIAT, IVSS, INCES, FAOV, entre otros).",
            190.0
        ),
        (
            "Consultoría Legal Corporativa",
            "Soporte integral en derecho empresarial y societario para la protección de activos y la correcta gobernanza de la organización.",
            170.0
        ),
    ]

    for nombre, desc, costo in servicios_iniciales:
        cursor.execute("SELECT COUNT(*) FROM servicios WHERE nombre_servicio = ?", (nombre,))
        if cursor.fetchone()[0] == 0:
            cursor.execute(
                "INSERT INTO servicios (nombre_servicio, descripcion, costo_base) VALUES (?, ?, ?)",
                (nombre, desc, costo)
            )
    # 4. Tabla de solicitudes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS solicitudes_servicio (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER NOT NULL,
        servicio_id INTEGER NOT NULL,
        fecha_solicitud DATE,
        monto REAL,                 
        iva_usd REAL DEFAULT 0.0,   
        igtf_usd REAL DEFAULT 0.0,  
        total_usd REAL DEFAULT 0.0, 
        tasa_bcv REAL DEFAULT 1.0,  
        total_bs REAL DEFAULT 0.0,  
        metodo_pago TEXT,           
        activo TEXT DEFAULT 'Activo',
        FOREIGN KEY(cliente_id) REFERENCES clientes(id),
        FOREIGN KEY(servicio_id) REFERENCES servicios(id)
    )
    """)

    # 4b. Tabla para solicitudes web (formularios desde Landing Page)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS solicitudes_web (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha_solicitud TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        cliente_potencial TEXT NOT NULL,
        servicio_interes TEXT,
        descripcion TEXT,
        correo TEXT,
        telefono TEXT,
        estado TEXT DEFAULT 'Pendiente'
    )
    """)

    # 5. Tabla de notas internas entre usuarios
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS notas_sistema (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        remitente TEXT NOT NULL,
        destinatario TEXT NOT NULL,
        titulo TEXT NOT NULL,
        mensaje TEXT NOT NULL,
        tipo TEXT DEFAULT 'nota',
        servicio_id INTEGER,
        leida INTEGER DEFAULT 0,
        fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
        # ============================================================
    # 6. SISTEMA DE CHAT INTERNO
    # ============================================================

    # Conversaciones
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS conversaciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        creada_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Usuarios que participan en cada conversación
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS participantes_chat (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        conversacion_id INTEGER NOT NULL,
        usuario_id INTEGER NOT NULL,
        ultimo_mensaje_leido INTEGER DEFAULT 0,

        FOREIGN KEY(conversacion_id)
            REFERENCES conversaciones(id)
            ON DELETE CASCADE,

        FOREIGN KEY(usuario_id)
            REFERENCES usuarios(id)
            ON DELETE CASCADE,

        UNIQUE(conversacion_id, usuario_id)
    )
    """)

    # Mensajes
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS mensajes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,

        conversacion_id INTEGER NOT NULL,
        usuario_id INTEGER NOT NULL,

        contenido TEXT,

        tipo TEXT DEFAULT 'texto',

        archivo_nombre TEXT,
        archivo_ruta TEXT,
        archivo_mime TEXT,

        leido INTEGER DEFAULT 0,

        enviado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

        FOREIGN KEY(conversacion_id)
            REFERENCES conversaciones(id)
            ON DELETE CASCADE,

        FOREIGN KEY(usuario_id)
            REFERENCES usuarios(id)
            ON DELETE CASCADE
    )
    """)

    # 6. Tabla Historial de Actividades (Auditoría)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS historial_actividades (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        usuario TEXT NOT NULL,
        accion TEXT NOT NULL,
        detalles TEXT,
        fecha_hora TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Asegurar un usuario administrador inicial
    cursor.execute("SELECT id FROM usuarios WHERE username = ?", ("admin",))
    if cursor.fetchone() is None:
        # Usar hash para la contraseña por defecto
        admin_pw_hash = hash_password("arepaconqueso")
        cursor.execute(
            "INSERT INTO usuarios (username, password_hash, rol) VALUES (?, ?, ?)",
            ("admin", admin_pw_hash, "Administrador")
        )
    
    conexion.commit()
    conexion.close()
    print("Base de datos estructurada e inicializada con éxito.")

def registrar_accion(usuario, accion, detalles=""):
    """Inserta de forma segura una traza de auditoría en el sistema."""
    try:
        conn = obtener_conexion()
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO historial_actividades (usuario, accion, detalles)
            VALUES (?, ?, ?)
        """, (usuario, accion, detalles))
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"Error al guardar bitácora de auditoría: {e}")