import os
import shutil
import mimetypes
import uuid
import subprocess
import platform
import tkinter as tk
from tkinter import messagebox, filedialog

from conexion import obtener_conexion


class VentanaChat(tk.Toplevel):

    # ============================================================
    # CONFIGURACIÓN
    # ============================================================

    MAX_ARCHIVO_MB = 100

    EXTENSIONES_IMAGEN = {
        ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"
    }

    EXTENSIONES_VIDEO = {
        ".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".m4v"
    }

    EXTENSIONES_DOCUMENTO = {
        ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv",
        ".txt", ".rtf", ".ppt", ".pptx", ".zip", ".rar",
        ".7z"
    }

    def __init__(self, parent, usuario_actual):

        super().__init__(parent)

        self.parent = parent
        self.usuario_actual = usuario_actual

        self.usuario_id = None
        self.conversacion_id = None
        self.usuario_destino_id = None
        self.usuario_destino = None

        self.colores = {
            "sidebar": "#6b1426",
            "vino": "#6b1426",
            "vino_oscuro": "#4f0e1c",
            "oro": "#d4af37",
            "fondo": "#f5f5f5",
            "blanco": "#ffffff",
            "gris": "#777777",
            "gris_claro": "#eeeeee",
            "texto": "#222222",
            "mensaje_mio": "#f1d7dc",
            "mensaje_otro": "#eeeeee",
            "verde": "#2a9d8f",
            "rojo": "#9e2a2b"
        }

        self.title("Chats - Sistema de Gestión de Costos")
        self.geometry("950x620")
        self.minsize(800, 520)

        self.protocol(
            "WM_DELETE_WINDOW",
            self.cerrar
        )

        # Carpeta raíz de archivos del chat.
        # Se crea al enviar el primer archivo.
        self.carpeta_archivos = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "archivos_chat"
        )

        self.obtener_usuario_actual()
        self.crear_interfaz()
        self.cargar_usuarios()

    # ============================================================
    # USUARIO ACTUAL
    # ============================================================

    def obtener_usuario_actual(self):

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                SELECT id, username
                FROM usuarios
                WHERE username = ?
                """,
                (self.usuario_actual,)
            )

            resultado = cursor.fetchone()

            conn.close()

            if resultado:

                self.usuario_id = resultado[0]
                self.usuario_actual = resultado[1]

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo obtener el usuario actual.\n\n{e}"
            )

    # ============================================================
    # INTERFAZ
    # ============================================================

    def crear_interfaz(self):

        self.configure(
            bg=self.colores["fondo"]
        )

        # --------------------------------------------------------
        # BARRA LATERAL
        # --------------------------------------------------------

        self.sidebar = tk.Frame(
            self,
            bg=self.colores["sidebar"],
            width=280
        )

        self.sidebar.pack(
            side="left",
            fill="y"
        )

        self.sidebar.pack_propagate(False)

        titulo = tk.Label(
            self.sidebar,
            text="💬 Chats",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 18, "bold")
        )

        titulo.pack(
            anchor="w",
            padx=20,
            pady=(20, 10)
        )

        self.entry_buscar = tk.Entry(
            self.sidebar,
            font=("Arial", 10),
            relief="flat",
            bg="white"
        )

        self.entry_buscar.pack(
            fill="x",
            padx=15,
            pady=(0, 10),
            ipady=7
        )

        self.entry_buscar.insert(
            0,
            "Buscar usuario..."
        )

        self.entry_buscar.bind(
            "<FocusIn>",
            self.limpiar_busqueda
        )

        self.entry_buscar.bind(
            "<KeyRelease>",
            lambda event: self.cargar_usuarios()
        )

        # Lista de usuarios
        frame_lista = tk.Frame(
            self.sidebar,
            bg=self.colores["sidebar"]
        )

        frame_lista.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        scrollbar = tk.Scrollbar(
            frame_lista
        )

        scrollbar.pack(
            side="right",
            fill="y"
        )

        self.lista_usuarios = tk.Listbox(
            frame_lista,
            bg=self.colores["sidebar"],
            fg="white",
            selectbackground=self.colores["oro"],
            selectforeground="black",
            font=("Arial", 11),
            relief="flat",
            bd=0,
            activestyle="none",
            yscrollcommand=scrollbar.set
        )

        self.lista_usuarios.pack(
            side="left",
            fill="both",
            expand=True
        )

        scrollbar.config(
            command=self.lista_usuarios.yview
        )

        self.lista_usuarios.bind(
            "<<ListboxSelect>>",
            self.usuario_seleccionado
        )

        # --------------------------------------------------------
        # PANEL PRINCIPAL
        # --------------------------------------------------------

        self.panel_chat = tk.Frame(
            self,
            bg=self.colores["fondo"]
        )

        self.panel_chat.pack(
            side="right",
            fill="both",
            expand=True
        )

        # Cabecera
        self.cabecera_chat = tk.Frame(
            self.panel_chat,
            bg=self.colores["blanco"],
            height=70
        )

        self.cabecera_chat.pack(
            fill="x"
        )

        self.cabecera_chat.pack_propagate(False)

        self.lbl_usuario_chat = tk.Label(
            self.cabecera_chat,
            text="Selecciona un usuario",
            bg=self.colores["blanco"],
            fg=self.colores["vino"],
            font=("Arial", 15, "bold")
        )

        self.lbl_usuario_chat.pack(
            anchor="w",
            padx=20,
            pady=(12, 0)
        )

        self.lbl_estado_chat = tk.Label(
            self.cabecera_chat,
            text="",
            bg=self.colores["blanco"],
            fg=self.colores["gris"],
            font=("Arial", 9)
        )

        self.lbl_estado_chat.pack(
            anchor="w",
            padx=20
        )

        # --------------------------------------------------------
        # ÁREA DE MENSAJES
        # --------------------------------------------------------

        self.frame_mensajes = tk.Frame(
            self.panel_chat,
            bg=self.colores["fondo"]
        )

        self.frame_mensajes.pack(
            fill="both",
            expand=True
        )

        self.canvas_mensajes = tk.Canvas(
            self.frame_mensajes,
            bg=self.colores["fondo"],
            highlightthickness=0
        )

        self.scroll_mensajes = tk.Scrollbar(
            self.frame_mensajes,
            orient="vertical",
            command=self.canvas_mensajes.yview
        )

        self.canvas_mensajes.configure(
            yscrollcommand=self.scroll_mensajes.set
        )

        self.scroll_mensajes.pack(
            side="right",
            fill="y"
        )

        self.canvas_mensajes.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.contenedor_mensajes = tk.Frame(
            self.canvas_mensajes,
            bg=self.colores["fondo"]
        )

        self.ventana_canvas = self.canvas_mensajes.create_window(
            (0, 0),
            window=self.contenedor_mensajes,
            anchor="nw"
        )

        self.contenedor_mensajes.bind(
            "<Configure>",
            self.actualizar_scroll
        )

        self.canvas_mensajes.bind(
            "<Configure>",
            self.ajustar_ancho_canvas
        )

        # --------------------------------------------------------
        # BARRA DE ESCRITURA
        # --------------------------------------------------------

        self.frame_escribir = tk.Frame(
            self.panel_chat,
            bg=self.colores["blanco"],
            height=65
        )

        self.frame_escribir.pack(
            fill="x"
        )

        self.frame_escribir.pack_propagate(False)

        # Botón adjuntar
        self.btn_adjuntar = tk.Button(
            self.frame_escribir,
            text="📎",
            bg=self.colores["blanco"],
            fg=self.colores["vino"],
            font=("Arial", 15, "bold"),
            relief="flat",
            cursor="hand2",
            command=self.seleccionar_archivo
        )

        self.btn_adjuntar.pack(
            side="left",
            padx=(12, 5),
            pady=10
        )

        self.entry_mensaje = tk.Entry(
            self.frame_escribir,
            font=("Arial", 11),
            relief="solid",
            bd=1
        )

        self.entry_mensaje.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(3, 8),
            pady=12,
            ipady=7
        )

        self.entry_mensaje.bind(
            "<Return>",
            lambda event: self.enviar_mensaje()
        )

        self.btn_enviar = tk.Button(
            self.frame_escribir,
            text="➤",
            bg=self.colores["vino"],
            fg="white",
            font=("Arial", 14, "bold"),
            relief="flat",
            width=4,
            cursor="hand2",
            command=self.enviar_mensaje
        )

        self.btn_enviar.pack(
            side="right",
            padx=(0, 15),
            pady=10
        )

        self.mostrar_mensaje_inicial()

    # ============================================================
    # BÚSQUEDA
    # ============================================================

    def limpiar_busqueda(self, event=None):

        if self.entry_buscar.get() == "Buscar usuario...":

            self.entry_buscar.delete(
                0,
                tk.END
            )

    # ============================================================
    # CARGAR USUARIOS
    # ============================================================

    def cargar_usuarios(self):

        if not self.usuario_id:
            return

        busqueda = self.entry_buscar.get().strip()

        if busqueda == "Buscar usuario...":
            busqueda = ""

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            consulta = """
                SELECT
                    u.id,
                    u.username,

                    (
                        SELECT
                            CASE
                                WHEN COALESCE(m.tipo, 'texto') = 'texto'
                                    THEN COALESCE(m.contenido, '')
                                WHEN m.tipo = 'imagen'
                                    THEN '📷 ' || COALESCE(m.archivo_nombre, 'Imagen')
                                WHEN m.tipo = 'video'
                                    THEN '🎥 ' || COALESCE(m.archivo_nombre, 'Video')
                                WHEN m.tipo = 'archivo'
                                    THEN '📎 ' || COALESCE(m.archivo_nombre, 'Archivo')
                                ELSE
                                    COALESCE(m.contenido, 'Archivo')
                            END
                        FROM mensajes m

                        INNER JOIN participantes_chat pc1
                            ON pc1.conversacion_id = m.conversacion_id

                        INNER JOIN participantes_chat pc2
                            ON pc2.conversacion_id = m.conversacion_id

                        WHERE pc1.usuario_id = ?
                          AND pc2.usuario_id = u.id

                        ORDER BY m.id DESC
                        LIMIT 1
                    ) AS ultimo_mensaje,

                    (
                        SELECT COUNT(*)
                        FROM mensajes m2

                        INNER JOIN participantes_chat pc3
                            ON pc3.conversacion_id = m2.conversacion_id

                        INNER JOIN participantes_chat pc4
                            ON pc4.conversacion_id = m2.conversacion_id

                        WHERE pc3.usuario_id = ?
                          AND pc4.usuario_id = u.id
                          AND m2.usuario_id != ?
                          AND m2.leido = 0
                    ) AS no_leidos

                FROM usuarios u

                WHERE u.id != ?
                  AND u.username LIKE ?

                ORDER BY u.username COLLATE NOCASE
            """

            parametro_busqueda = f"%{busqueda}%"

            cursor.execute(
                consulta,
                (
                    self.usuario_id,
                    self.usuario_id,
                    self.usuario_id,
                    self.usuario_id,
                    parametro_busqueda
                )
            )

            usuarios = cursor.fetchall()

            conn.close()

            self.lista_usuarios.delete(
                0,
                tk.END
            )

            self.mapa_usuarios = {}

            for usuario_id, username, ultimo, no_leidos in usuarios:

                if ultimo:

                    ultimo = ultimo.replace(
                        "\n",
                        " "
                    )

                    if len(ultimo) > 30:
                        ultimo = ultimo[:30] + "..."

                else:

                    ultimo = "Sin mensajes"

                if no_leidos and no_leidos > 0:

                    texto = (
                        f"● {username}\n"
                        f"   {ultimo}   🔴 {no_leidos}"
                    )

                else:

                    texto = (
                        f"   {username}\n"
                        f"   {ultimo}"
                    )

                self.lista_usuarios.insert(
                    tk.END,
                    texto
                )

                self.mapa_usuarios[
                    self.lista_usuarios.size() - 1
                ] = (
                    usuario_id,
                    username
                )

        except Exception as e:

            print(
                f"Error cargando usuarios del chat: {e}"
            )

    # ============================================================
    # SELECCIONAR USUARIO
    # ============================================================

    def usuario_seleccionado(self, event=None):

        seleccion = self.lista_usuarios.curselection()

        if not seleccion:
            return

        indice = seleccion[0]

        if indice not in self.mapa_usuarios:
            return

        usuario_id, username = self.mapa_usuarios[indice]

        self.usuario_destino_id = usuario_id
        self.usuario_destino = username

        self.lbl_usuario_chat.config(
            text=f"👤 {username}"
        )

        self.lbl_estado_chat.config(
            text="Chat interno"
        )

        self.obtener_o_crear_conversacion()

        self.cargar_mensajes()

        self.marcar_mensajes_leidos()

        self.cargar_usuarios()

        # Volver a dejar seleccionada la persona.
        for idx, datos in self.mapa_usuarios.items():

            if datos[0] == self.usuario_destino_id:

                self.lista_usuarios.selection_set(
                    idx
                )

                self.lista_usuarios.see(
                    idx
                )

                break

    # ============================================================
    # OBTENER / CREAR CONVERSACIÓN
    # ============================================================

    def obtener_o_crear_conversacion(self):

        if not self.usuario_id or not self.usuario_destino_id:
            return

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                SELECT pc1.conversacion_id
                FROM participantes_chat pc1
                INNER JOIN participantes_chat pc2
                    ON pc1.conversacion_id = pc2.conversacion_id
                WHERE pc1.usuario_id = ?
                  AND pc2.usuario_id = ?
                LIMIT 1
                """,
                (
                    self.usuario_id,
                    self.usuario_destino_id
                )
            )

            resultado = cursor.fetchone()

            if resultado:

                self.conversacion_id = resultado[0]

            else:

                cursor.execute(
                    """
                    INSERT INTO conversaciones
                    DEFAULT VALUES
                    """
                )

                self.conversacion_id = cursor.lastrowid

                cursor.execute(
                    """
                    INSERT INTO participantes_chat
                    (
                        conversacion_id,
                        usuario_id
                    )
                    VALUES (?, ?)
                    """,
                    (
                        self.conversacion_id,
                        self.usuario_id
                    )
                )

                cursor.execute(
                    """
                    INSERT INTO participantes_chat
                    (
                        conversacion_id,
                        usuario_id
                    )
                    VALUES (?, ?)
                    """,
                    (
                        self.conversacion_id,
                        self.usuario_destino_id
                    )
                )

                conn.commit()

            conn.close()

        except Exception as e:

            self.conversacion_id = None

            messagebox.showerror(
                "Error",
                f"No se pudo abrir la conversación.\n\n{e}"
            )

    # ============================================================
    # CARGAR MENSAJES
    # ============================================================

    def cargar_mensajes(self):

        for widget in self.contenedor_mensajes.winfo_children():

            widget.destroy()

        if not self.conversacion_id:
            return

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                SELECT
                    m.id,
                    m.usuario_id,
                    u.username,
                    m.contenido,
                    m.enviado_en,
                    COALESCE(m.tipo, 'texto'),
                    m.archivo_nombre,
                    m.archivo_ruta,
                    m.archivo_mime

                FROM mensajes m

                INNER JOIN usuarios u
                    ON u.id = m.usuario_id

                WHERE m.conversacion_id = ?

                ORDER BY m.id ASC
                """,
                (
                    self.conversacion_id,
                )
            )

            mensajes = cursor.fetchall()

            conn.close()

            if not mensajes:

                lbl = tk.Label(
                    self.contenedor_mensajes,
                    text=(
                        "No hay mensajes todavía.\n"
                        "Escribe el primer mensaje."
                    ),
                    bg=self.colores["fondo"],
                    fg=self.colores["gris"],
                    font=("Arial", 11)
                )

                lbl.pack(
                    pady=50
                )

            else:

                for mensaje in mensajes:

                    self.crear_burbuja_mensaje(
                        mensaje
                    )

            self.after(
                100,
                self.bajar_scroll
            )

        except Exception as e:

            print(
                f"Error cargando mensajes: {e}"
            )

    # ============================================================
    # BURBUJA DEL MENSAJE
    # ============================================================

    def crear_burbuja_mensaje(self, mensaje):

        mensaje_id = mensaje[0]
        usuario_id = mensaje[1]
        username = mensaje[2]
        contenido = mensaje[3]
        fecha = mensaje[4]
        tipo = mensaje[5] or "texto"
        archivo_nombre = mensaje[6]
        archivo_ruta = mensaje[7]
        archivo_mime = mensaje[8]

        es_mio = (
            usuario_id == self.usuario_id
        )

        fila = tk.Frame(
            self.contenedor_mensajes,
            bg=self.colores["fondo"]
        )

        fila.pack(
            fill="x",
            padx=15,
            pady=5
        )

        fondo = (
            self.colores["mensaje_mio"]
            if es_mio
            else self.colores["mensaje_otro"]
        )

        if es_mio:

            marco = tk.Frame(
                fila,
                bg=fondo,
                bd=0
            )

            marco.pack(
                side="right",
                padx=(80, 0)
            )

        else:

            marco = tk.Frame(
                fila,
                bg=fondo,
                bd=0
            )

            marco.pack(
                side="left",
                padx=(0, 80)
            )

        if not es_mio:

            lbl_nombre = tk.Label(
                marco,
                text=username,
                bg=fondo,
                fg=self.colores["vino"],
                font=("Arial", 9, "bold")
            )

            lbl_nombre.pack(
                anchor="w",
                padx=10,
                pady=(7, 0)
            )

        # --------------------------------------------------------
        # MENSAJE DE TEXTO
        # --------------------------------------------------------

        if tipo == "texto":

            lbl_mensaje = tk.Label(
                marco,
                text=contenido or "",
                bg=fondo,
                fg=self.colores["texto"],
                font=("Arial", 10),
                justify="left",
                anchor="w",
                wraplength=450
            )

            lbl_mensaje.pack(
                anchor="w",
                padx=10,
                pady=(6, 3)
            )

        # --------------------------------------------------------
        # MENSAJE CON ARCHIVO
        # --------------------------------------------------------

        else:

            self.crear_burbuja_archivo(
                marco,
                fondo,
                tipo,
                archivo_nombre,
                archivo_ruta,
                archivo_mime,
                contenido
            )

        lbl_fecha = tk.Label(
            marco,
            text=str(fecha),
            bg=fondo,
            fg=self.colores["gris"],
            font=("Arial", 8)
        )

        lbl_fecha.pack(
            anchor="e",
            padx=10,
            pady=(0, 7)
        )

    # ============================================================
    # BURBUJA DE ARCHIVO
    # ============================================================

    def crear_burbuja_archivo(
        self,
        marco,
        fondo,
        tipo,
        archivo_nombre,
        archivo_ruta,
        archivo_mime,
        contenido
    ):

        if tipo == "imagen":
            icono = "📷"
            etiqueta = "Imagen"

        elif tipo == "video":
            icono = "🎥"
            etiqueta = "Video"

        else:
            icono = "📎"
            etiqueta = "Archivo"

        nombre = (
            archivo_nombre
            or "Archivo adjunto"
        )

        if archivo_ruta:

            ruta_completa = self.obtener_ruta_archivo(
                archivo_ruta
            )

        else:

            ruta_completa = None

        tarjeta = tk.Frame(
            marco,
            bg=fondo
        )

        tarjeta.pack(
            fill="x",
            padx=8,
            pady=(7, 3)
        )

        tk.Label(
            tarjeta,
            text=icono,
            bg=fondo,
            fg=self.colores["vino"],
            font=("Arial", 22)
        ).pack(
            side="left",
            padx=(4, 8)
        )

        info = tk.Frame(
            tarjeta,
            bg=fondo
        )

        info.pack(
            side="left",
            fill="x",
            expand=True
        )

        tk.Label(
            info,
            text=etiqueta,
            bg=fondo,
            fg=self.colores["vino"],
            font=("Arial", 9, "bold")
        ).pack(
            anchor="w"
        )

        lbl_nombre = tk.Label(
            info,
            text=nombre,
            bg=fondo,
            fg=self.colores["texto"],
            font=("Arial", 10, "bold"),
            wraplength=330,
            justify="left"
        )

        lbl_nombre.pack(
            anchor="w"
        )

        # Mostrar tamaño si el archivo existe.
        if ruta_completa and os.path.exists(ruta_completa):

            try:

                tamano = os.path.getsize(
                    ruta_completa
                )

                tamano_texto = self.formatear_tamano(
                    tamano
                )

                tk.Label(
                    info,
                    text=tamano_texto,
                    bg=fondo,
                    fg=self.colores["gris"],
                    font=("Arial", 8)
                ).pack(
                    anchor="w"
                )

            except Exception:
                pass

            tk.Button(
                tarjeta,
                text="Abrir",
                bg=self.colores["vino"],
                fg="white",
                activebackground=self.colores["vino_oscuro"],
                activeforeground="white",
                relief="flat",
                cursor="hand2",
                command=lambda ruta=ruta_completa: (
                    self.abrir_archivo(ruta)
                )
            ).pack(
                side="right",
                padx=(5, 4)
            )

        else:

            tk.Label(
                tarjeta,
                text="Archivo no encontrado",
                bg=fondo,
                fg=self.colores["rojo"],
                font=("Arial", 8)
            ).pack(
                side="right",
                padx=5
            )

    # ============================================================
    # SELECCIONAR ARCHIVO
    # ============================================================

    def seleccionar_archivo(self):

        if not self.usuario_destino_id:

            messagebox.showwarning(
                "Chats",
                "Primero selecciona un usuario."
            )

            return

        tipos = [
            (
                "Todos los archivos",
                "*.*"
            ),
            (
                "Imágenes",
                "*.jpg *.jpeg *.png *.gif *.bmp *.webp"
            ),
            (
                "Videos",
                "*.mp4 *.mov *.avi *.mkv *.webm *.wmv *.m4v"
            ),
            (
                "Documentos",
                "*.pdf *.doc *.docx *.xls *.xlsx *.csv "
                "*.txt *.rtf *.ppt *.pptx *.zip *.rar *.7z"
            )
        ]

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo para enviar",
            filetypes=tipos
        )

        if not ruta:
            return

        self.enviar_archivo(
            ruta
        )

    # ============================================================
    # ENVIAR ARCHIVO
    # ============================================================

    def enviar_archivo(self, ruta_origen):

        if not os.path.isfile(ruta_origen):

            messagebox.showerror(
                "Archivo",
                "El archivo seleccionado no existe."
            )

            return

        try:

            tamano = os.path.getsize(
                ruta_origen
            )

            limite = (
                self.MAX_ARCHIVO_MB
                * 1024
                * 1024
            )

            if tamano > limite:

                messagebox.showwarning(
                    "Archivo demasiado grande",
                    (
                        f"El archivo pesa "
                        f"{self.formatear_tamano(tamano)}.\n\n"
                        f"El límite permitido es "
                        f"{self.MAX_ARCHIVO_MB} MB."
                    )
                )

                return

            if not self.conversacion_id:

                self.obtener_o_crear_conversacion()

            if not self.conversacion_id:
                return

            nombre_original = os.path.basename(
                ruta_origen
            )

            extension = (
                os.path.splitext(
                    nombre_original
                )[1].lower()
            )

            tipo = self.determinar_tipo_archivo(
                extension
            )

            mime = (
                mimetypes.guess_type(
                    nombre_original
                )[0]
                or "application/octet-stream"
            )

            # ----------------------------------------------------
            # Crear carpeta correspondiente
            # ----------------------------------------------------

            if tipo == "imagen":

                subcarpeta = "imagenes"

            elif tipo == "video":

                subcarpeta = "videos"

            else:

                subcarpeta = "documentos"

            carpeta_destino = os.path.join(
                self.carpeta_archivos,
                subcarpeta
            )

            os.makedirs(
                carpeta_destino,
                exist_ok=True
            )

            # ----------------------------------------------------
            # Nombre único para evitar colisiones
            # ----------------------------------------------------

            nombre_guardado = (
                f"{uuid.uuid4().hex}"
                f"{extension}"
            )

            ruta_destino = os.path.join(
                carpeta_destino,
                nombre_guardado
            )

            shutil.copy2(
                ruta_origen,
                ruta_destino
            )

            # Ruta relativa que guardamos en SQLite.
            ruta_relativa = os.path.relpath(
                ruta_destino,
                os.path.dirname(
                    os.path.abspath(__file__)
                )
            )

            # ----------------------------------------------------
            # Insertar mensaje
            # ----------------------------------------------------

            conn = obtener_conexion()
            cursor = conn.cursor()

            contenido = (
                f"📷 {nombre_original}"
                if tipo == "imagen"
                else
                f"🎥 {nombre_original}"
                if tipo == "video"
                else
                f"📎 {nombre_original}"
            )

            cursor.execute(
                """
                INSERT INTO mensajes
                (
                    conversacion_id,
                    usuario_id,
                    contenido,
                    tipo,
                    archivo_nombre,
                    archivo_ruta,
                    archivo_mime,
                    leido
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 0)
                """,
                (
                    self.conversacion_id,
                    self.usuario_id,
                    contenido,
                    tipo,
                    nombre_original,
                    ruta_relativa,
                    mime
                )
            )

            conn.commit()
            conn.close()

            self.cargar_mensajes()
            self.cargar_usuarios()

            # Restaurar selección.
            for indice, datos in self.mapa_usuarios.items():

                if datos[0] == self.usuario_destino_id:

                    self.lista_usuarios.selection_set(
                        indice
                    )

                    self.lista_usuarios.see(
                        indice
                    )

                    break

        except Exception as e:

            messagebox.showerror(
                "Error enviando archivo",
                f"No se pudo enviar el archivo.\n\n{e}"
            )

    # ============================================================
    # DETERMINAR TIPO
    # ============================================================

    def determinar_tipo_archivo(self, extension):

        if extension in self.EXTENSIONES_IMAGEN:
            return "imagen"

        if extension in self.EXTENSIONES_VIDEO:
            return "video"

        return "archivo"

    # ============================================================
    # RUTA DE ARCHIVO
    # ============================================================

    def obtener_ruta_archivo(self, ruta_guardada):

        if not ruta_guardada:
            return None

        base = os.path.dirname(
            os.path.abspath(__file__)
        )

        return os.path.abspath(
            os.path.join(
                base,
                ruta_guardada
            )
        )

    # ============================================================
    # ABRIR ARCHIVO
    # ============================================================

    def abrir_archivo(self, ruta):

        if not ruta:

            messagebox.showerror(
                "Archivo",
                "No se encontró la ruta del archivo."
            )

            return

        if not os.path.exists(ruta):

            messagebox.showerror(
                "Archivo",
                "El archivo ya no existe en el sistema."
            )

            return

        try:

            if platform.system() == "Windows":

                os.startfile(ruta)

            elif platform.system() == "Darwin":

                subprocess.Popen(
                    ["open", ruta]
                )

            else:

                subprocess.Popen(
                    ["xdg-open", ruta]
                )

        except Exception as e:

            messagebox.showerror(
                "Archivo",
                f"No se pudo abrir el archivo.\n\n{e}"
            )

    # ============================================================
    # FORMATEAR TAMAÑO
    # ============================================================

    def formatear_tamano(self, bytes_archivo):

        if bytes_archivo < 1024:

            return f"{bytes_archivo} B"

        if bytes_archivo < 1024 * 1024:

            return (
                f"{bytes_archivo / 1024:.1f} KB"
            )

        if bytes_archivo < 1024 * 1024 * 1024:

            return (
                f"{bytes_archivo / (1024 * 1024):.1f} MB"
            )

        return (
            f"{bytes_archivo / (1024 * 1024 * 1024):.1f} GB"
        )

    # ============================================================
    # ENVIAR MENSAJE DE TEXTO
    # ============================================================

    def enviar_mensaje(self):

        if not self.usuario_destino_id:

            messagebox.showwarning(
                "Chats",
                "Primero selecciona un usuario."
            )

            return

        contenido = (
            self.entry_mensaje.get().strip()
        )

        if not contenido:
            return

        if not self.conversacion_id:

            self.obtener_o_crear_conversacion()

        if not self.conversacion_id:
            return

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                INSERT INTO mensajes
                (
                    conversacion_id,
                    usuario_id,
                    contenido,
                    tipo,
                    leido
                )
                VALUES (?, ?, ?, 'texto', 0)
                """,
                (
                    self.conversacion_id,
                    self.usuario_id,
                    contenido
                )
            )

            conn.commit()
            conn.close()

            self.entry_mensaje.delete(
                0,
                tk.END
            )

            self.cargar_mensajes()
            self.cargar_usuarios()

            # Mantener seleccionado al destinatario.
            for indice, datos in self.mapa_usuarios.items():

                if datos[0] == self.usuario_destino_id:

                    self.lista_usuarios.selection_set(
                        indice
                    )

                    self.lista_usuarios.see(
                        indice
                    )

                    break

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo enviar el mensaje.\n\n{e}"
            )

    # ============================================================
    # MARCAR MENSAJES COMO LEÍDOS
    # ============================================================

    def marcar_mensajes_leidos(self):

        if not self.conversacion_id:
            return

        try:

            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                UPDATE mensajes
                SET leido = 1
                WHERE conversacion_id = ?
                  AND usuario_id != ?
                  AND leido = 0
                """,
                (
                    self.conversacion_id,
                    self.usuario_id
                )
            )

            cursor.execute(
                """
                UPDATE participantes_chat
                SET ultimo_mensaje_leido = COALESCE(
                    (
                        SELECT MAX(id)
                        FROM mensajes
                        WHERE conversacion_id = ?
                    ),
                    0
                )
                WHERE conversacion_id = ?
                  AND usuario_id = ?
                """,
                (
                    self.conversacion_id,
                    self.conversacion_id,
                    self.usuario_id
                )
            )

            conn.commit()
            conn.close()

        except Exception as e:

            print(
                f"Error marcando mensajes leídos: {e}"
            )

    # ============================================================
    # MENSAJE INICIAL
    # ============================================================

    def mostrar_mensaje_inicial(self):

        lbl = tk.Label(
            self.contenedor_mensajes,
            text=(
                "Selecciona un usuario "
                "para comenzar un chat."
            ),
            bg=self.colores["fondo"],
            fg=self.colores["gris"],
            font=("Arial", 12)
        )

        lbl.pack(
            pady=80
        )

    # ============================================================
    # SCROLL
    # ============================================================

    def actualizar_scroll(self, event=None):

        self.canvas_mensajes.configure(
            scrollregion=self.canvas_mensajes.bbox("all")
        )

    def ajustar_ancho_canvas(self, event):

        try:

            self.canvas_mensajes.itemconfig(
                self.ventana_canvas,
                width=event.width
            )

        except Exception:
            pass

    def bajar_scroll(self):

        try:

            self.canvas_mensajes.update_idletasks()

            self.canvas_mensajes.yview_moveto(
                1.0
            )

        except Exception:
            pass

    # ============================================================
    # CERRAR
    # ============================================================

    def cerrar(self):

        self.destroy()
