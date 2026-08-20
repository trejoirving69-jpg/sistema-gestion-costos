import os
import mimetypes
import uuid
import subprocess
import platform
import tkinter as tk
from tkinter import messagebox, filedialog

import requests
import certifi

from chat_config import CHAT_SERVER_URL


class VentanaChat(tk.Toplevel):
    MAX_ARCHIVO_MB = 100
    EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
    EXTENSIONES_VIDEO = {".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".m4v"}

    POLL_MS = 2500
    REFRESCO_USUARIOS_CADA = 4  # cada ~10 s con POLL_MS=2500

    def __init__(self, parent, usuario_actual, password_actual):
        super().__init__(parent)

        self.parent = parent
        self.usuario_actual = usuario_actual
        self.password_actual = password_actual

        self.usuario_id = None
        self.conversacion_id = None
        self.usuario_destino_id = None
        self.usuario_destino = None

        self.mapa_usuarios = {}
        self.snapshot_usuarios = None

        self.after_id = 0
        self.token = None
        self.sesion_activa = False
        self.poll_id = None
        self.ciclos_poll = 0

        self.ids_mensajes_renderizados = set()
        self.ultimo_id_conversacion = 0

        self.colores = {
            "vino": "#6b1426",
            "vino_oscuro": "#4a0d19",
            "vino_claro": "#8a2438",
            "oro": "#d4af37",
            "oro_suave": "#eadba7",
            "fondo": "#f6f2ea",
            "panel": "#ffffff",
            "panel_alt": "#fbf8f2",
            "texto": "#241d1f",
            "texto_suave": "#766b6d",
            "borde": "#ddd5c8",
            "mio": "#f3dde2",
            "otro": "#ffffff",
            "verde": "#2a9d6f",
            "rojo": "#b23a48",
            "gris_icono": "#8a7f81",
        }

        self.title("Chat Interno · Macilitano Consulting Group C.A.")
        self.geometry("1180x760")
        self.minsize(980, 620)
        self.configure(bg=self.colores["fondo"])

        try:
            self.transient(parent)
        except Exception:
            pass

        self.http = requests.Session()
        self.http.verify = certifi.where()
        self.headers = {}

        self.carpeta_archivos = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "archivos_chat",
        )

        self.protocol("WM_DELETE_WINDOW", self.cerrar)

        if not self.iniciar_sesion_red():
            self.after(50, self.cerrar)
            return

        self.crear_interfaz()
        self.cargar_usuarios(forzar=True)
        self.poll_id = self.after(self.POLL_MS, self.refrescar_red)

    # ============================================================
    # RED
    # ============================================================
    def api(self, method, path, **kwargs):
        if not self.sesion_activa or not self.token:
            raise RuntimeError("La sesión del chat no está iniciada.")

        headers = dict(self.headers)
        extra_headers = kwargs.pop("headers", None)
        if extra_headers:
            headers.update(extra_headers)

        response = self.http.request(
            method,
            f"{CHAT_SERVER_URL.rstrip('/')}{path}",
            headers=headers,
            timeout=15,
            verify=certifi.where(),
            **kwargs,
        )

        response.raise_for_status()
        data = response.json()

        if not data.get("ok"):
            raise RuntimeError(data.get("error", "Error del servidor"))

        return data

    def iniciar_sesion_red(self):
        if not CHAT_SERVER_URL.strip():
            messagebox.showerror(
                "Chat interno",
                "No se ha configurado el servidor del chat.",
                parent=self,
            )
            return False

        try:
            response = self.http.post(
                f"{CHAT_SERVER_URL.rstrip('/')}/api/session",
                json={
                    "username": self.usuario_actual,
                    "password": self.password_actual or "",
                },
                timeout=15,
                verify=certifi.where(),
            )
            response.raise_for_status()

            data = response.json()
            if not data.get("ok"):
                raise RuntimeError(
                    data.get("error", "Usuario o contraseña incorrectos")
                )

            self.usuario_id = int(data["user_id"])
            self.usuario_actual = data["username"]
            self.token = data["token"]
            self.headers = {"Authorization": f"Bearer {self.token}"}
            self.sesion_activa = True
            return True

        except requests.RequestException as e:
            messagebox.showerror(
                "Chat interno",
                f"No se pudo conectar con el servidor.\n\n{e}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Chat interno",
                f"No se pudo iniciar la sesión del chat.\n\n{e}",
                parent=self,
            )

        return False

    # ============================================================
    # INTERFAZ
    # ============================================================
    def crear_interfaz(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # ---------------- SIDEBAR ----------------
        self.sidebar = tk.Frame(
            self,
            bg=self.colores["vino_oscuro"],
            width=320,
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        marca = tk.Frame(self.sidebar, bg=self.colores["vino_oscuro"])
        marca.pack(fill="x", padx=20, pady=(22, 14))

        tk.Label(
            marca,
            text="MACILITANO",
            bg=self.colores["vino_oscuro"],
            fg="white",
            font=("Georgia", 15, "bold"),
        ).pack(anchor="w")

        tk.Label(
            marca,
            text="CHAT INTERNO",
            bg=self.colores["vino_oscuro"],
            fg=self.colores["oro"],
            font=("Arial", 9, "bold"),
        ).pack(anchor="w", pady=(2, 0))

        buscador_frame = tk.Frame(
            self.sidebar,
            bg=self.colores["panel"],
            highlightbackground=self.colores["oro"],
            highlightthickness=1,
        )
        buscador_frame.pack(fill="x", padx=16, pady=(0, 12))

        tk.Label(
            buscador_frame,
            text="⌕",
            bg=self.colores["panel"],
            fg=self.colores["gris_icono"],
            font=("Arial", 14),
        ).pack(side="left", padx=(10, 4))

        self.entry_buscar = tk.Entry(
            buscador_frame,
            font=("Arial", 10),
            relief="flat",
            bd=0,
            bg=self.colores["panel"],
            fg=self.colores["texto"],
            insertbackground=self.colores["vino"],
        )
        self.entry_buscar.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 8),
            pady=9,
        )
        self.entry_buscar.insert(0, "Buscar usuario...")
        self.entry_buscar.bind("<FocusIn>", self.limpiar_busqueda)
        self.entry_buscar.bind("<FocusOut>", self.restaurar_busqueda)
        self.entry_buscar.bind("<KeyRelease>", lambda e: self.cargar_usuarios(forzar=True))

        encabezado_conversaciones = tk.Frame(
            self.sidebar,
            bg=self.colores["vino_oscuro"]
        )
        encabezado_conversaciones.pack(fill="x", padx=17, pady=(0, 5))

        tk.Label(
            encabezado_conversaciones,
            text="Conversaciones",
            bg=self.colores["vino_oscuro"],
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(side="left")

        self.lbl_conexion = tk.Label(
            encabezado_conversaciones,
            text="● En línea",
            bg=self.colores["vino_oscuro"],
            fg="#77d6a7",
            font=("Arial", 8, "bold"),
        )
        self.lbl_conexion.pack(side="right")

        lista_frame = tk.Frame(
            self.sidebar,
            bg=self.colores["vino_oscuro"],
        )
        lista_frame.pack(fill="both", expand=True, padx=10, pady=(0, 12))

        scroll_usuarios = tk.Scrollbar(
            lista_frame,
            bd=0,
            relief="flat",
        )
        scroll_usuarios.pack(side="right", fill="y")

        self.lista_usuarios = tk.Listbox(
            lista_frame,
            bg=self.colores["vino_oscuro"],
            fg="white",
            selectbackground=self.colores["oro"],
            selectforeground=self.colores["vino_oscuro"],
            font=("Arial", 10),
            relief="flat",
            bd=0,
            highlightthickness=0,
            activestyle="none",
            yscrollcommand=scroll_usuarios.set,
            selectborderwidth=0,
        )
        self.lista_usuarios.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(5, 0),
        )
        scroll_usuarios.config(command=self.lista_usuarios.yview)
        self.lista_usuarios.bind("<<ListboxSelect>>", self.usuario_seleccionado)

        usuario_footer = tk.Frame(
            self.sidebar,
            bg=self.colores["vino"],
            height=62,
        )
        usuario_footer.pack(fill="x", side="bottom")
        usuario_footer.pack_propagate(False)

        avatar = tk.Label(
            usuario_footer,
            text=self._inicial_usuario(self.usuario_actual),
            bg=self.colores["oro"],
            fg=self.colores["vino_oscuro"],
            width=3,
            height=1,
            font=("Arial", 12, "bold"),
        )
        avatar.pack(side="left", padx=(16, 10), pady=13)

        info_user = tk.Frame(usuario_footer, bg=self.colores["vino"])
        info_user.pack(side="left", fill="y", pady=10)

        tk.Label(
            info_user,
            text=self.usuario_actual,
            bg=self.colores["vino"],
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(anchor="w")

        tk.Label(
            info_user,
            text="Sesión activa",
            bg=self.colores["vino"],
            fg=self.colores["oro_suave"],
            font=("Arial", 8),
        ).pack(anchor="w")

        # ---------------- PANEL CHAT ----------------
        self.panel_chat = tk.Frame(
            self,
            bg=self.colores["fondo"],
        )
        self.panel_chat.grid(row=0, column=1, sticky="nsew")
        self.panel_chat.grid_rowconfigure(1, weight=1)
        self.panel_chat.grid_columnconfigure(0, weight=1)

        # Header conversación
        head = tk.Frame(
            self.panel_chat,
            bg=self.colores["panel"],
            height=78,
            highlightbackground=self.colores["borde"],
            highlightthickness=1,
        )
        head.grid(row=0, column=0, sticky="ew")
        head.grid_propagate(False)

        self.avatar_chat = tk.Label(
            head,
            text="?",
            bg=self.colores["vino"],
            fg="white",
            font=("Arial", 13, "bold"),
            width=3,
        )
        self.avatar_chat.pack(side="left", padx=(22, 12), pady=17)

        info_chat = tk.Frame(head, bg=self.colores["panel"])
        info_chat.pack(side="left", fill="y", pady=13)

        self.lbl_usuario_chat = tk.Label(
            info_chat,
            text="Selecciona una conversación",
            bg=self.colores["panel"],
            fg=self.colores["vino"],
            font=("Arial", 14, "bold"),
        )
        self.lbl_usuario_chat.pack(anchor="w")

        self.lbl_estado_chat = tk.Label(
            info_chat,
            text="Elige un usuario de la lista para comenzar",
            bg=self.colores["panel"],
            fg=self.colores["texto_suave"],
            font=("Arial", 9),
        )
        self.lbl_estado_chat.pack(anchor="w", pady=(2, 0))

        # Área mensajes
        fm = tk.Frame(
            self.panel_chat,
            bg=self.colores["fondo"],
        )
        fm.grid(row=1, column=0, sticky="nsew")
        fm.grid_rowconfigure(0, weight=1)
        fm.grid_columnconfigure(0, weight=1)

        self.canvas_mensajes = tk.Canvas(
            fm,
            bg=self.colores["fondo"],
            highlightthickness=0,
            bd=0,
        )
        self.canvas_mensajes.grid(row=0, column=0, sticky="nsew")

        sbm = tk.Scrollbar(
            fm,
            orient="vertical",
            command=self.canvas_mensajes.yview,
            bd=0,
            relief="flat",
        )
        sbm.grid(row=0, column=1, sticky="ns")

        self.canvas_mensajes.configure(yscrollcommand=sbm.set)

        self.contenedor_mensajes = tk.Frame(
            self.canvas_mensajes,
            bg=self.colores["fondo"],
        )

        self.ventana_canvas = self.canvas_mensajes.create_window(
            (0, 0),
            window=self.contenedor_mensajes,
            anchor="nw",
        )

        self.contenedor_mensajes.bind("<Configure>", self.actualizar_scroll)
        self.canvas_mensajes.bind("<Configure>", self.ajustar_ancho_canvas)
        self.canvas_mensajes.bind_all("<MouseWheel>", self._rueda_mouse)

        # Composer
        composer = tk.Frame(
            self.panel_chat,
            bg=self.colores["panel"],
            height=82,
            highlightbackground=self.colores["borde"],
            highlightthickness=1,
        )
        composer.grid(row=2, column=0, sticky="ew")
        composer.grid_propagate(False)

        btn_adjuntar = tk.Button(
            composer,
            text="📎",
            bg=self.colores["panel"],
            fg=self.colores["vino"],
            font=("Arial", 16),
            relief="flat",
            bd=0,
            cursor="hand2",
            activebackground=self.colores["panel_alt"],
            command=self.seleccionar_archivo,
        )
        btn_adjuntar.pack(side="left", padx=(18, 8), pady=17)

        entrada_frame = tk.Frame(
            composer,
            bg=self.colores["panel_alt"],
            highlightbackground=self.colores["borde"],
            highlightthickness=1,
        )
        entrada_frame.pack(
            side="left",
            fill="x",
            expand=True,
            pady=16,
        )

        self.entry_mensaje = tk.Entry(
            entrada_frame,
            font=("Arial", 11),
            relief="flat",
            bd=0,
            bg=self.colores["panel_alt"],
            fg=self.colores["texto"],
            insertbackground=self.colores["vino"],
        )
        self.entry_mensaje.pack(
            fill="x",
            expand=True,
            padx=13,
            pady=10,
        )
        self.entry_mensaje.bind("<Return>", lambda event: self.enviar_mensaje())

        btn_enviar = tk.Button(
            composer,
            text="Enviar  ➤",
            bg=self.colores["vino"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            bd=0,
            padx=18,
            pady=9,
            cursor="hand2",
            activebackground=self.colores["vino_claro"],
            activeforeground="white",
            command=self.enviar_mensaje,
        )
        btn_enviar.pack(side="right", padx=(10, 20), pady=17)

        self.mostrar_mensaje_inicial()

    # ============================================================
    # USUARIOS
    # ============================================================
    def _inicial_usuario(self, nombre):
        nombre = (nombre or "?").strip()
        return nombre[:1].upper() if nombre else "?"

    def limpiar_busqueda(self, event=None):
        if self.entry_buscar.get() == "Buscar usuario...":
            self.entry_buscar.delete(0, tk.END)
            self.entry_buscar.config(fg=self.colores["texto"])

    def restaurar_busqueda(self, event=None):
        if not self.entry_buscar.get().strip():
            self.entry_buscar.insert(0, "Buscar usuario...")
            self.entry_buscar.config(fg=self.colores["texto_suave"])

    def cargar_usuarios(self, forzar=False):
        if not self.usuario_id or not self.sesion_activa:
            return

        try:
            query = self.entry_buscar.get().strip()
            if query == "Buscar usuario...":
                query = ""

            data = self.api(
                "GET",
                "/api/users",
                params={"search": query},
            )

            usuarios = data.get("usuarios", [])

            snapshot = tuple(
                (
                    int(user["id"]),
                    user["username"],
                    user.get("ultimo_mensaje") or "",
                    int(user.get("no_leidos") or 0),
                )
                for user in usuarios
            )

            # Si nada cambió, NO reconstruimos la lista.
            if not forzar and snapshot == self.snapshot_usuarios:
                return

            seleccionado_id = self.usuario_destino_id

            self.lista_usuarios.delete(0, tk.END)
            self.mapa_usuarios = {}

            indice_a_reseleccionar = None

            for user in usuarios:
                last = (user.get("ultimo_mensaje") or "Sin mensajes").replace("\n", " ")
                if len(last) > 34:
                    last = last[:34] + "…"

                unread = int(user.get("no_leidos") or 0)
                nombre = user["username"]

                if unread:
                    texto = f"●  {nombre}   ({unread})\n    {last}"
                else:
                    texto = f"   {nombre}\n    {last}"

                self.lista_usuarios.insert(tk.END, texto)

                index = self.lista_usuarios.size() - 1
                user_id = int(user["id"])
                self.mapa_usuarios[index] = (user_id, nombre)

                if seleccionado_id and user_id == seleccionado_id:
                    indice_a_reseleccionar = index

            if indice_a_reseleccionar is not None:
                self.lista_usuarios.selection_set(indice_a_reseleccionar)
                self.lista_usuarios.activate(indice_a_reseleccionar)
                self.lista_usuarios.see(indice_a_reseleccionar)

            self.snapshot_usuarios = snapshot
            self.lbl_conexion.config(text="● En línea", fg="#77d6a7")

        except Exception as e:
            self.lbl_conexion.config(text="● Sin conexión", fg="#ff9b9b")
            print("Error cargando usuarios:", e)

    def usuario_seleccionado(self, event=None):
        selection = self.lista_usuarios.curselection()
        if not selection:
            return

        index = selection[0]
        if index not in self.mapa_usuarios:
            return

        nuevo_id, nuevo_nombre = self.mapa_usuarios[index]

        # Evita recargar la conversación si la selección no cambió.
        if (
            self.usuario_destino_id == nuevo_id
            and self.conversacion_id is not None
        ):
            return

        self.usuario_destino_id = nuevo_id
        self.usuario_destino = nuevo_nombre

        self.lbl_usuario_chat.config(text=self.usuario_destino)
        self.lbl_estado_chat.config(text="Conversación interna segura")
        self.avatar_chat.config(text=self._inicial_usuario(self.usuario_destino))

        try:
            self.obtener_conversacion()
            self.cargar_mensajes_completos()
            self.marcar_leidos()
            self.cargar_usuarios(forzar=True)
            self.entry_mensaje.focus_set()

        except Exception as e:
            messagebox.showerror(
                "Chat interno",
                f"No se pudo abrir la conversación.\n\n{e}",
                parent=self,
            )

    def obtener_conversacion(self):
        data = self.api(
            "POST",
            "/api/conversations",
            json={"target_user_id": self.usuario_destino_id},
        )

        nueva_conv = int(data["conversacion_id"])

        if nueva_conv != self.conversacion_id:
            self.conversacion_id = nueva_conv
            self.ids_mensajes_renderizados.clear()
            self.ultimo_id_conversacion = 0

    # ============================================================
    # MENSAJES
    # ============================================================
    def _limpiar_area_mensajes(self):
        for widget in self.contenedor_mensajes.winfo_children():
            widget.destroy()

    def cargar_mensajes_completos(self):
        if not self.conversacion_id:
            return

        try:
            data = self.api(
                "GET",
                f"/api/conversations/{self.conversacion_id}/messages",
            )
            mensajes = data.get("mensajes", [])

            self._limpiar_area_mensajes()
            self.ids_mensajes_renderizados.clear()
            self.ultimo_id_conversacion = 0

            if not mensajes:
                self.mostrar_conversacion_vacia()
                return

            for mensaje in mensajes:
                self._renderizar_si_nuevo(mensaje)

            self.after(70, self.bajar_scroll)

        except Exception as e:
            print("Error cargando mensajes:", e)

    def actualizar_mensajes_incremental(self):
        """Consulta la conversación pero solo inserta mensajes que aún no están pintados."""
        if not self.conversacion_id:
            return False

        try:
            data = self.api(
                "GET",
                f"/api/conversations/{self.conversacion_id}/messages",
            )
            mensajes = data.get("mensajes", [])

            nuevos = [
                m for m in mensajes
                if int(m.get("id") or 0) not in self.ids_mensajes_renderizados
            ]

            if not nuevos:
                return False

            # Si estaba visible el texto de conversación vacía, quitarlo.
            if not self.ids_mensajes_renderizados:
                self._limpiar_area_mensajes()

            estaba_abajo = self._esta_scroll_abajo()

            for mensaje in nuevos:
                self._renderizar_si_nuevo(mensaje)

            self.marcar_leidos()

            # Solo baja automáticamente si el usuario ya estaba abajo.
            if estaba_abajo:
                self.after(50, self.bajar_scroll)

            return True

        except Exception as e:
            print("Error actualizando conversación:", e)
            return False

    def _renderizar_si_nuevo(self, mensaje):
        message_id = int(mensaje.get("id") or 0)

        if message_id and message_id in self.ids_mensajes_renderizados:
            return

        self.crear_burbuja(mensaje)

        if message_id:
            self.ids_mensajes_renderizados.add(message_id)
            self.ultimo_id_conversacion = max(
                self.ultimo_id_conversacion,
                message_id,
            )

    def crear_burbuja(self, message):
        fila = tk.Frame(
            self.contenedor_mensajes,
            bg=self.colores["fondo"],
        )
        fila.pack(fill="x", padx=22, pady=5)

        mio = int(message["usuario_id"]) == self.usuario_id
        fondo = self.colores["mio"] if mio else self.colores["otro"]

        # "Tarjeta" de mensaje
        marco = tk.Frame(
            fila,
            bg=fondo,
            highlightbackground=(
                self.colores["vino_claro"] if mio else self.colores["borde"]
            ),
            highlightthickness=1,
            bd=0,
        )
        marco.pack(
            side="right" if mio else "left",
            padx=(150, 0) if mio else (0, 150),
        )

        if not mio:
            tk.Label(
                marco,
                text=message.get("username") or "",
                bg=fondo,
                fg=self.colores["vino"],
                font=("Arial", 9, "bold"),
            ).pack(
                anchor="w",
                padx=12,
                pady=(8, 0),
            )

        tipo = message.get("tipo", "texto")

        if tipo == "texto":
            tk.Label(
                marco,
                text=message.get("contenido") or "",
                bg=fondo,
                fg=self.colores["texto"],
                font=("Arial", 10),
                justify="left",
                anchor="w",
                wraplength=520,
            ).pack(
                anchor="w",
                padx=12,
                pady=(7, 4),
            )
        else:
            self.crear_archivo(
                marco,
                fondo,
                message,
            )

        fecha = self._formatear_fecha(message.get("enviado_en"))

        tk.Label(
            marco,
            text=fecha,
            bg=fondo,
            fg=self.colores["texto_suave"],
            font=("Arial", 8),
        ).pack(
            anchor="e",
            padx=12,
            pady=(0, 8),
        )

    def _formatear_fecha(self, valor):
        if not valor:
            return ""
        texto = str(valor)
        # Conserva compatibilidad con el formato que devuelve el servidor.
        if len(texto) >= 16:
            return texto[:16].replace("T", " ")
        return texto

    def crear_archivo(self, marco, fondo, message):
        tipo = message.get("tipo")

        if tipo == "imagen":
            icono = "▣"
            etiqueta = "Imagen"
        elif tipo == "video":
            icono = "▶"
            etiqueta = "Video"
        else:
            icono = "▤"
            etiqueta = "Archivo"

        card = tk.Frame(
            marco,
            bg=fondo,
        )
        card.pack(fill="x", padx=10, pady=(9, 4))

        icon = tk.Label(
            card,
            text=icono,
            bg=self.colores["vino"],
            fg="white",
            width=3,
            font=("Arial", 14, "bold"),
        )
        icon.pack(side="left", padx=(0, 10), ipady=5)

        info = tk.Frame(card, bg=fondo)
        info.pack(side="left", fill="x", expand=True)

        tk.Label(
            info,
            text=etiqueta,
            bg=fondo,
            fg=self.colores["vino"],
            font=("Arial", 8, "bold"),
        ).pack(anchor="w")

        tk.Label(
            info,
            text=message.get("archivo_nombre") or "Archivo adjunto",
            bg=fondo,
            fg=self.colores["texto"],
            font=("Arial", 9, "bold"),
            wraplength=350,
            justify="left",
        ).pack(anchor="w")

        tk.Button(
            card,
            text="Abrir",
            bg=self.colores["vino"],
            fg="white",
            relief="flat",
            bd=0,
            padx=12,
            pady=5,
            cursor="hand2",
            command=lambda: self.descargar_archivo(
                message.get("archivo_ruta"),
                message.get("archivo_nombre"),
            ),
        ).pack(side="right", padx=(10, 0))

    def mostrar_mensaje_inicial(self):
        self._limpiar_area_mensajes()

        bloque = tk.Frame(
            self.contenedor_mensajes,
            bg=self.colores["fondo"],
        )
        bloque.pack(expand=True, pady=140)

        tk.Label(
            bloque,
            text="💬",
            bg=self.colores["fondo"],
            fg=self.colores["vino"],
            font=("Arial", 32),
        ).pack()

        tk.Label(
            bloque,
            text="Chat interno",
            bg=self.colores["fondo"],
            fg=self.colores["vino"],
            font=("Arial", 16, "bold"),
        ).pack(pady=(5, 3))

        tk.Label(
            bloque,
            text="Selecciona una conversación para comenzar.",
            bg=self.colores["fondo"],
            fg=self.colores["texto_suave"],
            font=("Arial", 10),
        ).pack()

    def mostrar_conversacion_vacia(self):
        self._limpiar_area_mensajes()

        bloque = tk.Frame(
            self.contenedor_mensajes,
            bg=self.colores["fondo"],
        )
        bloque.pack(expand=True, pady=120)

        tk.Label(
            bloque,
            text="Aún no hay mensajes",
            bg=self.colores["fondo"],
            fg=self.colores["vino"],
            font=("Arial", 14, "bold"),
        ).pack()

        tk.Label(
            bloque,
            text="Escribe el primer mensaje para iniciar la conversación.",
            bg=self.colores["fondo"],
            fg=self.colores["texto_suave"],
            font=("Arial", 10),
        ).pack(pady=(4, 0))

    def enviar_mensaje(self):
        if not self.usuario_destino_id:
            messagebox.showwarning(
                "Chat interno",
                "Primero selecciona una conversación.",
                parent=self,
            )
            return

        contenido = self.entry_mensaje.get().strip()
        if not contenido:
            return

        try:
            if not self.conversacion_id:
                self.obtener_conversacion()

            data = self.api(
                "POST",
                "/api/messages",
                json={
                    "conversacion_id": self.conversacion_id,
                    "contenido": contenido,
                    "tipo": "texto",
                },
            )

            self.entry_mensaje.delete(0, tk.END)

            # Actualiza sin reconstruir toda la pantalla.
            self.actualizar_mensajes_incremental()
            self.cargar_usuarios(forzar=True)
            self.after(40, self.bajar_scroll)

        except Exception as e:
            messagebox.showerror(
                "Chat interno",
                f"No se pudo enviar el mensaje.\n\n{e}",
                parent=self,
            )

    # ============================================================
    # ARCHIVOS
    # ============================================================
    def seleccionar_archivo(self):
        if not self.usuario_destino_id:
            messagebox.showwarning(
                "Chat interno",
                "Primero selecciona una conversación.",
                parent=self,
            )
            return

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[
                ("Todos los archivos", "*.*"),
                ("Imágenes", "*.jpg *.jpeg *.png *.gif *.bmp *.webp"),
                ("Videos", "*.mp4 *.mov *.avi *.mkv *.webm *.wmv *.m4v"),
                (
                    "Documentos",
                    "*.pdf *.doc *.docx *.xls *.xlsx *.csv "
                    "*.txt *.rtf *.ppt *.pptx *.zip *.rar *.7z",
                ),
            ],
        )

        if ruta:
            self.enviar_archivo(ruta)

    def enviar_archivo(self, ruta):
        if os.path.getsize(ruta) > self.MAX_ARCHIVO_MB * 1024 * 1024:
            messagebox.showwarning(
                "Archivo demasiado grande",
                f"El archivo supera el límite de {self.MAX_ARCHIVO_MB} MB.",
                parent=self,
            )
            return

        try:
            if not self.conversacion_id:
                self.obtener_conversacion()

            nombre = os.path.basename(ruta)
            ext = os.path.splitext(nombre)[1].lower()

            if ext in self.EXTENSIONES_IMAGEN:
                tipo = "imagen"
            elif ext in self.EXTENSIONES_VIDEO:
                tipo = "video"
            else:
                tipo = "archivo"

            mime = mimetypes.guess_type(nombre)[0] or "application/octet-stream"

            with open(ruta, "rb") as archivo:
                data = self.api(
                    "POST",
                    "/api/files",
                    files={"file": (nombre, archivo, mime)},
                )

            if tipo == "imagen":
                contenido = f"Imagen: {nombre}"
            elif tipo == "video":
                contenido = f"Video: {nombre}"
            else:
                contenido = f"Archivo: {nombre}"

            self.api(
                "POST",
                "/api/messages",
                json={
                    "conversacion_id": self.conversacion_id,
                    "contenido": contenido,
                    "tipo": tipo,
                    "archivo_nombre": data["archivo_nombre"],
                    "archivo_ruta": data["archivo_ruta"],
                    "archivo_mime": data["archivo_mime"],
                },
            )

            self.actualizar_mensajes_incremental()
            self.cargar_usuarios(forzar=True)
            self.after(40, self.bajar_scroll)

        except Exception as e:
            messagebox.showerror(
                "Archivo",
                f"No se pudo enviar el archivo.\n\n{e}",
                parent=self,
            )

    def descargar_archivo(self, ruta, nombre):
        if not ruta:
            return

        try:
            carpeta = os.path.join(
                self.carpeta_archivos,
                "recibidos",
            )
            os.makedirs(carpeta, exist_ok=True)

            destino = os.path.join(
                carpeta,
                f"{uuid.uuid4().hex}_{os.path.basename(nombre or 'archivo')}",
            )

            response = self.http.get(
                f"{CHAT_SERVER_URL.rstrip('/')}/api/files/{ruta}",
                headers=self.headers,
                timeout=120,
                verify=certifi.where(),
            )
            response.raise_for_status()

            with open(destino, "wb") as archivo:
                archivo.write(response.content)

            self.abrir_archivo(destino)

        except Exception as e:
            messagebox.showerror(
                "Archivo",
                f"No se pudo abrir el archivo recibido.\n\n{e}",
                parent=self,
            )

    def abrir_archivo(self, ruta):
        try:
            if platform.system() == "Windows":
                os.startfile(ruta)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", ruta])
            else:
                subprocess.Popen(["xdg-open", ruta])
        except Exception as e:
            messagebox.showerror(
                "Archivo",
                str(e),
                parent=self,
            )

    def marcar_leidos(self):
        if not self.conversacion_id:
            return

        try:
            self.api(
                "POST",
                "/api/messages/read",
                json={"conversacion_id": self.conversacion_id},
            )
        except Exception:
            pass

    # ============================================================
    # REFRESCO SIN PARPADEO
    # ============================================================
    def refrescar_red(self):
        if not self.sesion_activa:
            return

        try:
            data = self.api(
                "GET",
                "/api/notifications",
                params={"after_id": self.after_id},
            )

            nuevos_globales = data.get("mensajes", [])

            for message in nuevos_globales:
                self.after_id = max(
                    self.after_id,
                    int(message.get("id") or 0),
                )

            # Solo tocamos la conversación si hay algo realmente nuevo.
            if nuevos_globales and self.conversacion_id:
                self.actualizar_mensajes_incremental()

            # La lista lateral ya no se reconstruye cada 2 segundos.
            self.ciclos_poll += 1
            if nuevos_globales or self.ciclos_poll >= self.REFRESCO_USUARIOS_CADA:
                self.ciclos_poll = 0
                self.cargar_usuarios(forzar=False)

            self.lbl_conexion.config(text="● En línea", fg="#77d6a7")

        except Exception as e:
            self.lbl_conexion.config(text="● Reconectando…", fg="#f0c36e")
            print("Monitor chat:", e)

        try:
            self.poll_id = self.after(
                self.POLL_MS,
                self.refrescar_red,
            )
        except Exception:
            self.poll_id = None

    # ============================================================
    # SCROLL / CIERRE
    # ============================================================
    def actualizar_scroll(self, event=None):
        bbox = self.canvas_mensajes.bbox("all")
        if bbox:
            self.canvas_mensajes.configure(scrollregion=bbox)

    def ajustar_ancho_canvas(self, event):
        self.canvas_mensajes.itemconfig(
            self.ventana_canvas,
            width=event.width,
        )

    def _esta_scroll_abajo(self):
        try:
            primero, ultimo = self.canvas_mensajes.yview()
            return ultimo >= 0.94
        except Exception:
            return True

    def bajar_scroll(self):
        try:
            self.canvas_mensajes.update_idletasks()
            self.canvas_mensajes.yview_moveto(1.0)
        except Exception:
            pass

    def _rueda_mouse(self, event):
        try:
            if self.winfo_exists():
                self.canvas_mensajes.yview_scroll(
                    int(-1 * (event.delta / 120)),
                    "units",
                )
        except Exception:
            pass

    def cerrar(self):
        self.sesion_activa = False

        if self.poll_id:
            try:
                self.after_cancel(self.poll_id)
            except Exception:
                pass
            self.poll_id = None

        try:
            self.canvas_mensajes.unbind_all("<MouseWheel>")
        except Exception:
            pass

        try:
            self.http.close()
        except Exception:
            pass

        try:
            self.destroy()
        except Exception:
            pass
