import os
import mimetypes
import uuid
import subprocess
import platform
import tkinter as tk
from tkinter import messagebox, filedialog

import requests

from chat_config import CHAT_SERVER_URL


class VentanaChat(tk.Toplevel):
    MAX_ARCHIVO_MB = 100
    EXTENSIONES_IMAGEN = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
    EXTENSIONES_VIDEO = {".mp4", ".mov", ".avi", ".mkv", ".webm", ".wmv", ".m4v"}

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
        self.after_id = 0
        self.token = None
        self.sesion_activa = False

        self.colores = {
            "sidebar": "#6b1426",
            "vino": "#6b1426",
            "vino_oscuro": "#4f0e1c",
            "oro": "#d4af37",
            "fondo": "#f5f5f5",
            "blanco": "#fff",
            "gris": "#777",
            "texto": "#222",
            "mio": "#f1d7dc",
            "otro": "#eee",
            "verde": "#2a9d8f",
            "rojo": "#9e2a2b",
        }

        self.title("Chats - Sistema de Gestión de Costos")
        self.geometry("950x620")
        self.minsize(800, 520)

        self.http = requests.Session()
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
        self.cargar_usuarios()
        self.after(2000, self.refrescar_red)

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
                "Chat en red",
                "CHAT_SERVER_URL está vacío en chat_config.py",
                parent=self,
            )
            return False

        try:
            # /api/session es el único endpoint que no necesita
            # token: aquí comprobamos las credenciales del sistema.
            response = self.http.post(
                f"{CHAT_SERVER_URL.rstrip('/')}/api/session",
                json={
                    "username": self.usuario_actual,
                    "password": self.password_actual or "",
                },
                timeout=15,
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

            self.headers = {
                "Authorization": f"Bearer {self.token}"
            }

            self.sesion_activa = True
            return True

        except requests.RequestException as e:
            messagebox.showerror(
                "Chat en red",
                f"No se pudo conectar al servidor.\n\n{e}",
                parent=self,
            )
        except Exception as e:
            messagebox.showerror(
                "Chat en red",
                f"No se pudo iniciar la sesión del chat.\n\n{e}",
                parent=self,
            )

        return False

    def crear_interfaz(self):
        self.configure(bg=self.colores["fondo"])

        self.sidebar = tk.Frame(
            self,
            bg=self.colores["sidebar"],
            width=280,
        )
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        tk.Label(
            self.sidebar,
            text="💬 Chats",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 18, "bold"),
        ).pack(anchor="w", padx=20, pady=(20, 10))

        self.entry_buscar = tk.Entry(
            self.sidebar,
            font=("Arial", 10),
            relief="flat",
            bg="white",
        )
        self.entry_buscar.pack(
            fill="x",
            padx=15,
            pady=(0, 10),
            ipady=7,
        )
        self.entry_buscar.insert(0, "Buscar usuario...")
        self.entry_buscar.bind("<FocusIn>", self.limpiar_busqueda)
        self.entry_buscar.bind(
            "<KeyRelease>",
            lambda event: self.cargar_usuarios(),
        )

        fr = tk.Frame(self.sidebar, bg=self.colores["sidebar"])
        fr.pack(fill="both", expand=True, padx=10, pady=5)

        sb = tk.Scrollbar(fr)
        sb.pack(side="right", fill="y")

        self.lista_usuarios = tk.Listbox(
            fr,
            bg=self.colores["sidebar"],
            fg="white",
            selectbackground=self.colores["oro"],
            selectforeground="black",
            font=("Arial", 11),
            relief="flat",
            bd=0,
            activestyle="none",
            yscrollcommand=sb.set,
        )
        self.lista_usuarios.pack(side="left", fill="both", expand=True)
        sb.config(command=self.lista_usuarios.yview)
        self.lista_usuarios.bind(
            "<<ListboxSelect>>",
            self.usuario_seleccionado,
        )

        self.panel_chat = tk.Frame(
            self,
            bg=self.colores["fondo"],
        )
        self.panel_chat.pack(side="right", fill="both", expand=True)

        head = tk.Frame(
            self.panel_chat,
            bg="white",
            height=70,
        )
        head.pack(fill="x")
        head.pack_propagate(False)

        self.lbl_usuario_chat = tk.Label(
            head,
            text="Selecciona un usuario",
            bg="white",
            fg=self.colores["vino"],
            font=("Arial", 15, "bold"),
        )
        self.lbl_usuario_chat.pack(
            anchor="w",
            padx=20,
            pady=(12, 0),
        )

        self.lbl_estado_chat = tk.Label(
            head,
            text="",
            bg="white",
            fg=self.colores["gris"],
            font=("Arial", 9),
        )
        self.lbl_estado_chat.pack(anchor="w", padx=20)

        fm = tk.Frame(
            self.panel_chat,
            bg=self.colores["fondo"],
        )
        fm.pack(fill="both", expand=True)

        self.canvas_mensajes = tk.Canvas(
            fm,
            bg=self.colores["fondo"],
            highlightthickness=0,
        )

        sbm = tk.Scrollbar(
            fm,
            orient="vertical",
            command=self.canvas_mensajes.yview,
        )
        self.canvas_mensajes.configure(
            yscrollcommand=sbm.set
        )
        sbm.pack(side="right", fill="y")
        self.canvas_mensajes.pack(
            side="left",
            fill="both",
            expand=True,
        )

        self.contenedor_mensajes = tk.Frame(
            self.canvas_mensajes,
            bg=self.colores["fondo"],
        )

        self.ventana_canvas = self.canvas_mensajes.create_window(
            (0, 0),
            window=self.contenedor_mensajes,
            anchor="nw",
        )

        self.contenedor_mensajes.bind(
            "<Configure>",
            self.actualizar_scroll,
        )
        self.canvas_mensajes.bind(
            "<Configure>",
            self.ajustar_ancho_canvas,
        )

        fw = tk.Frame(
            self.panel_chat,
            bg="white",
            height=65,
        )
        fw.pack(fill="x")
        fw.pack_propagate(False)

        tk.Button(
            fw,
            text="📎",
            bg="white",
            fg=self.colores["vino"],
            font=("Arial", 15, "bold"),
            relief="flat",
            cursor="hand2",
            command=self.seleccionar_archivo,
        ).pack(
            side="left",
            padx=(12, 5),
            pady=10,
        )

        self.entry_mensaje = tk.Entry(
            fw,
            font=("Arial", 11),
            relief="solid",
            bd=1,
        )
        self.entry_mensaje.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(3, 8),
            pady=12,
            ipady=7,
        )
        self.entry_mensaje.bind(
            "<Return>",
            lambda event: self.enviar_mensaje(),
        )

        tk.Button(
            fw,
            text="➤",
            bg=self.colores["vino"],
            fg="white",
            font=("Arial", 14, "bold"),
            relief="flat",
            width=4,
            cursor="hand2",
            command=self.enviar_mensaje,
        ).pack(
            side="right",
            padx=(0, 15),
            pady=10,
        )

        self.mostrar_mensaje_inicial()

    def limpiar_busqueda(self, event=None):
        if self.entry_buscar.get() == "Buscar usuario...":
            self.entry_buscar.delete(0, tk.END)

    def cargar_usuarios(self):
        if not self.usuario_id or not self.sesion_activa:
            return

        try:
            query = self.entry_buscar.get().strip()
            if query == "Buscar usuario...":
                query = ""

            data = self.api(
                "GET",
                "/api/users",
                params={
                    "search": query,
                },
            )

            self.lista_usuarios.delete(0, tk.END)
            self.mapa_usuarios = {}

            for user in data["usuarios"]:
                last = (
                    user.get("ultimo_mensaje") or "Sin mensajes"
                ).replace("\n", " ")

                if len(last) > 30:
                    last = last[:30] + "..."

                unread = int(user.get("no_leidos") or 0)

                if unread:
                    text = (
                        f"● {user['username']}\n"
                        f"   {last}   🔴 {unread}"
                    )
                else:
                    text = (
                        f"   {user['username']}\n"
                        f"   {last}"
                    )

                self.lista_usuarios.insert(tk.END, text)

                index = self.lista_usuarios.size() - 1
                self.mapa_usuarios[index] = (
                    int(user["id"]),
                    user["username"],
                )

        except Exception as e:
            print("Error cargando usuarios:", e)

    def usuario_seleccionado(self, event=None):
        selection = self.lista_usuarios.curselection()

        if not selection:
            return

        index = selection[0]

        if index not in self.mapa_usuarios:
            return

        (
            self.usuario_destino_id,
            self.usuario_destino,
        ) = self.mapa_usuarios[index]

        self.lbl_usuario_chat.config(
            text=f"👤 {self.usuario_destino}"
        )
        self.lbl_estado_chat.config(
            text="Chat por red"
        )

        try:
            self.obtener_conversacion()
            self.cargar_mensajes()
            self.marcar_leidos()
        except Exception as e:
            messagebox.showerror(
                "Chats",
                f"No se pudo abrir la conversación.\n\n{e}",
                parent=self,
            )

    def obtener_conversacion(self):
        data = self.api(
            "POST",
            "/api/conversations",
            json={
                "target_user_id": self.usuario_destino_id,
            },
        )

        self.conversacion_id = int(
            data["conversacion_id"]
        )

    def cargar_mensajes(self):
        for widget in self.contenedor_mensajes.winfo_children():
            widget.destroy()

        if not self.conversacion_id:
            return

        try:
            data = self.api(
                "GET",
                f"/api/conversations/{self.conversacion_id}/messages",
            )

            messages = data["mensajes"]

            if not messages:
                tk.Label(
                    self.contenedor_mensajes,
                    text=(
                        "No hay mensajes todavía.\n"
                        "Escribe el primer mensaje."
                    ),
                    bg=self.colores["fondo"],
                    fg=self.colores["gris"],
                    font=("Arial", 11),
                ).pack(pady=50)

            for message in messages:
                self.crear_burbuja(message)

            self.after(100, self.bajar_scroll)

        except Exception as e:
            print("Error cargando mensajes:", e)

    def crear_burbuja(self, message):
        fila = tk.Frame(
            self.contenedor_mensajes,
            bg=self.colores["fondo"],
        )
        fila.pack(
            fill="x",
            padx=15,
            pady=5,
        )

        mio = (
            int(message["usuario_id"])
            == self.usuario_id
        )

        fondo = (
            self.colores["mio"]
            if mio
            else self.colores["otro"]
        )

        marco = tk.Frame(
            fila,
            bg=fondo,
        )
        marco.pack(
            side="right" if mio else "left",
            padx=(80, 0) if mio else (0, 80),
        )

        if not mio:
            tk.Label(
                marco,
                text=message["username"],
                bg=fondo,
                fg=self.colores["vino"],
                font=("Arial", 9, "bold"),
            ).pack(
                anchor="w",
                padx=10,
                pady=(7, 0),
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
                wraplength=450,
            ).pack(
                anchor="w",
                padx=10,
                pady=(6, 3),
            )
        else:
            self.crear_archivo(
                marco,
                fondo,
                message,
            )

        tk.Label(
            marco,
            text=str(message.get("enviado_en", "")),
            bg=fondo,
            fg=self.colores["gris"],
            font=("Arial", 8),
        ).pack(
            anchor="e",
            padx=10,
            pady=(0, 7),
        )

    def crear_archivo(self, marco, fondo, message):
        tipo = message.get("tipo")
        icon = (
            "📷"
            if tipo == "imagen"
            else "🎥"
            if tipo == "video"
            else "📎"
        )

        etiqueta = (
            "Imagen"
            if tipo == "imagen"
            else "Video"
            if tipo == "video"
            else "Archivo"
        )

        card = tk.Frame(marco, bg=fondo)
        card.pack(
            fill="x",
            padx=8,
            pady=(7, 3),
        )

        tk.Label(
            card,
            text=icon,
            bg=fondo,
            fg=self.colores["vino"],
            font=("Arial", 22),
        ).pack(
            side="left",
            padx=(4, 8),
        )

        info = tk.Frame(card, bg=fondo)
        info.pack(
            side="left",
            fill="x",
            expand=True,
        )

        tk.Label(
            info,
            text=etiqueta,
            bg=fondo,
            fg=self.colores["vino"],
            font=("Arial", 9, "bold"),
        ).pack(anchor="w")

        tk.Label(
            info,
            text=message.get("archivo_nombre")
            or "Archivo adjunto",
            bg=fondo,
            fg=self.colores["texto"],
            font=("Arial", 10, "bold"),
            wraplength=330,
            justify="left",
        ).pack(anchor="w")

        tk.Button(
            card,
            text="Abrir",
            bg=self.colores["vino"],
            fg="white",
            relief="flat",
            cursor="hand2",
            command=lambda: self.descargar_archivo(
                message.get("archivo_ruta"),
                message.get("archivo_nombre"),
            ),
        ).pack(
            side="right",
            padx=4,
        )

    def enviar_mensaje(self):
        if not self.usuario_destino_id:
            messagebox.showwarning(
                "Chats",
                "Primero selecciona un usuario.",
                parent=self,
            )
            return

        text = self.entry_mensaje.get().strip()

        if not text:
            return

        try:
            if not self.conversacion_id:
                self.obtener_conversacion()

            self.api(
                "POST",
                "/api/messages",
                json={
                    "conversacion_id": self.conversacion_id,
                    "contenido": text,
                    "tipo": "texto",
                },
            )

            self.entry_mensaje.delete(0, tk.END)
            self.cargar_mensajes()
            self.cargar_usuarios()

        except Exception as e:
            messagebox.showerror(
                "Chat",
                f"No se pudo enviar el mensaje.\n\n{e}",
                parent=self,
            )

    def seleccionar_archivo(self):
        if not self.usuario_destino_id:
            messagebox.showwarning(
                "Chats",
                "Primero selecciona un usuario.",
                parent=self,
            )
            return

        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo para enviar",
            filetypes=[
                ("Todos los archivos", "*.*"),
                (
                    "Imágenes",
                    "*.jpg *.jpeg *.png *.gif *.bmp *.webp",
                ),
                (
                    "Videos",
                    "*.mp4 *.mov *.avi *.mkv *.webm *.wmv *.m4v",
                ),
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
                "Archivo",
                "El archivo supera 100 MB.",
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

            mime = (
                mimetypes.guess_type(nombre)[0]
                or "application/octet-stream"
            )

            with open(ruta, "rb") as archivo:
                data = self.api(
                    "POST",
                    "/api/files",
                    files={
                        "file": (
                            nombre,
                            archivo,
                            mime,
                        )
                    },
                )

            if tipo == "imagen":
                contenido = f"📷 {nombre}"
            elif tipo == "video":
                contenido = f"🎥 {nombre}"
            else:
                contenido = f"📎 {nombre}"

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

            self.cargar_mensajes()
            self.cargar_usuarios()

        except Exception as e:
            messagebox.showerror(
                "Archivo",
                f"No se pudo enviar el archivo.\n\n{e}",
                parent=self,
            )

    def marcar_leidos(self):
        if not self.conversacion_id:
            return

        try:
            self.api(
                "POST",
                "/api/messages/read",
                json={
                    "conversacion_id": self.conversacion_id,
                },
            )
        except Exception:
            pass

    def refrescar_red(self):
        if not self.sesion_activa:
            return

        try:
            data = self.api(
                "GET",
                "/api/notifications",
                params={
                    "after_id": self.after_id,
                },
            )

            for message in data["mensajes"]:
                self.after_id = max(
                    self.after_id,
                    int(message["id"]),
                )

            self.cargar_usuarios()

            if self.conversacion_id:
                self.cargar_mensajes()

        except Exception as e:
            print("Monitor chat:", e)

        try:
            self.after(
                2000,
                self.refrescar_red,
            )
        except Exception:
            pass

    def descargar_archivo(self, ruta, nombre):
        if not ruta:
            return

        try:
            carpeta = os.path.join(
                self.carpeta_archivos,
                "recibidos",
            )
            os.makedirs(
                carpeta,
                exist_ok=True,
            )

            destino = os.path.join(
                carpeta,
                f"{uuid.uuid4().hex}_"
                f"{os.path.basename(nombre or 'archivo')}",
            )

            response = self.http.get(
                f"{CHAT_SERVER_URL.rstrip('/')}/api/files/{ruta}",
                headers=self.headers,
                timeout=120,
            )
            response.raise_for_status()

            with open(destino, "wb") as archivo:
                archivo.write(response.content)

            self.abrir_archivo(destino)

        except Exception as e:
            messagebox.showerror(
                "Archivo",
                f"No se pudo descargar el archivo.\n\n{e}",
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

    def mostrar_mensaje_inicial(self):
        tk.Label(
            self.contenedor_mensajes,
            text="Selecciona un usuario para comenzar a conversar.",
            bg=self.colores["fondo"],
            fg=self.colores["gris"],
            font=("Arial", 11),
        ).pack(pady=50)

    def actualizar_scroll(self, event=None):
        self.canvas_mensajes.configure(
            scrollregion=self.canvas_mensajes.bbox("all")
        )

    def ajustar_ancho_canvas(self, event):
        self.canvas_mensajes.itemconfig(
            self.ventana_canvas,
            width=event.width,
        )

    def bajar_scroll(self):
        self.canvas_mensajes.update_idletasks()
        self.canvas_mensajes.yview_moveto(1.0)

    def cerrar(self):
        try:
            self.sesion_activa = False
        except Exception:
            pass

        try:
            self.http.close()
        except Exception:
            pass

        self.destroy()
