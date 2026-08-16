import customtkinter as ctxt
import os
import tkinter as tk
from conexion import obtener_conexion, obtener_ruta_recurso, hash_password

# Intentamos cargar PIL para el logo circular
try:
    from PIL import Image
    PILLOW_DISPONIBLE = True
except ImportError:
    PILLOW_DISPONIBLE = False

class VentanaLogin(ctxt.CTk):
    def __init__(self):
        super().__init__()

        self.usuario_autenticado = None
        self.rol_autenticado = None
        self.password_autenticado = None

        # ==================== PALETA DE COLORES ====================
        self.COLOR_FONDO_DEGRADADO = "#290a0e"  # Vinotinto ultra oscuro (fondo de pantalla)
        self.COLOR_TARJETA_CENTRAL = "#3d1116"  # Tarjeta flotante (un poco más clara para el relieve)
        self.COLOR_INPUT_FONDO = "#21080b"      # Fondo oscuro de los campos de texto
        self.COLOR_INPUT_BORDE = "#541b21"      # Borde sutil vinotinto para las cajas
        self.COLOR_BOTON_ROJO = "#9c1c26"       # Rojo vivo del botón "Iniciar Sesión"
        self.COLOR_BOTON_HOVER = "#7a141b"      # Color del botón al pasar el mouse
        self.COLOR_TEXTO_GRIS = "#a69496"       # Gris claro para las etiquetas secundarias

        # Configuración del tamaño de la ventana 
        self.title("Iniciar Sesión - Macilitano Consulting Group")
        self.geometry("450x660")
        self.resizable(False, False)
        self.configure(fg_color=self.COLOR_FONDO_DEGRADADO)
        self.center_window()
        self.after(50, self.traer_ventana_al_frente)

        # ==================== TARJETA VERTICAL CENTRAL ====================
        self.card = ctxt.CTkFrame(
            self, width=380, height=590,
            fg_color=self.COLOR_TARJETA_CENTRAL,
            corner_radius=25,  
            border_width=0
        )
        self.card.place(relx=0.5, rely=0.5, anchor="center")
        self.card.pack_propagate(False)

        # 1. Textos de Identidad de la Firma
        self.lbl_empresa = ctxt.CTkLabel(
            self.card, text="Macilitano",
            font=("Arial", 22, "bold"), text_color="#ffffff"
        )
        self.lbl_empresa.pack(pady=(10, 0))

        self.lbl_subtitulo = ctxt.CTkLabel(
            self.card, text="Consulting Group C.A.",
            font=("Arial", 16, "bold"), text_color="#c42334"
        )
        self.lbl_subtitulo.pack(pady=(2, 12))

        # 2. Logo de la asesoría debajo del nombre
        self.cargar_logo_circular()

        # 3. Campo de Entrada: Usuario
        self.lbl_user_tag = ctxt.CTkLabel(
            self.card, text="Usuario",
            font=("Arial", 13), text_color=self.COLOR_TEXTO_GRIS
        )
        self.lbl_user_tag.pack(anchor="w", padx=35, pady=(5, 2))

        self.entry_usuario = ctxt.CTkEntry(
            self.card, width=310, height=42,
            placeholder_text="Ingrese su usuario",
            fg_color=self.COLOR_INPUT_FONDO, text_color="#ffffff",
            border_color=self.COLOR_INPUT_BORDE, border_width=1, corner_radius=10,
            placeholder_text_color="#5c4547"
        )
        self.entry_usuario.pack(pady=5)

        # 4. Campo de Entrada: Contraseña
        self.lbl_pass_tag = ctxt.CTkLabel(
            self.card, text="Contraseña",
            font=("Arial", 13), text_color=self.COLOR_TEXTO_GRIS
        )
        self.lbl_pass_tag.pack(anchor="w", padx=35, pady=(10, 2))

        self.entry_password = ctxt.CTkEntry(
            self.card, width=310, height=42,
            placeholder_text="••••••••", show="*",
            fg_color=self.COLOR_INPUT_FONDO, text_color="#ffffff",
            border_color=self.COLOR_INPUT_BORDE, border_width=1, corner_radius=10,
            placeholder_text_color="#5c4547"
        )
        self.entry_password.pack(pady=2)

        self.lbl_hint = ctxt.CTkLabel(
            self.card, text="Mínimo 5 caracteres",
            font=("Arial", 11), text_color="#705457"
        )
        self.lbl_hint.pack(anchor="w", padx=35, pady=(0, 10))

        # Espacio dinámico para mensajes de error
        self.lbl_error = ctxt.CTkLabel(self.card, text="", font=("Arial", 12))
        self.lbl_error.pack(pady=2)

        # 5. Botón Redondeado "Iniciar Sesión"
        self.btn_ingresar = ctxt.CTkButton(
            self.card, text="Iniciar Sesión",
            width=310, height=45, font=("Arial", 14, "bold"),
            fg_color=self.COLOR_BOTON_ROJO, text_color="#ffffff",
            hover_color=self.COLOR_BOTON_HOVER, corner_radius=12,  # Redondeado exacto
            command=self.validar_login
        )
        self.btn_ingresar.pack(pady=15)

        # 6. Texto de Soporte al Fondo de la Tarjeta
        self.lbl_soporte = ctxt.CTkLabel(
            self.card, text="¿Soporte técnico? Contactar al administrador",
            font=("Arial", 11), text_color=self.COLOR_TEXTO_GRIS
        )
        self.lbl_soporte.pack(side="bottom", pady=20)

    def cargar_logo_circular(self):
        """Carga el logo centrado respetando las dimensiones del círculo de la firma"""
        ruta_logo = obtener_ruta_recurso("logo.png")
        if PILLOW_DISPONIBLE and os.path.exists(ruta_logo):
            try:
                img_logo = Image.open(ruta_logo)
                # Tamaño calibrado para mantener la proporción perfecta en el eje vertical
                self.logo_ctk = ctxt.CTkImage(light_image=img_logo, dark_image=img_logo, size=(105, 105))
                self.lbl_logo = ctxt.CTkLabel(self.card, image=self.logo_ctk, text="")
                self.lbl_logo.pack(pady=(8, 15))
                return
            except Exception:
                pass

        
        if os.path.exists(ruta_logo):
            try:
                self._tk_logo = tk.PhotoImage(file=ruta_logo)
                
                self.lbl_logo = tk.Label(self.card, image=self._tk_logo, bg=self.COLOR_TARJETA_CENTRAL)
                self.lbl_logo.pack(pady=(8, 15))
                return
            except Exception:
                pass

        
        self.lbl_logo_backup = ctxt.CTkLabel(self.card, text="Ⓜ", font=("Arial", 45), text_color="#ffffff")
        self.lbl_logo_backup.pack(pady=(8, 15))

    def center_window(self):
        self.update_idletasks()
        ancho = 450
        alto = 660
        x = (self.winfo_screenwidth() - ancho) // 2
        y = (self.winfo_screenheight() - alto) // 2
        self.geometry(f"{ancho}x{alto}+{x}+{y}")

    def traer_ventana_al_frente(self):
        self.deiconify()
        self.lift()
        self.focus_force()
        self.wm_attributes('-topmost', True)
        self.after(100, lambda: self.wm_attributes('-topmost', False))

    def validar_login(self):
        usuario = self.entry_usuario.get().strip()
        password = self.entry_password.get().strip()

        if not usuario or not password:
            self.lbl_error.configure(text="Ingrese usuario y contraseña", text_color="#ff4a5a")
            return

        if len(password) < 5:
            self.lbl_error.configure(text="Contraseña demasiado corta", text_color="#ff4a5a")
            return

        try:
            db = obtener_conexion()
            cursor = db.cursor()
            pw_hash = hash_password(password)
            query = "SELECT username, rol FROM usuarios WHERE username = ? AND password_hash = ?"
            cursor.execute(query, (usuario, pw_hash))
            resultado = cursor.fetchone()

            # Si no se encuentra el usuario, intentar también con espacios reemplazados por guiones bajos
            if not resultado:
                alt_usuario = usuario.replace(' ', '_')
                if alt_usuario != usuario:
                    cursor.execute(query, (alt_usuario, pw_hash))
                    resultado = cursor.fetchone()
            cursor.close()
            db.close()

            if resultado:
                self.usuario_autenticado = resultado[0]
                self.rol_autenticado = resultado[1]
                self.password_autenticado = password
                self.destroy()
            else:
                self.usuario_autenticado = None
                self.rol_autenticado = None
                self.lbl_error.configure(text="Usuario o contraseña incorrectos", text_color="#ff4a5a")
        except Exception as e:
            self.usuario_autenticado = None
            self.rol_autenticado = None
            self.lbl_error.configure(text="Error al conectar con la base de datos", text_color="#ff4a5a")
            print(f"Login DB error: {e}")