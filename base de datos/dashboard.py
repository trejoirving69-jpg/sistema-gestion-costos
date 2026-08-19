import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import os
import sqlite3
import mysql.connector
import urllib3
import requests
from bs4 import BeautifulSoup
import re
import datetime
from openpyxl import Workbook
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import traceback
import certifi

# Importaciones desde conexion
from conexion import obtener_conexion, registrar_accion, hash_password, generar_id_cliente, obtener_ruta_recurso 
from chat_red import VentanaChat

TIPOS_EMPRESA = [
    "Industria manufacturera",
    "Comercio mayorista",
    "Comercio minorista / Retail",
    "Servicios profesionales",
    "Tecnología / Software",
    "Construcción",
    "Agricultura / Ganadería / Pesca",
    "Alimentos y bebidas",
    "Transporte / Logística",
    "Salud",
    "Educación",
    "Turismo / Hotelería",
    "Inmobiliario",
    "Finanzas / Seguros",
    "Telecomunicaciones",
    "Energía / Petróleo / Gas",
    "Minería",
    "Automotriz",
    "Textil / Moda",
    "Farmacéutica",
    "Medios / Publicidad",
    "Entretenimiento / Eventos",
    "Servicios personales",
    "ONG / Fundación",
    "Otro / No clasificado",
]

from chat_config import CHAT_SERVER_URL

# Nuevas importaciones para la Etapa 10 (Gráficos integrados en Tkinter)
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ================= AUTOMATIZACIÓN DE LA TASA BCV =================
def obtener_tasa_bcv_automatica():
    """Obtiene la tasa oficial del USD directamente desde el sitio del BCV."""
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
    try:
        url = "https://www.bcv.org.ve/"
        response = requests.get(url, headers=headers, verify=False, timeout=8)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            div_dolar = soup.find('div', id='dolar')
            if div_dolar:
                texto = div_dolar.get_text().strip()
                match = re.search(r'\d+,\d+', texto)
                if match:
                    return float(match.group().replace(',', '.'))
    except Exception:
        pass
    try:
        api_url = "https://open.er-api.com/v6/latest/USD"
        res = requests.get(api_url, timeout=5).json()
        if "rates" in res and "VES" in res["rates"]:
            return float(res["rates"]["VES"])
    except Exception:
        pass
    return None


class DashboardSGC:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Gestión de Costos SGC - Panel Principal")
        self.root.geometry("1150x700") 
        
        self.colores = {
            "sidebar": "#6b1426",       
            "botones_menu": "#1f1f1f",  
            "fondo_main": "#fdfaf2",    
            "oro": "#d4af37",           
            "rojo_eliminar": "#9e2a2b", 
            "verde_aprobar": "#2a9d8f",
            "blanco_tarjeta": "#ffffff",
            "texto_oscuro": "#333333"
        }
        
        self.estilo = ttk.Style()
        self.estilo.theme_use("default")
        
        self.estilo.configure("Treeview.Heading", 
                    background=self.colores["botones_menu"], 
                    foreground=self.colores["oro"], 
                    font=("Arial", 10, "bold"), 
                    relief="flat")
        self.estilo.map("Treeview.Heading", background=[('active', '#2a2a2a')])
        
        self.estilo.configure("Treeview", 
                background="white", 
                fieldbackground="white", 
                foreground="black", 
                font=("Arial", 10), 
                rowheight=24)
        self.estilo.map("Treeview", 
                background=[('selected', self.colores["sidebar"])], 
                foreground=[('selected', 'white')])
        
        self.ultimo_mensaje_notificado = 0
        self.notificaciones_chat = 0
        self.solicitudes_token = None
        self.ultima_solicitud_web_id = 0
        self.monitor_solicitudes_activo = False
        self.iniciar_monitor_mensajes()
        self.configurar_interfaz()

    def configurar_interfaz(self):
        sidebar = tk.Frame(self.root, bg=self.colores["sidebar"], width=220)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        
        ruta_logo = obtener_ruta_recurso("logo.png")
        if os.path.exists(ruta_logo):
            try:
                self._tk_logo = tk.PhotoImage(file=ruta_logo)
                self._tk_logo = self._tk_logo.subsample(2, 2)
                lbl_logo = tk.Label(sidebar, image=self._tk_logo, bg=self.colores["sidebar"])
            except Exception:
                lbl_logo = tk.Label(sidebar, text="MACILITANO\nCONSULTING", fg="white", bg=self.colores["sidebar"], font=("Arial", 14, "bold"))
        else:
            lbl_logo = tk.Label(sidebar, text="MACILITANO\nCONSULTING", fg="white", bg=self.colores["sidebar"], font=("Arial", 14, "bold"))
        lbl_logo.pack(pady=20)

        # Construir botones del menú y guardar referencias para control de permisos
        self.menu_buttons = {}
        self.notas_pendientes = 0

        def add_menu_button(texto, comando):
            btn = tk.Button(sidebar, text=texto, bg=self.colores["botones_menu"], fg=self.colores["oro"],
                            font=("Arial", 11, "bold"), relief="flat", activebackground=self.colores["oro"],
                            activeforeground="black", anchor="w", padx=20, command=comando)
            btn.pack(fill="x", pady=4, padx=10)
            self.menu_buttons[texto] = btn
            return btn

        add_menu_button("Inicio", self.mostrar_inicio)
        add_menu_button("Clientes", self.mostrar_clientes)
        add_menu_button("Asesores", self.mostrar_asesores)
        add_menu_button("Servicios", self.mostrar_servicios)
        add_menu_button("Gestión de Costos", self.mostrar_gestion_costos)
        add_menu_button("Solicitudes", self.mostrar_notificaciones)
        
        self.btn_chat = add_menu_button(
    "💬 Chats",
    self.abrir_chat
    )
    
        # Crear el botón de Historial pero su visibilidad/estado dependerá del rol
        self.btn_historial = add_menu_button("Historial Logs", self.mostrar_historial)

        self.area_principal = tk.Frame(self.root, bg=self.colores["fondo_main"])
        self.area_principal.pack(side="right", fill="both", expand=True)
        
        self.mostrar_inicio()
        self.actualizar_alertas_globales()
    
        # ============================================================
    # CHAT INTERNO
    # ============================================================

        # ============================================================
    # CHAT INTERNO
    # ============================================================

    def abrir_chat(self):

        usuario_actual = getattr(
            self,
            "usuario_autenticado",
            None
        )
        password_actual = getattr(
            self,
            "password_autenticado",
            None
        )

        if not usuario_actual:

            messagebox.showwarning(
                "Chats",
                "No se pudo identificar al usuario actual."
            )

            return

        try:

            ventana = VentanaChat(
                self.root,
                usuario_actual,
                password_actual
            )

            ventana.transient(
                self.root
            )

            # Guardamos referencia para evitar
            # que Python la destruya
            self.ventana_chat = ventana

            ventana.protocol(
                "WM_DELETE_WINDOW",
                lambda: self.cerrar_ventana_chat()
            )

        except Exception as e:

            messagebox.showerror(
                "Error",
                f"No se pudo abrir el módulo de Chats.\n\n{e}"
            )

    def cerrar_ventana_chat(self):

        if hasattr(self, "ventana_chat"):

            try:
                self.ventana_chat.cerrar()
            except Exception:
                pass

            self.ventana_chat = None
    def aplicar_permisos(self):
        """Aplica reglas de visibilidad/permiso en la UI según `self.rol_autenticado`."""
        rol = getattr(self, 'rol_autenticado', None)
        try:
            if hasattr(self, 'btn_historial'):
                if rol is None or str(rol).lower() != 'administrador':
                    # Ocultar el botón de historial si no es administrador
                    try:
                        self.btn_historial.pack_forget()
                    except Exception:
                        self.btn_historial.config(state='disabled')
                else:
                    # Asegurar que esté visible si es administrador
                    try:
                        self.btn_historial.pack(fill="x", pady=4, padx=10)
                    except Exception:
                        self.btn_historial.config(state='normal')
        except Exception:
            pass
        
    def crear_panel_principal(self, titulo):
        for widget in self.area_principal.winfo_children():
            widget.destroy()
        
        lbl_titulo = tk.Label(self.area_principal, text=titulo, bg=self.colores["fondo_main"], 
                fg=self.colores["sidebar"], font=("Arial", 20, "bold"))
        lbl_titulo.pack(anchor="w", padx=25, pady=20)
        
        return self.area_principal

    def actualizar_alertas_globales(self):
        self.notas_pendientes = 0
        if not getattr(self, "monitor_solicitudes_activo", False):
            self.monitor_solicitudes_activo = True
            self.root.after(1500, self.monitor_solicitudes_web)

    def _api_solicitudes(self, metodo, ruta, **kwargs):
        """Cliente centralizado para solicitudes web alojadas en PythonAnywhere."""
        if not CHAT_SERVER_URL.strip():
            raise RuntimeError("CHAT_SERVER_URL no está configurado")

        if not self.solicitudes_token:
            usuario = getattr(self, "usuario_autenticado", None)
            password = getattr(self, "password_autenticado", None)
            if not usuario or not password:
                raise RuntimeError("No hay credenciales de sesión disponibles")
            login = requests.post(
                f"{CHAT_SERVER_URL.rstrip('/')}/api/session",
                json={"username": usuario, "password": password},
                timeout=15,
                verify=certifi.where(),
            )
            login.raise_for_status()
            data_login = login.json()
            if not data_login.get("ok"):
                raise RuntimeError(data_login.get("error", "No se pudo iniciar sesión remota"))
            self.solicitudes_token = data_login["token"]

        headers = {"Authorization": f"Bearer {self.solicitudes_token}"}
        respuesta = requests.request(
            metodo,
            f"{CHAT_SERVER_URL.rstrip('/')}{ruta}",
            headers=headers,
            timeout=15,
            verify=certifi.where(),
            **kwargs,
        )
        if respuesta.status_code == 401:
            self.solicitudes_token = None
            return self._api_solicitudes(metodo, ruta, **kwargs)
        respuesta.raise_for_status()
        data = respuesta.json()
        if not data.get("ok"):
            raise RuntimeError(data.get("error", "Error en API de solicitudes"))
        return data

    def monitor_solicitudes_web(self):
        """Muestra una alerta estilo app cuando entra una nueva solicitud desde la web."""
        try:
            data = self._api_solicitudes("GET", "/api/solicitudes?estado=Pendiente")
            items = data.get("items", [])
            ids = [int(x.get("id", 0)) for x in items]
            max_id = max(ids, default=0)
            if self.ultima_solicitud_web_id == 0:
                self.ultima_solicitud_web_id = max_id
            else:
                nuevas = [x for x in items if int(x.get("id", 0)) > self.ultima_solicitud_web_id]
                for solicitud in sorted(nuevas, key=lambda x: int(x.get("id", 0))):
                    self.mostrar_notificacion_solicitud(solicitud)
                self.ultima_solicitud_web_id = max(self.ultima_solicitud_web_id, max_id)
        except Exception as e:
            print(f"Monitor solicitudes web: {e}")
        try:
            self.root.after(5000, self.monitor_solicitudes_web)
        except Exception:
            pass

    def mostrar_notificacion_solicitud(self, solicitud):
        n = tk.Toplevel(self.root)
        n.title("Nueva solicitud web")
        n.resizable(False, False)
        n.configure(bg="white")
        n.attributes("-topmost", True)
        ancho, alto = 385, 180
        n.update_idletasks()
        sw, sh = n.winfo_screenwidth(), n.winfo_screenheight()
        n.geometry(f"{ancho}x{alto}+{sw-ancho-25}+{max(20, sh-alto-70)}")
        tk.Label(n, text="🔔  Nueva solicitud de asesoría", bg="white", fg=self.colores["sidebar"], font=("Arial", 13, "bold")).pack(anchor="w", padx=15, pady=(12, 3))
        tk.Label(n, text=solicitud.get("cliente_potencial") or "Cliente", bg="white", fg="#222", font=("Arial", 10, "bold")).pack(anchor="w", padx=15)
        tk.Label(n, text=solicitud.get("servicio_interes") or "Servicio no especificado", bg="white", fg="#666", font=("Arial", 9), wraplength=345, justify="left").pack(anchor="w", padx=15, pady=(2, 8))
        tk.Button(n, text="Ver solicitud", bg=self.colores["sidebar"], fg="white", relief="flat", cursor="hand2", command=lambda: (n.destroy(), self.mostrar_notificaciones())).pack(anchor="e", padx=15, pady=(0, 10))
        n.after(9000, lambda: n.destroy() if n.winfo_exists() else None)
    
    def mostrar_inicio(self):
        panel = self.crear_panel_principal("Panel de Control Ejecutivo")
        
        total_clientes = 0
        total_servicios = 0
        total_solicitudes = 0
        ventas_mes = 0.0
        
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) FROM clientes")
            total_clientes = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM servicios")
            total_servicios = cursor.fetchone()[0]
            
            cursor.execute("SELECT COUNT(*) FROM solicitudes_servicio")
            total_solicitudes = cursor.fetchone()[0]
            
            cursor.execute("SELECT SUM(total_usd) FROM solicitudes_servicio WHERE activo = 'Activo'")
            monto_ventas = cursor.fetchone()[0]
            if monto_ventas:
                ventas_mes = monto_ventas
                
            conn.close()
        except Exception as e:
            print(f"Error consultando métricas locales: {e}")
            
        tasa_bcv = obtener_tasa_bcv_automatica() or 40.00 

        frame_tarjetas = tk.Frame(panel, bg=self.colores["fondo_main"])
        frame_tarjetas.pack(fill="x", padx=25, pady=10)
        
        metricas = [
            ("Clientes Activos", str(total_clientes), "#2a9d8f"),
            ("Catálogo Servicios", str(total_servicios), "#457b9d"),
            ("Solicitudes Proc.", str(total_solicitudes), "#e07a5f"),
            ("Ventas del Mes", f"$ {ventas_mes:,.2f}", "#6b1426"),
            ("Tasa Oficial BCV", f"{tasa_bcv:.2f} Bs.", "#d4af37")
        ]
        
        for idx, (titulo, valor, color_borde) in enumerate(metricas):
            tarjeta = tk.Frame(frame_tarjetas, bg=self.colores["blanco_tarjeta"], 
                            highlightbackground=color_borde, highlightthickness=2, bd=0)
            tarjeta.grid(row=0, column=idx, padx=8, pady=5, sticky="nsew")
            frame_tarjetas.grid_columnconfigure(idx, weight=1)
            
            tk.Label(tarjeta, text=titulo, fg="#777777", bg=self.colores["blanco_tarjeta"], font=("Arial", 10, "bold")).pack(pady=(10, 2), padx=10)
            tk.Label(tarjeta, text=valor, fg=self.colores["texto_oscuro"], bg=self.colores["blanco_tarjeta"], font=("Arial", 16, "bold")).pack(pady=(0, 10), padx=10)

        frame_grafico = tk.Frame(panel, bg=self.colores["blanco_tarjeta"], bd=1, relief="solid")
        frame_grafico.pack(fill="both", expand=True, padx=25, pady=20)
        
        lbl_grafico_titulo = tk.Label(frame_grafico, text="📊 Evolución de Ingresos Mensuales (USD)", 
                bg=self.colores["blanco_tarjeta"], fg=self.colores["sidebar"], font=("Arial", 12, "bold"))
        lbl_grafico_titulo.pack(anchor="w", padx=15, pady=10)
        
        figura = Figure(figsize=(6, 3), dpi=100)
        eje = figura.add_subplot(111)
        
        meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul']
        valores_ingresos = [1200, 2400, 1850, 3100, 2900, 4200, max(ventas_mes, 1500)]
        
        eje.plot(meses, valores_ingresos, marker='o', color=self.colores["sidebar"], linewidth=2, label="Ventas Realizadas")
        eje.fill_between(meses, valores_ingresos, color=self.colores["sidebar"], alpha=0.1)
        eje.set_facecolor('#fafafa')
        eje.grid(True, linestyle='--', alpha=0.5)
        
        canvas = FigureCanvasTkAgg(figura, master=frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=15, pady=(0, 15))
        

    def mostrar_clientes(self):
        panel = self.crear_panel_principal("Gestión de Clientes y Contabilidad")
        
        # Barra superior dividida en dos filas para evitar que los botones se solapen.
        # Fila 1: búsqueda y filtros.
        filtros_bar = tk.Frame(panel, bg=self.colores["fondo_main"])
        filtros_bar.pack(fill="x", padx=25, pady=(5, 5))

        lbl_buscar = tk.Label(
            filtros_bar,
            text="Nombre / empresa:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        )
        lbl_buscar.pack(side="left", padx=(0, 5))

        self.entry_buscar = tk.Entry(
            filtros_bar,
            font=("Arial", 10),
            width=22,
            relief="solid",
            bd=1
        )
        self.entry_buscar.pack(side="left", padx=5, ipady=3)
        self.entry_buscar.bind("<KeyRelease>", lambda event: self.actualizar_tabla_clientes())

        lbl_servicio = tk.Label(
            filtros_bar,
            text="Servicio:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        )
        lbl_servicio.pack(side="left", padx=(12, 5))

        self.combo_servicio = ttk.Combobox(
            filtros_bar,
            state="readonly",
            width=20
        )
        self.combo_servicio.pack(side="left", padx=5)
        self.combo_servicio.bind(
            "<<ComboboxSelected>>",
            lambda event: self.actualizar_tabla_clientes()
        )

        lbl_tipo = tk.Label(
            filtros_bar,
            text="Tipo cliente:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        )
        lbl_tipo.pack(side="left", padx=(12, 5))

        self.combo_tipo_cliente = ttk.Combobox(
            filtros_bar,
            state="readonly",
            width=22
        )
        self.combo_tipo_cliente.pack(side="left", padx=5)
        self.combo_tipo_cliente.bind(
            "<<ComboboxSelected>>",
            lambda event: self.actualizar_tabla_clientes()
        )

        try:
            conn = obtener_conexion()
            cursor = conn.cursor()

            cursor.execute(
                """
                SELECT nombre_servicio
                FROM servicios
                WHERE nombre_servicio IS NOT NULL
                  AND TRIM(nombre_servicio) != ''
                ORDER BY nombre_servicio
                """
            )
            servicios = [row[0] for row in cursor.fetchall()]

            cursor.execute(
                """
                SELECT DISTINCT servicio
                FROM clientes
                WHERE servicio IS NOT NULL
                  AND TRIM(servicio) != ''
                ORDER BY servicio
                """
            )
            for (servicio,) in cursor.fetchall():
                if servicio and servicio not in servicios:
                    servicios.append(servicio)

            tipos = ["Todos"] + TIPOS_EMPRESA

            cursor.execute(
                """
                SELECT DISTINCT industria
                FROM clientes
                WHERE industria IS NOT NULL
                  AND TRIM(industria) != ''
                ORDER BY industria
                """
            )
            for (tipo,) in cursor.fetchall():
                if tipo and tipo not in tipos:
                    tipos.append(tipo)

            conn.close()

        except Exception:
            servicios = []
            tipos = ["Todos"] + TIPOS_EMPRESA

        self.combo_servicio["values"] = ["Todos"] + servicios
        self.combo_servicio.set("Todos")

        self.combo_tipo_cliente["values"] = tipos
        self.combo_tipo_cliente.set("Todos")

        # Fila 2: acciones. "Agregar Cliente" queda separado y siempre visible.
        acciones_bar = tk.Frame(panel, bg=self.colores["fondo_main"])
        acciones_bar.pack(fill="x", padx=25, pady=(0, 10))

        btn_agregar = tk.Button(
            acciones_bar,
            text="➕ Agregar Cliente",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=14,
            pady=5,
            command=self.modal_agregar_cliente
        )
        btn_agregar.pack(side="left", padx=(0, 12))

        btn_ordenar = tk.Button(
            acciones_bar,
            text="↕ Ordenar columnas",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=self.ordenar_columnas_clientes
        )
        btn_ordenar.pack(side="right", padx=5)

        btn_pdf = tk.Button(
            acciones_bar,
            text="✅ Exportar PDF",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=self.exportar_clientes_pdf
        )
        btn_pdf.pack(side="right", padx=5)

        btn_excel = tk.Button(
            acciones_bar,
            text="✅ Exportar Excel",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=self.exportar_clientes_excel
        )
        btn_excel.pack(side="right", padx=5)

        btn_eliminar = tk.Button(
            acciones_bar,
            text="🗑 Eliminar",
            bg=self.colores["rojo_eliminar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=self.eliminar_cliente_bd
        )
        btn_eliminar.pack(side="right", padx=5)

        btn_editar = tk.Button(
            acciones_bar,
            text="✏ Editar",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=self.modal_editar_cliente
        )
        btn_editar.pack(side="right", padx=5)

        tk.Label(panel, text="📋 Directorio de Contactos y Servicios", bg=self.colores["fondo_main"], 
                fg=self.colores["sidebar"], font=("Arial", 12, "bold")).pack(anchor="w", padx=25, pady=(5, 2))
        
        # NUEVO ORDEN DE COLUMNAS REQUERIDO
        columnas_c = (
            "ID", 
            "Código Cliente", 
            "Nombre Cliente", 
            "Servicio Solicitado", 
            "Teléfono de Contacto", 
            "Nombre Persona Contacto", 
            "Correo", 
            "Status"
        )
        self.tabla_clientes = ttk.Treeview(panel, columns=columnas_c, show="headings", height=15)
        
        # Configuración del ancho de las columnas ajustado al nuevo orden
        for col in columnas_c:
            self.tabla_clientes.heading(col, text=col)
            if col in ["Nombre Cliente", "Nombre Persona Contacto", "Correo"]:
                self.tabla_clientes.column(col, anchor="w", width=150)
            elif col in ["Servicio Solicitado"]:
                self.tabla_clientes.column(col, anchor="w", width=210)
            elif col in ["ID", "Código Cliente"]:
                self.tabla_clientes.column(col, anchor="center", width=80)
            else:
                self.tabla_clientes.column(col, anchor="center", width=110)
                
        self.tabla_clientes.pack(fill="both", expand=True, padx=25, pady=(0, 15))
        self.tabla_clientes.bind("<Double-1>", self.alternar_estado_servicio)
        self.actualizar_tabla_clientes()

    def exportar_clientes_excel(self):
        try:
            if not hasattr(self, 'tabla_clientes'):
                return
            path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx'), ('All Files', '*.*')])
            if not path:
                return
            wb = Workbook()
            ws = wb.active
            ws.title = 'Clientes'
            columnas = self.tabla_clientes['columns']
            ws.append(columnas)
            for item in self.tabla_clientes.get_children():
                ws.append(self.tabla_clientes.item(item, 'values'))
            wb.save(path)
            messagebox.showinfo('Exportar Excel', f'Clientes exportados con éxito a:\n{path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar a Excel: {e}')

    def exportar_clientes_pdf(self):
        try:
            if not hasattr(self, 'tabla_clientes'):
                return
            path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[('PDF Files', '*.pdf'), ('All Files', '*.*')])
            if not path:
                return
            c = canvas.Canvas(path, pagesize=letter)
            width, height = letter
            columnas = list(self.tabla_clientes['columns'])
            x_positions = [50, 110, 200, 360, 520, 620, 720, 820]
            y = height - 50
            c.setFont('Helvetica-Bold', 10)
            for idx, col in enumerate(columnas):
                x = x_positions[idx] if idx < len(x_positions) else 50 + idx * 100
                c.drawString(x, y, str(col))
            c.setFont('Helvetica', 9)
            y -= 20
            for item in self.tabla_clientes.get_children():
                valores = self.tabla_clientes.item(item, 'values')
                if y < 50:
                    c.showPage()
                    y = height - 50
                    c.setFont('Helvetica-Bold', 10)
                    for idx, col in enumerate(columnas):
                        x = x_positions[idx] if idx < len(x_positions) else 50 + idx * 100
                        c.drawString(x, y, str(col))
                    c.setFont('Helvetica', 9)
                    y -= 20
                for idx, valor in enumerate(valores):
                    x = x_positions[idx] if idx < len(x_positions) else 50 + idx * 100
                    c.drawString(x, y, str(valor))
                y -= 18
            c.save()
            messagebox.showinfo('Exportar PDF', f'Clientes exportados con éxito a:\n{path}')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo exportar a PDF: {e}')

    def ordenar_columnas_clientes(self):
        try:
            if not hasattr(self, 'tabla_clientes'):
                return
            columna = 'Nombre Cliente'
            orden = getattr(self, 'orden_cliente_asc', True)
            items = [(self.tabla_clientes.set(k, columna), k) for k in self.tabla_clientes.get_children('')]
            items.sort(key=lambda t: str(t[0]).lower(), reverse=not orden)
            for index, (_, k) in enumerate(items):
                self.tabla_clientes.move(k, '', index)
            self.orden_cliente_asc = not orden
            direccion = 'ascendente' if orden else 'descendente'
            messagebox.showinfo('Ordenar columnas', f'Tabla ordenada por "{columna}" {direccion}.')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo ordenar la tabla: {e}')

    def actualizar_tabla_clientes(self):
        """Actualiza el directorio de clientes aplicando filtros en tiempo real."""
        for item in self.tabla_clientes.get_children(): 
            self.tabla_clientes.delete(item)
            
        texto_busqueda = self.entry_buscar.get().strip() if hasattr(self, 'entry_buscar') else ""
        servicio_filtro = self.combo_servicio.get() if hasattr(self, 'combo_servicio') else "Todos"
        tipo_filtro = self.combo_tipo_cliente.get() if hasattr(self, 'combo_tipo_cliente') else "Todos"

        condiciones = []
        params = []

        if texto_busqueda:
            condiciones.append("(c.nombre LIKE ? OR c.correo LIKE ? OR c.codigo LIKE ? OR c.telefono LIKE ? OR COALESCE(c.nombre_contacto, c.nombre) LIKE ?)")
            texto_like = f"%{texto_busqueda}%"
            params.extend([texto_like, texto_like, texto_like, texto_like, texto_like])

        if servicio_filtro and servicio_filtro != "Todos":
            condiciones.append("COALESCE(s.nombre_servicio, c.servicio, '') LIKE ?")
            params.append(f"%{servicio_filtro}%")

        if tipo_filtro and tipo_filtro != "Todos":
            condiciones.append("LOWER(COALESCE(c.industria, '')) LIKE ?")
            params.append(f"%{tipo_filtro.lower()}%")

        where_sql = f" WHERE {' AND '.join(condiciones)}" if condiciones else ""

        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            cursor.execute("PRAGMA table_info(clientes)")
            columnas = [col[1] for col in cursor.fetchall()]
            if "estado" not in columnas:
                cursor.execute("ALTER TABLE clientes ADD COLUMN estado TEXT DEFAULT 'Activo'")
                conn.commit()

            # Estructuración de datos 
            # ID | Código Cliente | Nombre Cliente | Servicio Solicitado | Teléfono | Persona Contacto | Correo | Status
            cursor.execute(f"""
                SELECT 
                    c.id, 
                    c.codigo, 
                    c.nombre, 
                    IFNULL(COALESCE(s.nombre_servicio, c.servicio), 'Ninguno'), 
                    IFNULL(c.telefono, 'No registrado'),
                    IFNULL(c.nombre_contacto, c.nombre), 
                    IFNULL(c.correo, 'No registrado'), 
                    IFNULL(c.estado, 'Activo')
                FROM clientes c
                LEFT JOIN solicitudes_servicio ss ON c.id = ss.cliente_id
                LEFT JOIN servicios s ON ss.servicio_id = s.id
                {where_sql}
                GROUP BY c.id
            """, params)
            
            for fila in cursor.fetchall():
                self.tabla_clientes.insert("", "end", values=fila)
            conn.close()
        except Exception as e:
            print(f"Error al actualizar listas: {e}")


    # ----------------- GESTIÓN DE COSTOS / FINANZAS -----------------
    def _obtener_configuracion_financiera(self):
        conn = obtener_conexion()
        cur = conn.cursor()
        cur.execute("""
            SELECT iva_pct, igtf_pct,
                   igtf_efectivo, igtf_transferencia, igtf_pago_movil
            FROM configuracion_financiera
            WHERE id = 1
        """)
        fila = cur.fetchone()
        conn.close()

        if not fila:
            return {
                "iva_pct": 16.0,
                "igtf_pct": 3.0,
                "igtf_efectivo": 1,
                "igtf_transferencia": 0,
                "igtf_pago_movil": 0,
            }

        return {
            "iva_pct": float(fila[0]),
            "igtf_pct": float(fila[1]),
            "igtf_efectivo": int(fila[2]),
            "igtf_transferencia": int(fila[3]),
            "igtf_pago_movil": int(fila[4]),
        }

    def _aplica_igtf(self, metodo_pago, config=None):
        config = config or self._obtener_configuracion_financiera()
        metodo = (metodo_pago or "").strip().lower()

        if "efectivo" in metodo:
            return bool(config["igtf_efectivo"])
        if "transferencia" in metodo:
            return bool(config["igtf_transferencia"])
        if "pago movil" in metodo or "pago móvil" in metodo:
            return bool(config["igtf_pago_movil"])

        return False

    def configurar_parametros_financieros(self):
        try:
            cfg = self._obtener_configuracion_financiera()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar la configuración financiera: {e}")
            return

        modal = tk.Toplevel(self.root)
        modal.title("Configuración financiera")
        modal.geometry("540x480")
        modal.configure(bg=self.colores["fondo_main"])
        modal.transient(self.root)
        modal.grab_set()
        modal.resizable(False, False)

        tk.Label(
            modal,
            text="Parámetros financieros",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 15, "bold")
        ).pack(pady=(22, 6))

        tk.Label(
            modal,
            text="Estas tasas y reglas se aplican a nuevas operaciones.",
            bg=self.colores["fondo_main"],
            fg=self.colores["texto_oscuro"],
            font=("Arial", 9, "italic")
        ).pack(pady=(0, 16))

        frame = tk.Frame(modal, bg=self.colores["fondo_main"])
        frame.pack(fill="x", padx=35)

        tk.Label(frame, text="IVA (%)", bg=self.colores["fondo_main"], font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=8)
        ent_iva = tk.Entry(frame, width=16)
        ent_iva.insert(0, str(cfg["iva_pct"]))
        ent_iva.grid(row=0, column=1, sticky="w", pady=8)

        tk.Label(frame, text="IGTF (%)", bg=self.colores["fondo_main"], font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=8)
        ent_igtf = tk.Entry(frame, width=16)
        ent_igtf.insert(0, str(cfg["igtf_pct"]))
        ent_igtf.grid(row=1, column=1, sticky="w", pady=8)

        tk.Label(
            frame,
            text="Aplicar IGTF según método de pago:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(18, 8))

        var_efectivo = tk.BooleanVar(value=bool(cfg["igtf_efectivo"]))
        var_transferencia = tk.BooleanVar(value=bool(cfg["igtf_transferencia"]))
        var_pago_movil = tk.BooleanVar(value=bool(cfg["igtf_pago_movil"]))

        tk.Checkbutton(frame, text="Efectivo", variable=var_efectivo, bg=self.colores["fondo_main"]).grid(row=3, column=0, columnspan=2, sticky="w", pady=4)
        tk.Checkbutton(frame, text="Transferencia", variable=var_transferencia, bg=self.colores["fondo_main"]).grid(row=4, column=0, columnspan=2, sticky="w", pady=4)
        tk.Checkbutton(frame, text="Pago Móvil", variable=var_pago_movil, bg=self.colores["fondo_main"]).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)

        def guardar():
            try:
                iva = float(ent_iva.get().strip().replace(",", "."))
                igtf = float(ent_igtf.get().strip().replace(",", "."))
                if not (0 <= iva <= 100 and 0 <= igtf <= 100):
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Valores inválidos", "Las tasas deben ser números entre 0 y 100.", parent=modal)
                return

            try:
                conn = obtener_conexion()
                cur = conn.cursor()
                cur.execute("""
                    UPDATE configuracion_financiera
                    SET iva_pct = ?, igtf_pct = ?,
                        igtf_efectivo = ?, igtf_transferencia = ?, igtf_pago_movil = ?
                    WHERE id = 1
                """, (
                    iva, igtf,
                    1 if var_efectivo.get() else 0,
                    1 if var_transferencia.get() else 0,
                    1 if var_pago_movil.get() else 0
                ))
                conn.commit()
                conn.close()

                registrar_accion(
                    getattr(self, "usuario_autenticado", "Sistema"),
                    "ACTUALIZÓ CONFIGURACIÓN FINANCIERA",
                    f"IVA {iva:.2f}% | IGTF {igtf:.2f}%."
                )

                messagebox.showinfo("Guardado", "Configuración financiera actualizada.", parent=modal)
                modal.destroy()
                self.mostrar_gestion_costos()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar: {e}", parent=modal)

        tk.Button(
            modal,
            text="Guardar configuración",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            command=guardar
        ).pack(pady=24)

    def mostrar_gestion_costos(self):
        panel = self.crear_panel_principal("Gestión de Costos y Cuentas por Cobrar")

        try:
            cfg = self._obtener_configuracion_financiera()
        except Exception:
            cfg = {"iva_pct": 16.0, "igtf_pct": 3.0}

        cabecera = tk.Frame(panel, bg=self.colores["fondo_main"])
        cabecera.pack(fill="x", padx=25, pady=(0, 10))

        tk.Label(
            cabecera,
            text=f"IVA configurado: {cfg['iva_pct']:.2f}%   |   IGTF configurado: {cfg['igtf_pct']:.2f}%",
            bg=self.colores["fondo_main"],
            fg=self.colores["texto_oscuro"],
            font=("Arial", 10, "italic")
        ).pack(side="left")

        tk.Button(
            cabecera,
            text="⚙ Configurar tasas",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            command=self.configurar_parametros_financieros
        ).pack(side="right")

        try:
            conn = obtener_conexion()
            cur = conn.cursor()

            cur.execute("SELECT COUNT(*), COALESCE(SUM(total_usd),0), COALESCE(SUM(total_bs),0) FROM solicitudes_servicio")
            total_ops, total_usd, total_bs = cur.fetchone()

            cur.execute("""
                SELECT COALESCE(SUM(total_usd),0)
                FROM solicitudes_servicio
                WHERE COALESCE(estado_pago,'Pendiente') = 'Pendiente'
            """)
            por_cobrar = float(cur.fetchone()[0] or 0)

            cur.execute("""
                SELECT COALESCE(SUM(total_usd),0)
                FROM solicitudes_servicio
                WHERE COALESCE(estado_pago,'Pendiente') = 'Pagado'
            """)
            cobrado = float(cur.fetchone()[0] or 0)

            conn.close()
        except Exception:
            total_ops, total_usd, total_bs, por_cobrar, cobrado = 0, 0, 0, 0, 0

        tarjetas = tk.Frame(panel, bg=self.colores["fondo_main"])
        tarjetas.pack(fill="x", padx=25, pady=(0, 12))

        metricas = [
            ("Operaciones", str(total_ops)),
            ("Facturado USD", f"$ {float(total_usd):,.2f}"),
            ("Cobrado USD", f"$ {cobrado:,.2f}"),
            ("Por cobrar USD", f"$ {por_cobrar:,.2f}"),
        ]

        for idx, (titulo, valor) in enumerate(metricas):
            tarjeta = tk.Frame(
                tarjetas,
                bg=self.colores["blanco_tarjeta"],
                highlightbackground=self.colores["oro"],
                highlightthickness=1
            )
            tarjeta.grid(row=0, column=idx, padx=5, sticky="nsew")
            tarjetas.grid_columnconfigure(idx, weight=1)

            tk.Label(tarjeta, text=titulo, bg=self.colores["blanco_tarjeta"], fg="#777", font=("Arial", 9, "bold")).pack(pady=(8, 2), padx=8)
            tk.Label(tarjeta, text=valor, bg=self.colores["blanco_tarjeta"], fg=self.colores["sidebar"], font=("Arial", 14, "bold")).pack(pady=(0, 8), padx=8)

        botones = tk.Frame(panel, bg=self.colores["fondo_main"])
        botones.pack(fill="x", padx=25, pady=(0, 8))

        columnas = (
            "ID", "Fecha", "Código", "Cliente", "Servicio", "Base USD",
            "IVA", "IGTF", "Total USD", "Tasa BCV", "Total Bs.",
            "Método", "Estado pago"
        )

        tabla = ttk.Treeview(panel, columns=columnas, show="headings", height=13)

        for col in columnas:
            tabla.heading(col, text=col)
            if col in ("Cliente", "Servicio"):
                tabla.column(col, width=180, anchor="w")
            elif col == "Método":
                tabla.column(col, width=105, anchor="center")
            else:
                tabla.column(col, width=85, anchor="center")

        sb = ttk.Scrollbar(panel, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", pady=8, padx=(0, 18))
        tabla.pack(fill="both", expand=True, padx=(25, 0), pady=8)

        try:
            conn = obtener_conexion()
            cur = conn.cursor()
            cur.execute("""
                SELECT
                    ss.id,
                    COALESCE(ss.fecha_solicitud,''),
                    COALESCE(c.codigo,''),
                    c.nombre,
                    s.nombre_servicio,
                    ss.monto,
                    ss.iva_usd,
                    ss.igtf_usd,
                    ss.total_usd,
                    ss.tasa_bcv,
                    ss.total_bs,
                    COALESCE(ss.metodo_pago,''),
                    COALESCE(ss.estado_pago,'Pendiente')
                FROM solicitudes_servicio ss
                INNER JOIN clientes c ON c.id = ss.cliente_id
                INNER JOIN servicios s ON s.id = ss.servicio_id
                ORDER BY ss.id DESC
            """)

            for f in cur.fetchall():
                tabla.insert("", "end", values=(
                    f[0], f[1], f[2], f[3], f[4],
                    f"$ {float(f[5] or 0):,.2f}",
                    f"$ {float(f[6] or 0):,.2f}",
                    f"$ {float(f[7] or 0):,.2f}",
                    f"$ {float(f[8] or 0):,.2f}",
                    f"{float(f[9] or 0):.4f}",
                    f"Bs. {float(f[10] or 0):,.2f}",
                    f[11], f[12]
                ))

            conn.close()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar Gestión de Costos: {e}")

        def cambiar_estado(nuevo_estado):
            sel = tabla.selection()

            if not sel:
                messagebox.showwarning("Seleccione una operación", "Seleccione una operación financiera.")
                return

            valores = tabla.item(sel[0], "values")
            operacion_id = int(valores[0])

            try:
                conn = obtener_conexion()
                cur = conn.cursor()

                if nuevo_estado == "Pagado":
                    cur.execute("""
                        UPDATE solicitudes_servicio
                        SET estado_pago = 'Pagado', fecha_pago = date('now')
                        WHERE id = ?
                    """, (operacion_id,))
                else:
                    cur.execute("""
                        UPDATE solicitudes_servicio
                        SET estado_pago = 'Pendiente', fecha_pago = NULL
                        WHERE id = ?
                    """, (operacion_id,))

                conn.commit()
                conn.close()

                registrar_accion(
                    getattr(self, "usuario_autenticado", "Sistema"),
                    "ACTUALIZÓ ESTADO DE PAGO",
                    f"Operación #{operacion_id}: {nuevo_estado}."
                )

                self.mostrar_gestion_costos()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cambiar el estado: {e}")

        tk.Button(
            botones,
            text="✓ Marcar Pagado",
            bg=self.colores["verde_aprobar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            command=lambda: cambiar_estado("Pagado")
        ).pack(side="left")

        tk.Button(
            botones,
            text="↩ Marcar Pendiente",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            command=lambda: cambiar_estado("Pendiente")
        ).pack(side="left", padx=8)

    # ----------------- GESTIÓN DE ASESORES -----------------
    def mostrar_asesores(self):
        panel = self.crear_panel_principal("Gestión de Asesores")

        top_bar = tk.Frame(panel, bg=self.colores["fondo_main"]) 
        top_bar.pack(fill="x", padx=25, pady=(5,10))

        btn_eliminar = tk.Button(top_bar, text="🗑 Eliminar", bg=self.colores["rojo_eliminar"], 
                            fg="white", font=("Arial", 10, "bold"), relief="flat", padx=10,
                                command=self.eliminar_asesor_bd)
        btn_eliminar.pack(side="right", padx=5)

        btn_editar = tk.Button(top_bar, text="✏ Editar", bg=self.colores["botones_menu"], 
                            fg=self.colores["oro"], font=("Arial", 10, "bold"), relief="flat", padx=10,
                            command=self.modal_editar_asesor)
        btn_editar.pack(side="right", padx=5)

        btn_agregar = tk.Button(top_bar, text="➕ Agregar Asesor", bg=self.colores["sidebar"], 
                                fg="white", font=("Arial", 10, "bold"), relief="flat", padx=10,
                                command=self.modal_agregar_asesor)
        btn_agregar.pack(side="right", padx=5)

        columnas = ("ID", "Usuario", "Nombre", "Correo", "Teléfono", "Especialidad", "Fecha Contratación", "Salario", "Activo")
        self.tabla_asesores = ttk.Treeview(panel, columns=columnas, show="headings", height=10)
        for col in columnas:
            self.tabla_asesores.heading(col, text=col)
            if col in ("Nombre", "Fecha Contratación"):
                self.tabla_asesores.column(col, anchor="w", width=150)
            elif col == "Salario":
                self.tabla_asesores.column(col, anchor="center", width=120)
            else:
                self.tabla_asesores.column(col, anchor="center", width=110)

        self.tabla_asesores.pack(fill="both", expand=True, padx=25, pady=(0,15))
        self.actualizar_tabla_asesores()

    def actualizar_tabla_asesores(self):
        if not hasattr(self, 'tabla_asesores'):
            return
        try:
            for item in self.tabla_asesores.get_children():
                self.tabla_asesores.delete(item)

            conn = obtener_conexion()
            cur = conn.cursor()
            cur.execute("SELECT a.id, u.username, a.nombre, a.correo, a.telefono, a.especialidad, a.fecha_contratacion, a.salario, a.activo FROM asesores a LEFT JOIN usuarios u ON a.usuario_id = u.id ORDER BY a.id")
            for fila in cur.fetchall():
                fila = list(fila)
                fila[7] = f"${fila[7]:,.2f}" if fila[7] is not None and fila[7] != 0 else ""
                fila[6] = fila[6] or ""
                self.tabla_asesores.insert("", "end", values=tuple(fila))
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los asesores: {e}")

    def modal_agregar_asesor(self):
        modal = tk.Toplevel(self.root)
        modal.title("Agregar Asesor")
        modal.geometry("520x520")
        modal.resizable(True, True)
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_force()
        modal.lift()

        tk.Label(modal, text="Crear Asesor / Vincular Usuario", font=("Arial", 12, "bold")).pack(pady=10)
        frame = tk.Frame(modal)
        frame.pack(padx=15, pady=5, fill="both", expand=True)
        frame.grid_columnconfigure(1, weight=1)

        tk.Label(frame, text="Usuario (login):").grid(row=0, column=0, sticky="w", pady=4)
        ent_user = tk.Entry(frame, width=35)
        ent_user.grid(row=0, column=1, pady=4, sticky="ew")

        tk.Label(frame, text="Contraseña (si crea cuenta):").grid(row=1, column=0, sticky="w")
        ent_pw = tk.Entry(frame, width=30, show='*')
        ent_pw.grid(row=1, column=1, pady=5)

        tk.Label(frame, text="Nombre completo:").grid(row=2, column=0, sticky="w")
        ent_nombre = tk.Entry(frame, width=30)
        ent_nombre.grid(row=2, column=1, pady=5)

        tk.Label(frame, text="Correo:").grid(row=3, column=0, sticky="w")
        ent_correo = tk.Entry(frame, width=30)
        ent_correo.grid(row=3, column=1, pady=5)

        tk.Label(frame, text="Teléfono:").grid(row=4, column=0, sticky="w")
        ent_tel = tk.Entry(frame, width=30)
        ent_tel.grid(row=4, column=1, pady=5)

        tk.Label(frame, text="Especialidad:").grid(row=5, column=0, sticky="w")
        ent_esp = tk.Entry(frame, width=30)
        ent_esp.grid(row=5, column=1, pady=5)

        tk.Label(frame, text="Fecha de contratación:").grid(row=6, column=0, sticky="w")
        ent_fecha = tk.Entry(frame, width=30)
        ent_fecha.grid(row=6, column=1, pady=5)

        tk.Label(frame, text="Salario mensual:").grid(row=7, column=0, sticky="w", pady=4)
        ent_salario = tk.Entry(frame, width=35)
        ent_salario.grid(row=7, column=1, pady=4, sticky="ew")

        def validar_fecha(texto):
            texto = texto.strip()
            if not texto:
                return False
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    datetime.datetime.strptime(texto, fmt)
                    return True
                except ValueError:
                    continue
            return False

        def validar_nombre_completo(texto):
            partes = [p for p in texto.split() if p.strip()]
            return len(partes) >= 2

        def validar_telefono(texto):
            digitos = re.sub(r"\D", "", texto)
            return len(digitos) == 11

        def validar_email(texto):
            return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", texto) is not None

        def guardar():
            usuario = str(ent_user.get()).strip()
            pw = str(ent_pw.get()).strip()
            nombre = str(ent_nombre.get()).strip()
            correo = str(ent_correo.get()).strip()
            tel = str(ent_tel.get()).strip()
            esp = str(ent_esp.get()).strip()
            fecha_contratacion = str(ent_fecha.get()).strip()
            salario_texto = str(ent_salario.get()).strip()
            salario = 0.0
            if salario_texto:
                try:
                    salario = float(salario_texto.replace('$', '').replace(',', '').strip())
                except ValueError:
                    messagebox.showwarning("Salario inválido", "Ingrese un salario numérico válido")
                    return

            if not usuario or not nombre:
                messagebox.showwarning("Datos incompletos", "Usuario y Nombre son obligatorios")
                return
            if not validar_nombre_completo(nombre):
                messagebox.showwarning("Nombre inválido", "Ingrese nombre completo y apellido")
                return
            if not validar_email(correo):
                messagebox.showwarning("Correo inválido", "Ingrese un correo válido como ejemplo@gmail.com")
                return
            if not validar_telefono(tel):
                messagebox.showwarning("Teléfono inválido", "Ingrese un teléfono con exactamente 11 dígitos")
                return
            if not validar_fecha(fecha_contratacion):
                messagebox.showwarning("Fecha inválida", "Ingrese una fecha válida en formato YYYY-MM-DD o DD/MM/YYYY")
                return

            try:
                conn = obtener_conexion()
                cur = conn.cursor()
                # comprobar usuario
                cur.execute('SELECT id FROM usuarios WHERE username=?', (usuario,))
                row = cur.fetchone()
                if row:
                    usuario_id = int(row[0])
                else:
                    # crear usuario con rol Asesor
                    pw_store = hash_password(pw) if pw else hash_password('default123')
                    cur.execute('INSERT INTO usuarios (username, password_hash, rol) VALUES (?,?,?)', (usuario, pw_store, 'Asesor'))
                    usuario_id = int(cur.lastrowid)

                cur.execute('INSERT INTO asesores (usuario_id, nombre, correo, telefono, especialidad, fecha_contratacion, salario) VALUES (?,?,?,?,?,?,?)', (usuario_id, nombre, correo, tel, esp, fecha_contratacion, salario))
                conn.commit()
                conn.close()
                registrar_accion(usuario, 'Crear Asesor', f'Nombre:{nombre}')
                modal.destroy()
                self.actualizar_tabla_asesores()
                messagebox.showinfo('OK','Asesor creado')
            except Exception as e:
                detalle = traceback.format_exc()
                print(detalle)
                messagebox.showerror('Error', f'No se pudo crear el asesor:\n{detalle}')

        btn_guardar = tk.Button(modal, text='Guardar', command=guardar, bg=self.colores['sidebar'], fg='white', font=('Arial', 10, 'bold'), relief='raised', padx=12, pady=5)
        btn_guardar.pack(pady=10, anchor='center')

    def modal_editar_asesor(self):
        sel = self.tabla_asesores.selection()
        if not sel:
            messagebox.showwarning('Seleccionar', 'Seleccione un asesor para editar')
            return
        vals = self.tabla_asesores.item(sel, 'values')
        asesor_id = vals[0]

        modal = tk.Toplevel(self.root)
        modal.title('Editar Asesor')
        modal.geometry('520x480')
        modal.resizable(True, True)
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_force()

        frame = tk.Frame(modal)
        frame.pack(padx=15, pady=10, fill='both', expand=True)
        frame.grid_columnconfigure(1, weight=1)

        tk.Label(frame, text='Nombre completo:').grid(row=0,column=0,sticky='w', pady=4)
        ent_nombre = tk.Entry(frame,width=35); ent_nombre.grid(row=0,column=1,pady=4, sticky='ew')

        tk.Label(frame, text='Correo:').grid(row=1,column=0,sticky='w')
        ent_correo = tk.Entry(frame,width=30); ent_correo.grid(row=1,column=1,pady=5)

        tk.Label(frame, text='Teléfono:').grid(row=2,column=0,sticky='w')
        ent_tel = tk.Entry(frame,width=30); ent_tel.grid(row=2,column=1,pady=5)

        tk.Label(frame, text='Especialidad:').grid(row=3,column=0,sticky='w')
        ent_esp = tk.Entry(frame,width=30); ent_esp.grid(row=3,column=1,pady=5)

        tk.Label(frame, text='Fecha de contratación:').grid(row=4,column=0,sticky='w')
        ent_fecha = tk.Entry(frame,width=30); ent_fecha.grid(row=4,column=1,pady=5)

        tk.Label(frame, text='Salario mensual:').grid(row=5,column=0,sticky='w', pady=4)
        ent_salario = tk.Entry(frame,width=35); ent_salario.grid(row=5,column=1,pady=4, sticky='ew')

        try:
            conn = obtener_conexion(); cur = conn.cursor()
            cur.execute('SELECT nombre, correo, telefono, especialidad, fecha_contratacion, salario FROM asesores WHERE id=?', (asesor_id,))
            row = cur.fetchone()
            conn.close()
            if row:
                ent_nombre.insert(0,row[0] or '')
                ent_correo.insert(0,row[1] or '')
                ent_tel.insert(0,row[2] or '')
                ent_esp.insert(0,row[3] or '')
                ent_fecha.insert(0,row[4] or '')
                ent_salario.insert(0, str(row[5]) if row[5] is not None else '')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo cargar asesor: {e}'); modal.destroy(); return

        def validar_fecha(texto):
            texto = texto.strip()
            if not texto:
                return False
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    datetime.datetime.strptime(texto, fmt)
                    return True
                except ValueError:
                    continue
            return False

        def validar_nombre_completo(texto):
            partes = [p for p in texto.split() if p.strip()]
            return len(partes) >= 2

        def validar_telefono(texto):
            digitos = re.sub(r"\D", "", texto)
            return len(digitos) == 11

        def validar_email(texto):
            return re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", texto) is not None

        def guardar_edicion():
            nombre = str(ent_nombre.get()).strip(); correo = str(ent_correo.get()).strip(); tel = str(ent_tel.get()).strip(); esp = str(ent_esp.get()).strip(); fecha_contratacion = str(ent_fecha.get()).strip(); salario_texto = str(ent_salario.get()).strip()
            salario = 0.0
            if salario_texto:
                try:
                    salario = float(salario_texto.replace('$','').replace(',','').strip())
                except ValueError:
                    messagebox.showwarning('Salario inválido', 'Ingrese un salario numérico válido')
                    return
            if not nombre:
                messagebox.showwarning('Datos incompletos', 'Nombre es obligatorio')
                return
            if not validar_nombre_completo(nombre):
                messagebox.showwarning('Nombre inválido', 'Ingrese nombre completo y apellido')
                return
            if not validar_email(correo):
                messagebox.showwarning('Correo inválido', 'Ingrese un correo válido como ejemplo@gmail.com')
                return
            if not validar_telefono(tel):
                messagebox.showwarning('Teléfono inválido', 'Ingrese un teléfono con exactamente 11 dígitos')
                return
            if not validar_fecha(fecha_contratacion):
                messagebox.showwarning('Fecha inválida', 'Ingrese una fecha válida en formato YYYY-MM-DD o DD/MM/YYYY')
                return
            try:
                conn = obtener_conexion(); cur = conn.cursor()
                cur.execute('UPDATE asesores SET nombre=?, correo=?, telefono=?, especialidad=?, fecha_contratacion=?, salario=? WHERE id=?', (nombre, correo, tel, esp, fecha_contratacion, salario, asesor_id))
                conn.commit(); conn.close()
                registrar_accion('system','Editar Asesor', f'id:{asesor_id}')
                modal.destroy(); self.actualizar_tabla_asesores(); messagebox.showinfo('OK','Asesor actualizado')
            except Exception as e:
                detalle = traceback.format_exc()
                print(detalle)
                messagebox.showerror('Error', f'No se pudo actualizar el asesor:\n{detalle}')

        btn_guardar = tk.Button(modal, text='Guardar', command=guardar_edicion, bg=self.colores['botones_menu'], fg=self.colores['oro'], font=('Arial', 10, 'bold'), relief='raised', padx=12, pady=5)
        btn_guardar.pack(pady=10)

    def eliminar_asesor_bd(self):
        sel = self.tabla_asesores.selection()
        if not sel:
            messagebox.showwarning('Seleccionar','Seleccione un asesor a eliminar')
            return
        vals = self.tabla_asesores.item(sel, 'values')
        asesor_id = vals[0]
        if not messagebox.askyesno('Confirmar','Eliminar asesor seleccionado?'):
            return
        try:
            conn = obtener_conexion(); cur = conn.cursor()
            cur.execute('DELETE FROM asesores WHERE id=?', (asesor_id,))
            conn.commit(); conn.close()
            registrar_accion('system','Eliminar Asesor', f'id:{asesor_id}')
            self.actualizar_tabla_asesores(); messagebox.showinfo('OK','Asesor eliminado')
        except Exception as e:
            messagebox.showerror('Error', f'No se pudo eliminar: {e}')

    def alternar_estado_servicio(self, event):
        """Alterna el estado nativo del cliente entre Activo e Inactivo con doble clic."""
        seleccion = self.tabla_clientes.selection()
        if not seleccion:
            return

        valores = self.tabla_clientes.item(seleccion, "values")
        id_cliente = valores[0]
        estado_actual = valores[7]  # Ajustado al nuevo índice en la tupla (Status es el índice 7)

        nuevo_estado = "Inactivo" if estado_actual == "Activo" else "Activo"

        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            cursor.execute("""
                UPDATE clientes 
                SET estado = ? 
                WHERE id = ?
            """, (nuevo_estado, id_cliente))
            
            conn.commit()
            conn.close()
            
            self.actualizar_tabla_clientes()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo modificar el estado del cliente: {e}")

    def modal_agregar_cliente(self):
        """Abre una ventana emergente para registrar manualmente un nuevo cliente con ID e Industria dinámicos."""
        modal = tk.Toplevel(self.root)
        modal.title("Agregar Nuevo Cliente")
        modal.geometry("440x510")
        modal.configure(bg=self.colores["fondo_main"])
        modal.transient(self.root)
        modal.grab_set()

        tk.Label(modal, text="Registro Manual de Cliente", font=("Arial", 12, "bold"), 
            bg=self.colores["fondo_main"], fg=self.colores["sidebar"]).pack(pady=15)

        frame_campos = tk.Frame(modal, bg=self.colores["fondo_main"])
        frame_campos.pack(fill="both", expand=True, padx=20)

        # 1. Nombre Cliente
        tk.Label(frame_campos, text="Nombre Cliente / Empresa:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        entry_nom = tk.Entry(frame_campos, width=25, relief="solid")
        entry_nom.grid(row=0, column=1, pady=5)

        # 2. Industria (Prefijos de Código dinámicos: 1-001, 2-001, 3-001)
        tk.Label(frame_campos, text="Tipo de Industria:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        combo_ind = ttk.Combobox(frame_campos, width=22, state="readonly", 
                                values=TIPOS_EMPRESA)
        combo_ind.grid(row=1, column=1, pady=5)
        combo_ind.current(0)  # Por defecto manufacturera

        # 3. Servicio del Cliente (Cargados del listado oficial del Landing Page)
        tk.Label(frame_campos, text="Servicio Solicitado:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        
        servicios_disponibles = [
            "Planificación Estratégica Empresarial",
            "Gestión Financiera, Costos y Tesorería",
            "Optimización de Procesos Administrativos y Operacionales",
            "Transformación Digital (Sistema Homologado)",
            "Outsourcing Contable y Gestión de Nómina",
            "Cumplimiento Fiscal, Tributario y Parafiscal",
            "Consultoría Legal Corporativa"
        ]

        combo_serv = ttk.Combobox(frame_campos, width=22, state="readonly", values=servicios_disponibles)
        combo_serv.grid(row=2, column=1, pady=5)
        combo_serv.current(0)

        # 4. Teléfono
        tk.Label(frame_campos, text="Teléfono de Contacto:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=3, column=0, sticky="w", pady=5)
        entry_tel = tk.Entry(frame_campos, width=25, relief="solid")
        entry_tel.grid(row=3, column=1, pady=5)

        # 5. Nombre Persona Contacto
        tk.Label(frame_campos, text="Nombre Persona Contacto:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=4, column=0, sticky="w", pady=5)
        entry_per_con = tk.Entry(frame_campos, width=25, relief="solid")
        entry_per_con.grid(row=4, column=1, pady=5)

        # 6. Correo Electrónico
        tk.Label(frame_campos, text="Correo:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=5, column=0, sticky="w", pady=5)
        entry_cor = tk.Entry(frame_campos, width=25, relief="solid")
        entry_cor.grid(row=5, column=1, pady=5)

        def guardar_nuevo():
            nom = entry_nom.get().strip()
            cor = entry_cor.get().strip() or "No registrado"
            tel = entry_tel.get().strip() or "No registrado"
            per_con = entry_per_con.get().strip() or nom
            industria_sel = combo_ind.get()
            servicio_sel = combo_serv.get()

            if not nom:
                messagebox.showerror("Error", "El nombre de la empresa es obligatorio.", parent=modal)
                return

            try:
                conn = obtener_conexion()
                cursor = conn.cursor()
                
                # Generar código de cliente según industria (centralizado)
                codigo = generar_id_cliente(industria_sel, conn)

                # Comprobación preventiva de campos en BD
                cursor.execute("PRAGMA table_info(clientes)")
                columnas = [col[1] for col in cursor.fetchall()]
                if "nombre_contacto" not in columnas:
                    cursor.execute("ALTER TABLE clientes ADD COLUMN nombre_contacto TEXT")
                    conn.commit()

                # Insertar en tabla local
                cursor.execute("""
                    INSERT INTO clientes (codigo, nombre, correo, telefono, industria, servicio, nombre_contacto, fecha_registro, estado)
                    VALUES (?, ?, ?, ?, ?, ?, ?, date('now'), 'Activo')
                """, (codigo, nom, cor, tel, industria_sel, servicio_sel, per_con))
                
                conn.commit()
                
                usuario_actual = getattr(self, 'usuario_autenticado', 'Sistema')
                registrar_accion(usuario_actual, "AGREGÓ CLIENTE", f"Registró a: {nom} [{industria_sel}] bajo el código {codigo}")

                conn.close()
                messagebox.showinfo("Éxito", f"Cliente guardado correctamente.\nCódigo Asignado: {codigo}", parent=modal)
                modal.destroy()
                self.actualizar_tabla_clientes()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar en la base de datos: {e}", parent=modal)

        tk.Button(modal, text="Guardar Cliente", bg=self.colores["verde_aprobar"], fg="white", 
                font=("Arial", 10, "bold"), relief="flat", command=guardar_nuevo).pack(pady=15)

    def modal_editar_cliente(self):
        """Abre una ventana emergente moderna para editar los datos del cliente seleccionado."""
        seleccion = self.tabla_clientes.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Por favor, seleccione un cliente de la lista para editar.")
            return

        valores = self.tabla_clientes.item(seleccion, "values")
        id_cliente = valores[0]
        codigo_act = valores[1]
        nombre_act = valores[2]
        servicio_act = valores[3]
        telefono_act = valores[4]
        per_con_act = valores[5]
        correo_act = valores[6]

        # Consultar datos actuales específicos de la base de datos
        industria_act = "Servicios"
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("SELECT industria FROM clientes WHERE id = ?", (id_cliente,))
            fila = cursor.fetchone()
            if fila:
                industria_act = fila[0]
            conn.close()
        except Exception as e:
            print(f"Error obteniendo detalles del cliente: {e}")

        modal = tk.Toplevel(self.root)
        modal.title("Editar Cliente")
        modal.geometry("440x510")
        modal.configure(bg=self.colores["fondo_main"])
        modal.transient(self.root)
        modal.grab_set()

        tk.Label(modal, text=f"Modificar Datos del Cliente ID: {id_cliente}\nCódigo actual: {codigo_act}", 
                 font=("Arial", 11, "bold"), bg=self.colores["fondo_main"], fg=self.colores["sidebar"]).pack(pady=15)

        frame_campos = tk.Frame(modal, bg=self.colores["fondo_main"])
        frame_campos.pack(fill="both", expand=True, padx=20)

        # 1. Nombre Cliente
        tk.Label(frame_campos, text="Nombre Cliente / Empresa:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        entry_nom = tk.Entry(frame_campos, width=25, relief="solid")
        entry_nom.insert(0, nombre_act)
        entry_nom.grid(row=0, column=1, pady=5)

        # 2. Tipo de Industria
        tk.Label(frame_campos, text="Tipo de Industria:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        combo_ind = ttk.Combobox(frame_campos, width=22, state="readonly", 
                            values=TIPOS_EMPRESA)
        combo_ind.grid(row=1, column=1, pady=5)
        
        if industria_act in combo_ind["values"]:
            combo_ind.set(industria_act)
        else:
            combo_ind.current(2)

        # 3. Servicio del Cliente
        tk.Label(frame_campos, text="Servicio Solicitado:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=2, column=0, sticky="w", pady=5)
        
        servicios_disponibles = [
            "Planificación Estratégica Empresarial",
            "Gestión Financiera, Costos y Tesorería",
            "Optimización de Procesos Administrativos y Operacionales",
            "Transformación Digital (Sistema Homologado)",
            "Outsourcing Contable y Gestión de Nómina",
            "Cumplimiento Fiscal, Tributario y Parafiscal",
            "Consultoría Legal Corporativa"
        ]

        combo_serv = ttk.Combobox(frame_campos, width=22, state="readonly", values=servicios_disponibles)
        combo_serv.grid(row=2, column=1, pady=5)
        
        if servicio_act in combo_serv["values"]:
            combo_serv.set(servicio_act)
        else:
            combo_serv.current(0)

        # 4. Teléfono
        tk.Label(frame_campos, text="Teléfono de Contacto:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=3, column=0, sticky="w", pady=5)
        entry_tel = tk.Entry(frame_campos, width=25, relief="solid")
        entry_tel.insert(0, telefono_act)
        entry_tel.grid(row=3, column=1, pady=5)

        # 5. Nombre Persona Contacto
        tk.Label(frame_campos, text="Nombre Persona Contacto:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=4, column=0, sticky="w", pady=5)
        entry_per_con = tk.Entry(frame_campos, width=25, relief="solid")
        entry_per_con.insert(0, per_con_act)
        entry_per_con.grid(row=4, column=1, pady=5)

        # 6. Correo Electrónico
        tk.Label(frame_campos, text="Correo:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=5, column=0, sticky="w", pady=5)
        entry_cor = tk.Entry(frame_campos, width=25, relief="solid")
        entry_cor.insert(0, correo_act)
        entry_cor.grid(row=5, column=1, pady=5)

        def actualizar_datos():
            nom = entry_nom.get().strip()
            cor = entry_cor.get().strip()
            tel = entry_tel.get().strip()
            per_con = entry_per_con.get().strip() or nom
            nueva_industria = combo_ind.get()
            nuevo_servicio = combo_serv.get()

            if not nom:
                messagebox.showerror("Error", "El nombre es obligatorio.", parent=modal)
                return

            try:
                conn = obtener_conexion()
                cursor = conn.cursor()
                
                # RECALCULAR ID AUTOMÁTICAMENTE SI CAMBIA LA INDUSTRIA
                if nueva_industria != industria_act:
                    nuevo_codigo = generar_id_cliente(nueva_industria, conn)
                else:
                    nuevo_codigo = codigo_act

                cursor.execute("""
                    UPDATE clientes 
                    SET nombre = ?, correo = ?, telefono = ?, industria = ?, servicio = ?, nombre_contacto = ?, codigo = ?
                    WHERE id = ?
                """, (nom, cor, tel, nueva_industria, nuevo_servicio, per_con, nuevo_codigo, id_cliente))
                
                conn.commit()
                
                usuario_actual = getattr(self, 'usuario_autenticado', 'Sistema')
                registrar_accion(usuario_actual, "EDITÓ CLIENTE", f"Modificó datos de {nom} (Nuevo Código: {nuevo_codigo})")
                
                conn.close()
                messagebox.showinfo("Éxito", f"Cliente actualizado con éxito.\nCódigo final: {nuevo_codigo}", parent=modal)
                modal.destroy()
                self.actualizar_tabla_clientes()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo actualizar: {e}", parent=modal)

        tk.Button(modal, text="Guardar Cambios", bg=self.colores["sidebar"], fg="white", 
                font=("Arial", 10, "bold"), relief="flat", command=actualizar_datos).pack(pady=15)

    def eliminar_cliente_bd(self):
        """Elimina de forma segura un cliente y sus registros financieros."""
        seleccion = self.tabla_clientes.selection()
        if not seleccion:
            messagebox.showwarning("Atención", "Por favor, seleccione un cliente del directorio para eliminar.")
            return

        valores = self.tabla_clientes.item(seleccion, "values")
        id_cliente = valores[0]
        nombre_cliente = valores[2] # Ajustado al índice 2 para el Nombre en el nuevo esquema

        confirmacion = messagebox.askyesno("Confirmación de Seguridad", 
                                            f"¿Está completamente seguro de que desea eliminar al cliente '{nombre_cliente}'?\n\nEsta acción borrará también sus registros financieros asociados de forma permanente.")
        if not confirmacion:
            return

        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            
            cursor.execute("DELETE FROM solicitudes_servicio WHERE cliente_id = ?", (id_cliente,))
            cursor.execute("DELETE FROM clientes WHERE id = ?", (id_cliente,))
            
            conn.commit()
            
            usuario_actual = getattr(self, 'usuario_autenticado', 'Sistema')
            registrar_accion(usuario_actual, "ELIMINÓ CLIENTE", f"Removió de forma permanente a: {nombre_cliente} (ID: {id_cliente})")
            
            conn.close()
            messagebox.showinfo("Éxito", "El cliente y su historial financiero han sido removidos del sistema.")
            self.actualizar_tabla_clientes()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo completar la transacción de borrado: {e}")
            
    def mostrar_servicios(self):
        panel = self.crear_panel_principal("Catálogo de Servicios y Precios Base")

        tk.Label(
            panel,
            text="Los precios base se utilizan automáticamente al aprobar una asesoría web.",
            bg=self.colores["fondo_main"],
            fg=self.colores["texto_oscuro"],
            font=("Arial", 10, "italic")
        ).pack(anchor="w", padx=25, pady=(0, 8))

        top_bar = tk.Frame(panel, bg=self.colores["fondo_main"])
        top_bar.pack(fill="x", padx=25, pady=(0, 10))

        tabla_s = None

        def servicio_seleccionado():
            seleccion = tabla_s.selection()
            if not seleccion:
                messagebox.showwarning(
                    "Seleccione un servicio",
                    "Primero seleccione un servicio de la tabla."
                )
                return None
            return tabla_s.item(seleccion[0], "values")

        def editar_servicio():
            valores = servicio_seleccionado()
            if not valores:
                return

            servicio_id = int(valores[0])

            try:
                conn = obtener_conexion()
                cur = conn.cursor()
                cur.execute(
                    "SELECT nombre_servicio, descripcion, costo_base FROM servicios WHERE id = ?",
                    (servicio_id,)
                )
                fila = cur.fetchone()
                conn.close()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el servicio: {e}")
                return

            if not fila:
                messagebox.showerror("Error", "El servicio seleccionado ya no existe.")
                return

            modal = tk.Toplevel(self.root)
            modal.title("Editar Servicio")
            modal.geometry("560x390")
            modal.configure(bg=self.colores["fondo_main"])
            modal.transient(self.root)
            modal.grab_set()
            modal.resizable(False, False)

            tk.Label(
                modal,
                text="Editar servicio y precio base",
                bg=self.colores["fondo_main"],
                fg=self.colores["sidebar"],
                font=("Arial", 14, "bold")
            ).pack(pady=(20, 14))

            frame = tk.Frame(modal, bg=self.colores["fondo_main"])
            frame.pack(fill="both", expand=True, padx=25)

            tk.Label(frame, text="Nombre:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="nw", pady=7)
            ent_nombre = tk.Entry(frame, width=48)
            ent_nombre.insert(0, fila[0] or "")
            ent_nombre.grid(row=0, column=1, sticky="ew", pady=7)

            tk.Label(frame, text="Descripción:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="nw", pady=7)
            txt_desc = tk.Text(frame, width=46, height=7, wrap="word")
            txt_desc.insert("1.0", fila[1] or "")
            txt_desc.grid(row=1, column=1, sticky="ew", pady=7)

            tk.Label(frame, text="Precio base (USD):", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=2, column=0, sticky="w", pady=7)
            ent_costo = tk.Entry(frame, width=20)
            ent_costo.insert(0, f"{float(fila[2]):.2f}")
            ent_costo.grid(row=2, column=1, sticky="w", pady=7)

            def guardar_edicion():
                nombre = ent_nombre.get().strip()
                descripcion = txt_desc.get("1.0", "end").strip()
                costo_txt = ent_costo.get().strip().replace(",", ".")

                if not nombre or not descripcion or not costo_txt:
                    messagebox.showwarning("Datos incompletos", "Complete nombre, descripción y precio.")
                    return

                try:
                    costo = float(costo_txt)
                    if costo <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showwarning("Precio inválido", "Ingrese un precio base mayor que cero.")
                    return

                try:
                    conn = obtener_conexion()
                    cur = conn.cursor()
                    cur.execute(
                        """
                        SELECT id FROM servicios
                        WHERE LOWER(TRIM(nombre_servicio)) = LOWER(TRIM(?))
                          AND id <> ?
                        """,
                        (nombre, servicio_id)
                    )
                    if cur.fetchone():
                        conn.close()
                        messagebox.showwarning("Servicio duplicado", "Ya existe otro servicio con ese nombre.")
                        return

                    cur.execute(
                        """
                        UPDATE servicios
                        SET nombre_servicio = ?, descripcion = ?, costo_base = ?
                        WHERE id = ?
                        """,
                        (nombre, descripcion, costo, servicio_id)
                    )
                    conn.commit()
                    conn.close()

                    registrar_accion(
                        getattr(self, "usuario_autenticado", "Sistema"),
                        "EDITÓ SERVICIO",
                        f"Actualizó '{nombre}' con precio base ${costo:,.2f}."
                    )

                    messagebox.showinfo("Éxito", "Servicio actualizado correctamente.")
                    modal.destroy()
                    self.mostrar_servicios()

                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo actualizar el servicio: {e}")

            tk.Button(
                modal,
                text="Guardar cambios",
                bg=self.colores["verde_aprobar"],
                fg="white",
                font=("Arial", 10, "bold"),
                relief="flat",
                padx=18,
                pady=8,
                command=guardar_edicion
            ).pack(pady=16)

        def eliminar_servicio_seleccionado():
            valores = servicio_seleccionado()
            if not valores:
                return

            servicio_id = int(valores[0])
            nombre = valores[1]

            if not messagebox.askyesno(
                "Confirmar eliminación",
                f"¿Desea eliminar el servicio?\n\n{nombre}"
            ):
                return

            try:
                conn = obtener_conexion()
                cur = conn.cursor()

                cur.execute(
                    "SELECT COUNT(*) FROM solicitudes_servicio WHERE servicio_id = ?",
                    (servicio_id,)
                )
                usos = int(cur.fetchone()[0])

                if usos > 0:
                    conn.close()
                    messagebox.showwarning(
                        "Servicio en uso",
                        "Este servicio ya posee movimientos asociados y no puede eliminarse.\n"
                        "Puede editar su nombre, descripción o precio base."
                    )
                    return

                cur.execute("DELETE FROM servicios WHERE id = ?", (servicio_id,))
                conn.commit()
                conn.close()

                registrar_accion(
                    getattr(self, "usuario_autenticado", "Sistema"),
                    "ELIMINÓ SERVICIO",
                    f"Eliminó el servicio '{nombre}'."
                )

                messagebox.showinfo("Éxito", "Servicio eliminado correctamente.")
                self.mostrar_servicios()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo eliminar el servicio: {e}")

        tk.Button(
            top_bar,
            text="➕ Agregar Servicio",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            pady=6,
            command=self.agregar_servicio
        ).pack(side="left")

        tk.Button(
            top_bar,
            text="✏ Editar Servicio",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            pady=6,
            command=editar_servicio
        ).pack(side="left", padx=8)

        tk.Button(
            top_bar,
            text="🗑 Eliminar Servicio",
            bg=self.colores["rojo_eliminar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            pady=6,
            command=eliminar_servicio_seleccionado
        ).pack(side="left")

        columnas = ("ID", "Nombre Servicio", "Descripción", "Precio Base USD")
        tabla_s = ttk.Treeview(panel, columns=columnas, show="headings", height=15)

        tabla_s.heading("ID", text="ID")
        tabla_s.heading("Nombre Servicio", text="Servicio")
        tabla_s.heading("Descripción", text="Descripción")
        tabla_s.heading("Precio Base USD", text="Precio Base")

        tabla_s.column("ID", width=55, anchor="center", stretch=False)
        tabla_s.column("Nombre Servicio", width=270, anchor="w")
        tabla_s.column("Descripción", width=520, anchor="w")
        tabla_s.column("Precio Base USD", width=130, anchor="e", stretch=False)

        scroll_y = ttk.Scrollbar(panel, orient="vertical", command=tabla_s.yview)
        tabla_s.configure(yscrollcommand=scroll_y.set)

        scroll_y.pack(side="right", fill="y", pady=10, padx=(0, 20))
        tabla_s.pack(fill="both", expand=True, padx=(25, 0), pady=10)

        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, nombre_servicio, descripcion, costo_base
                FROM servicios
                ORDER BY nombre_servicio COLLATE NOCASE
                """
            )
            for fila in cursor.fetchall():
                tabla_s.insert(
                    "",
                    "end",
                    values=(
                        fila[0],
                        fila[1],
                        fila[2],
                        f"$ {float(fila[3]):,.2f}"
                    )
                )
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el catálogo: {e}")

    def agregar_servicio(self):
        modal = tk.Toplevel(self.root)
        modal.title("Agregar Servicio")
        modal.geometry("560x390")
        modal.configure(bg=self.colores["fondo_main"])
        modal.transient(self.root)
        modal.grab_set()
        modal.resizable(False, False)

        tk.Label(
            modal,
            text="Agregar nuevo servicio",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 14, "bold")
        ).pack(pady=(20, 14))

        frame = tk.Frame(modal, bg=self.colores["fondo_main"])
        frame.pack(fill="both", expand=True, padx=25)

        tk.Label(frame, text="Nombre:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="nw", pady=7)
        entry_nombre = tk.Entry(frame, width=48)
        entry_nombre.grid(row=0, column=1, pady=7)

        tk.Label(frame, text="Descripción:", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="nw", pady=7)
        txt_desc = tk.Text(frame, width=46, height=7, wrap="word")
        txt_desc.grid(row=1, column=1, pady=7)

        tk.Label(frame, text="Precio base (USD):", bg=self.colores["fondo_main"], font=("Arial", 9, "bold")).grid(row=2, column=0, sticky="w", pady=7)
        entry_costo = tk.Entry(frame, width=20)
        entry_costo.grid(row=2, column=1, sticky="w", pady=7)

        def guardar():
            nombre = entry_nombre.get().strip()
            desc = txt_desc.get("1.0", "end").strip()
            costo_txt = entry_costo.get().strip().replace(",", ".")

            if not nombre or not desc or not costo_txt:
                messagebox.showwarning("Atención", "Complete todos los campos.")
                return

            try:
                costo = float(costo_txt)
                if costo <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Precio inválido", "Ingrese un precio base mayor que cero.")
                return

            try:
                conn = obtener_conexion()
                cursor = conn.cursor()

                cursor.execute(
                    """
                    SELECT id FROM servicios
                    WHERE LOWER(TRIM(nombre_servicio)) = LOWER(TRIM(?))
                    """,
                    (nombre,)
                )
                if cursor.fetchone():
                    conn.close()
                    messagebox.showwarning("Servicio duplicado", "Ya existe un servicio con ese nombre.")
                    return

                cursor.execute(
                    """
                    INSERT INTO servicios (nombre_servicio, descripcion, costo_base)
                    VALUES (?, ?, ?)
                    """,
                    (nombre, desc, costo)
                )
                conn.commit()
                conn.close()

                registrar_accion(
                    getattr(self, "usuario_autenticado", "Sistema"),
                    "AGREGÓ SERVICIO",
                    f"Agregó '{nombre}' con precio base ${costo:,.2f}."
                )

                messagebox.showinfo("Éxito", "Servicio agregado correctamente.")
                modal.destroy()
                self.mostrar_servicios()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo agregar el servicio: {e}")

        tk.Button(
            modal,
            text="Guardar Servicio",
            bg=self.colores["verde_aprobar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=18,
            pady=8,
            command=guardar
        ).pack(pady=16)

    def mostrar_notificaciones(self):
        panel = self.crear_panel_principal("Solicitudes y Citas Web")

        tk.Label(
            panel,
            text="Bandeja activa de formularios recibidos desde la página web.",
            bg=self.colores["fondo_main"],
            font=("Arial", 10, "italic")
        ).pack(anchor="w", padx=25, pady=(0, 8))

        self._mostrar_alerta_servicios_nuevos(panel)

        barra = tk.Frame(panel, bg=self.colores["fondo_main"])
        barra.pack(fill="x", padx=25, pady=(0, 8))

        tk.Button(
            barra,
            text="📞 Marcar Contactado",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            pady=7,
            command=lambda: self.marcar_solicitud_contactada(tabla_web)
        ).pack(side="left")

        tk.Button(
            barra,
            text="🕘 Historial de Solicitudes",
            bg=self.colores["botones_menu"],
            fg=self.colores["oro"],
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=12,
            pady=7,
            command=self.mostrar_historial_solicitudes
        ).pack(side="right")

        columnas_w = (
            "ID",
            "Cliente Potencial",
            "Servicio de Interés",
            "Descripción del Negocio",
            "Estado",
            "Correo",
            "Teléfono"
        )

        tabla_web = ttk.Treeview(
            panel,
            columns=columnas_w,
            show="headings",
            height=12
        )

        for col in columnas_w:
            tabla_web.heading(col, text=col)

            if col == "Descripción del Negocio":
                tabla_web.column(col, anchor="w", width=280)
            elif col in ("Correo", "Teléfono"):
                tabla_web.column(col, anchor="w", width=160)
            elif col == "ID":
                tabla_web.column(col, anchor="center", width=85)
            else:
                tabla_web.column(col, anchor="center", width=145)

        tabla_web.pack(fill="both", expand=True, padx=25, pady=10)

        frame_botones = tk.Frame(panel, bg=self.colores["fondo_main"])
        frame_botones.pack(fill="x", padx=25, pady=10)

        btn_aprobar = tk.Button(
            frame_botones,
            text="✔ Aprobar Asesoría",
            bg=self.colores["verde_aprobar"],
            fg="white",
            font=("Arial", 11, "bold"),
            relief="flat",
            padx=15,
            pady=8,
            command=lambda: self.aprobar_y_calcular_solicitud(tabla_web)
        )
        btn_aprobar.pack(side="left")

        btn_rechazar = tk.Button(
            frame_botones,
            text="✖ Rechazar Servicio",
            bg="#9e2a2b",
            fg="white",
            font=("Arial", 11, "bold"),
            relief="flat",
            padx=15,
            pady=8,
            command=lambda: self.rechazar_solicitud(tabla_web)
        )
        btn_rechazar.pack(side="left", padx=8)

        try:
            data = self._api_solicitudes("GET", "/api/solicitudes")

            for fila in data.get("items", []):
                estado = fila.get("estado") or "Pendiente"

                # Pendiente y Contactado siguen siendo solicitudes activas.
                if estado.lower() not in {"pendiente", "contactado"}:
                    continue

                tabla_web.insert(
                    "",
                    "end",
                    values=(
                        f"ID Web: {fila.get('id')}",
                        fila.get("cliente_potencial") or "No registrado",
                        fila.get("servicio_interes") or "No especificado",
                        fila.get("descripcion") or "",
                        estado,
                        fila.get("correo") or "No registrado",
                        fila.get("telefono") or "No registrado"
                    )
                )

        except Exception as e:
            tabla_web.insert(
                "",
                "end",
                values=("Error", "Servidor web", str(e), "", "", "", "")
            )

    def marcar_solicitud_contactada(self, tabla):
        seleccion = tabla.selection()

        if not seleccion:
            messagebox.showwarning(
                "Atención",
                "Seleccione una solicitud."
            )
            return

        valores = tabla.item(seleccion[0], "values")
        estado_actual = str(valores[4] or "").strip()

        if estado_actual.lower() == "contactado":
            messagebox.showinfo(
                "Información",
                "Esta solicitud ya está marcada como Contactado."
            )
            return

        if estado_actual.lower() != "pendiente":
            messagebox.showinfo(
                "Solicitud procesada",
                "Esta solicitud ya no está pendiente."
            )
            return

        try:
            id_web = int(str(valores[0]).replace("ID Web: ", ""))

            self._api_solicitudes(
                "PATCH",
                f"/api/solicitudes/{id_web}",
                json={"estado": "Contactado"}
            )

            registrar_accion(
                getattr(self, "usuario_autenticado", "Sistema"),
                "CONTACTÓ SOLICITUD WEB",
                f"Marcó como Contactado el formulario web ID {id_web}."
            )

            messagebox.showinfo(
                "Contacto registrado",
                "La solicitud fue marcada como Contactado.\n"
                "Ahora puede aprobarse o rechazarse cuando corresponda."
            )

            self.mostrar_notificaciones()

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo actualizar la solicitud: {e}"
            )

    def mostrar_historial_solicitudes(self):
        panel = self.crear_panel_principal("Historial de Solicitudes Web")

        tk.Label(
            panel,
            text="Consulta de solicitudes procesadas y su trazabilidad.",
            bg=self.colores["fondo_main"],
            font=("Arial", 10, "italic")
        ).pack(anchor="w", padx=25, pady=(0, 10))

        filtros = tk.Frame(panel, bg=self.colores["fondo_main"])
        filtros.pack(fill="x", padx=25, pady=(0, 10))

        tk.Label(
            filtros,
            text="Estado:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        ).pack(side="left")

        combo_estado = ttk.Combobox(
            filtros,
            state="readonly",
            width=18,
            values=["Todos", "Pendiente", "Contactado", "Aprobado", "Rechazado"]
        )
        combo_estado.set("Todos")
        combo_estado.pack(side="left", padx=(6, 15))

        tk.Label(
            filtros,
            text="Buscar:",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 10, "bold")
        ).pack(side="left")

        entrada_buscar = tk.Entry(filtros, width=28)
        entrada_buscar.pack(side="left", padx=6, ipady=3)

        columnas = (
            "ID",
            "Fecha solicitud",
            "Cliente",
            "Servicio",
            "Estado",
            "Procesado por",
            "Última actualización",
            "Correo",
            "Teléfono"
        )

        tabla = ttk.Treeview(
            panel,
            columns=columnas,
            show="headings",
            height=16
        )

        for col in columnas:
            tabla.heading(col, text=col)

            if col in ("Cliente", "Servicio", "Correo"):
                tabla.column(col, width=180, anchor="w")
            elif col == "ID":
                tabla.column(col, width=60, anchor="center")
            else:
                tabla.column(col, width=125, anchor="center")

        sb = ttk.Scrollbar(panel, orient="vertical", command=tabla.yview)
        tabla.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y", padx=(0, 18), pady=5)
        tabla.pack(fill="both", expand=True, padx=(25, 0), pady=5)

        def cargar():
            for item in tabla.get_children():
                tabla.delete(item)

            try:
                data = self._api_solicitudes("GET", "/api/solicitudes")
                estado_filtro = combo_estado.get().strip()
                texto = entrada_buscar.get().strip().lower()

                for fila in data.get("items", []):
                    estado = fila.get("estado") or "Pendiente"

                    if estado_filtro != "Todos" and estado != estado_filtro:
                        continue

                    cadena = " ".join([
                        str(fila.get("cliente_potencial") or ""),
                        str(fila.get("servicio_interes") or ""),
                        str(fila.get("correo") or ""),
                        str(fila.get("telefono") or ""),
                        str(fila.get("procesado_por") or "")
                    ]).lower()

                    if texto and texto not in cadena:
                        continue

                    tabla.insert(
                        "",
                        "end",
                        values=(
                            fila.get("id"),
                            fila.get("fecha_solicitud") or "",
                            fila.get("cliente_potencial") or "",
                            fila.get("servicio_interes") or "",
                            estado,
                            fila.get("procesado_por") or "—",
                            fila.get("actualizado_en") or "—",
                            fila.get("correo") or "",
                            fila.get("telefono") or ""
                        )
                    )

            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"No se pudo cargar el historial: {e}"
                )

        combo_estado.bind("<<ComboboxSelected>>", lambda e: cargar())
        entrada_buscar.bind("<KeyRelease>", lambda e: cargar())

        tk.Button(
            filtros,
            text="↻ Actualizar",
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            padx=10,
            command=cargar
        ).pack(side="right")

        cargar()

    def rechazar_solicitud(self, tabla):
        item_seleccionado = tabla.selection()

        if not item_seleccionado:
            messagebox.showwarning(
                "Atención",
                "Seleccione una solicitud para rechazarla."
            )
            return

        valores = tabla.item(item_seleccionado, "values")
        estado_actual = valores[4]

        # Evitar procesar una solicitud más de una vez.
        if str(estado_actual).lower() not in {"pendiente", "contactado"}:
            messagebox.showinfo(
                "Solicitud procesada",
                "Esta solicitud ya fue procesada."
            )
            return

        ref_id = valores[0]
        id_web_real = ref_id.replace("ID Web: ", "")

        confirmacion = messagebox.askyesno(
            "Confirmar",
            "¿Desea rechazar esta solicitud?"
        )
        if not confirmacion:
            return

        try:
            self._api_solicitudes(
                "PATCH",
                f"/api/solicitudes/{int(id_web_real)}",
                json={"estado": "Rechazado"}
            )

            usuario_actual = getattr(self, "usuario_autenticado", "Sistema")
            registrar_accion(
                usuario_actual,
                "RECHAZÓ SOLICITUD WEB",
                f"Rechazó la solicitud web ID {id_web_real}."
            )

            messagebox.showinfo(
                "Éxito",
                "Solicitud rechazada correctamente."
            )

            # Al recargar, ya no aparecerá porque dejó de estar Pendiente.
            self.mostrar_notificaciones()

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"No se pudo rechazar la solicitud: {e}"
            )

    def _mostrar_alerta_servicios_nuevos(self, panel):
        try:
            conn = obtener_conexion()
            cursor = conn.cursor()
            cursor.execute("SELECT id, nombre_servicio FROM servicios ORDER BY id DESC LIMIT 5")
            servicios = cursor.fetchall()
            conn.close()
        except Exception:
            servicios = []

        if not servicios:
            return

        frame_alerta = tk.Frame(panel, bg="#fff3cd", highlightbackground="#f0b429", highlightthickness=1)
        frame_alerta.pack(fill="x", padx=25, pady=(0, 10))

        tk.Label(frame_alerta, text="📣 Nuevos servicios en el catálogo", bg="#fff3cd", fg="#7a4b00", font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(8, 2))
        for servicio_id, nombre in servicios:
            tk.Label(frame_alerta, text=f"• {nombre}", bg="#fff3cd", fg="#3b2f2f", font=("Arial", 10)).pack(anchor="w", padx=14, pady=1)

    def seleccionar_metodo_pago(self):
        ventana = tk.Toplevel(self.root)
        ventana.title("Método de pago")
        ventana.geometry("450x220")
        ventana.resizable(False, False)
        ventana.transient(self.root)
        ventana.grab_set()

        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 225
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 110
        ventana.geometry(f"450x220+{x}+{y}")

        tk.Label(ventana, text="Seleccione el método de pago del cliente:", font=("Arial", 11, "bold")).pack(pady=(18, 12))

        resultado = {"valor": None}

        def elegir(valor):
            resultado["valor"] = valor
            ventana.destroy()

        frame_botones = tk.Frame(ventana)
        frame_botones.pack(pady=8)

        tk.Button(frame_botones, text="Transferencia", width=16, bg="#6b1426", fg="white", font=("Arial", 10, "bold"), relief="flat", command=lambda: elegir("transferencia")).pack(side="left", padx=8)
        tk.Button(frame_botones, text="Efectivo", width=16, bg="#2a9d8f", fg="white", font=("Arial", 10, "bold"), relief="flat", command=lambda: elegir("efectivo")).pack(side="left", padx=8)
        tk.Button(frame_botones, text="Pago Móvil", width=16, bg="#d4af37", fg="black", font=("Arial", 10, "bold"), relief="flat", command=lambda: elegir("pago movil")).pack(side="left", padx=8)

        ventana.wait_window()
        return resultado["valor"]

    def seleccionar_industria_cliente(self):
        """Muestra un selector único y ordenado para clasificar la empresa del cliente."""
        ventana = tk.Toplevel(self.root)
        ventana.title("Clasificar empresa")
        ventana.geometry("520x260")
        ventana.configure(bg=self.colores["fondo_main"])
        ventana.transient(self.root)
        ventana.grab_set()
        ventana.resizable(False, False)

        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 260
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 130
        ventana.geometry(f"520x260+{x}+{y}")

        tk.Label(
            ventana,
            text="Clasificación de la empresa",
            bg=self.colores["fondo_main"],
            fg=self.colores["sidebar"],
            font=("Arial", 15, "bold")
        ).pack(pady=(24, 5))

        tk.Label(
            ventana,
            text="Seleccione una sola actividad o sector principal del cliente.",
            bg=self.colores["fondo_main"],
            fg=self.colores["texto_oscuro"],
            font=("Arial", 10)
        ).pack(pady=(0, 18))

        combo = ttk.Combobox(
            ventana,
            state="readonly",
            width=42,
            values=TIPOS_EMPRESA,
            font=("Arial", 10)
        )
        combo.pack(ipady=4)
        combo.current(0)

        resultado = {"valor": None}

        def confirmar():
            valor = combo.get().strip()
            if not valor:
                messagebox.showwarning(
                    "Clasificación requerida",
                    "Seleccione el tipo de empresa.",
                    parent=ventana
                )
                return
            resultado["valor"] = valor
            ventana.destroy()

        frame_botones = tk.Frame(ventana, bg=self.colores["fondo_main"])
        frame_botones.pack(pady=22)

        tk.Button(
            frame_botones,
            text="Confirmar",
            width=15,
            bg=self.colores["sidebar"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            command=confirmar
        ).pack(side="left", padx=6)

        tk.Button(
            frame_botones,
            text="Cancelar",
            width=15,
            bg=self.colores["botones_menu"],
            fg="white",
            font=("Arial", 10, "bold"),
            relief="flat",
            command=ventana.destroy
        ).pack(side="left", padx=6)

        ventana.wait_window()
        return resultado["valor"]

    def aprobar_y_calcular_solicitud(self, tabla):
        item_seleccionado = tabla.selection()
        if not item_seleccionado:
            messagebox.showwarning("Atención", "Por favor, seleccione una solicitud de la lista.")
            return
            
        valores = tabla.item(item_seleccionado, "values")
        ref_id, cliente_nombre, servicio_interes, descripcion_negocio, estado_actual, correo_real, telefono_real = valores

        if str(estado_actual).lower() not in {"pendiente", "contactado"}:
            messagebox.showinfo(
                "Solicitud procesada",
                "Esta solicitud ya fue procesada."
            )
            return

        id_web_real = ref_id.replace("ID Web: ", "")

        industria_cliente = self.seleccionar_industria_cliente()
        if not industria_cliente:
            return

        tasa_bcv = obtener_tasa_bcv_automatica()
        if not tasa_bcv:
            tasa_bcv_input = simpledialog.askstring("Tasa de Cambio", "Ingrese tasa BCV manual:")
            if not tasa_bcv_input: return
            tasa_bcv = float(tasa_bcv_input)

        metodo_pago = self.seleccionar_metodo_pago()
        if not metodo_pago:
            return

        try:
            conn_local = obtener_conexion()
            cursor_local = conn_local.cursor()

            codigo_cliente = generar_id_cliente(industria_cliente, conn_local)
            nombre_contacto = simpledialog.askstring(
                "Persona de contacto",
                "Nombre de la persona de contacto del cliente:",
                initialvalue=cliente_nombre
            )
            if not nombre_contacto:
                conn_local.close()
                return

            cursor_local.execute("""
                INSERT INTO clientes (codigo, nombre, correo, telefono, industria, servicio, nombre_contacto, fecha_registro, estado)
                VALUES (?, ?, ?, ?, ?, ?, ?, date('now'), 'Activo')
            """, (codigo_cliente, cliente_nombre, correo_real, telefono_real, industria_cliente, servicio_interes, nombre_contacto))
            
            cliente_id = cursor_local.lastrowid

            # Resolver el servicio exacto del catálogo.
            # Se mantiene compatibilidad con el antiguo texto "Sistema Holgado".
            aliases_servicio = {
                "Transformación Digital (Sistema Holgado)": "Transformación Digital (Sistema Homologado)",
            }
            servicio_catalogo = aliases_servicio.get(servicio_interes, servicio_interes)

            cursor_local.execute(
                """
                SELECT id, costo_base
                FROM servicios
                WHERE LOWER(TRIM(nombre_servicio)) = LOWER(TRIM(?))
                """,
                (servicio_catalogo,)
            )
            res_serv = cursor_local.fetchone()

            # Si llegó "Otro servicio" desde la web, el asesor define el precio
            # una sola vez y el servicio queda incorporado al catálogo.
            if not res_serv:
                precio_personalizado = simpledialog.askstring(
                    "Nuevo servicio detectado",
                    (
                        f"El servicio solicitado no existe todavía en el catálogo:\n\n"
                        f"{servicio_interes}\n\n"
                        "Ingrese su precio base en USD para incorporarlo:"
                    )
                )
                if not precio_personalizado:
                    conn_local.close()
                    return

                try:
                    monto_base_usd = float(
                        precio_personalizado.strip().replace(",", ".")
                    )
                    if monto_base_usd <= 0:
                        raise ValueError
                except ValueError:
                    conn_local.close()
                    messagebox.showwarning(
                        "Precio inválido",
                        "El precio base debe ser un número mayor que cero."
                    )
                    return

                cursor_local.execute(
                    """
                    INSERT INTO servicios (nombre_servicio, descripcion, costo_base)
                    VALUES (?, ?, ?)
                    """,
                    (
                        servicio_interes,
                        "Servicio incorporado desde una solicitud personalizada de la página web.",
                        monto_base_usd
                    )
                )
                servicio_id = cursor_local.lastrowid

            else:
                servicio_id = int(res_serv[0])
                monto_base_usd = float(res_serv[1])

            config_fin = self._obtener_configuracion_financiera()
            iva_tasa = config_fin["iva_pct"] / 100.0
            igtf_tasa = config_fin["igtf_pct"] / 100.0

            iva_usd = monto_base_usd * iva_tasa
            igtf_usd = (
                monto_base_usd * igtf_tasa
                if self._aplica_igtf(metodo_pago, config_fin)
                else 0.0
            )

            total_usd = monto_base_usd + iva_usd + igtf_usd
            total_bs = total_usd * tasa_bcv

            cursor_local.execute("""
                INSERT INTO solicitudes_servicio (
                    cliente_id, servicio_id, fecha_solicitud, monto,
                    iva_usd, igtf_usd, total_usd, tasa_bcv, total_bs,
                    metodo_pago, activo, estado_pago
                ) VALUES (?, ?, date('now'), ?, ?, ?, ?, ?, ?, ?, 'Activo', 'Pendiente')
            """, (
                cliente_id, servicio_id, monto_base_usd, iva_usd, igtf_usd,
                total_usd, tasa_bcv, total_bs, metodo_pago
            ))

            conn_local.commit()
            conn_local.close()

            # Marcar la solicitud web como aprobada en el servidor central.
            self._api_solicitudes(
                "PATCH",
                f"/api/solicitudes/{int(id_web_real)}",
                json={"estado": "Aprobado"}
            )

            usuario_actual = getattr(self, 'usuario_autenticado', 'Sistema')
            registrar_accion(usuario_actual, "APROBÓ SERVICIO WEB", f"Aprobó orden para {cliente_nombre}. Servicio: {servicio_interes}. Total: {total_bs:,.2f} Bs.")

            messagebox.showinfo("Éxito", f"¡Sincronización Completada!\n\nTasa BCV: {tasa_bcv:.4f} Bs.\nTotal: {total_bs:,.2f} Bs.")
            self.mostrar_notificaciones()

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo completar el proceso: {e}")
            
    def mostrar_historial(self):
        """Renderiza la pantalla de auditoría interna de la firma."""
        panel = self.crear_panel_principal("Historial de Actividades y Auditoría")
        
        tk.Label(panel, text="Registro cronológico inmutable de acciones del personal corporativo:", 
                bg=self.colores["fondo_main"], font=("Arial", 10, "italic")).pack(anchor="w", padx=25, pady=(0, 10))

        columnas_h = ("ID Registro", "Fecha y Hora", "Asesor / Usuario", "Operación Realizada", "Detalles del Movimiento")
        tabla_historial = ttk.Treeview(panel, columns=columnas_h, show="headings", height=18)
        
        for col in columnas_h:
            tabla_historial.heading(col, text=col)
            if col == "Detalles del Movimiento":
                tabla_historial.column(col, anchor="w", width=420)
            elif col == "Fecha y Hora":
                tabla_historial.column(col, anchor="center", width=160)
            else:
                tabla_historial.column(col, anchor="center", width=130)
                
        tabla_historial.pack(fill="both", expand=True, padx=25, pady=10)

        tk.Button(panel, text="🔄 Actualizar Bitácora", bg=self.colores["botones_menu"], fg=self.colores["oro"],
                font=("Arial", 10, "bold"), relief="flat", padx=10, pady=5,
                command=lambda: cargar_logs()).pack(anchor="w", padx=25, pady=5)

        def cargar_logs():
            for item in tabla_historial.get_children():
                tabla_historial.delete(item)
            try:
                conn = obtener_conexion()
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT id, datetime(fecha_hora, 'localtime'), usuario, accion, detalles 
                    FROM historial_actividades 
                    ORDER BY id DESC
                """)
                for fila in cursor.fetchall():
                    tabla_historial.insert("", "end", values=fila)
                conn.close()
            except Exception as e:
                print(f"Error cargando historial: {e}")

        cargar_logs()
        
        # ============================================================
    # ============================================================
    # MONITOR GLOBAL DE MENSAJES
    # ============================================================

    def iniciar_monitor_mensajes(self):
        """Inicia la vigilancia global de mensajes nuevos."""
        try:
            conn = obtener_conexion(); cursor = conn.cursor()
            cursor.execute("SELECT COALESCE(MAX(id), 0) FROM mensajes")
            resultado = cursor.fetchone()
            self.ultimo_mensaje_notificado = resultado[0] if resultado else 0
            conn.close()
        except Exception as e:
            print(f"Error iniciando monitor de mensajes: {e}")
            self.ultimo_mensaje_notificado = 0
        self.root.after(2000, self.monitor_mensajes)

    def monitor_mensajes(self):
        """Comprueba periódicamente si llegaron mensajes para el usuario actual."""
        try:
            usuario_actual = getattr(self, "usuario_autenticado", None)
            if not usuario_actual:
                self.root.after(2000, self.monitor_mensajes); return
            conn = obtener_conexion(); cursor = conn.cursor()
            cursor.execute("SELECT id FROM usuarios WHERE username = ?", (usuario_actual,))
            usuario = cursor.fetchone()
            if not usuario:
                conn.close(); self.root.after(2000, self.monitor_mensajes); return
            usuario_id = usuario[0]
            cursor.execute("""
                SELECT m.id, m.contenido, u.username, m.conversacion_id
                FROM mensajes m
                INNER JOIN usuarios u ON u.id = m.usuario_id
                INNER JOIN participantes_chat pc ON pc.conversacion_id = m.conversacion_id
                WHERE m.id > ? AND pc.usuario_id = ? AND m.usuario_id != ?
                  AND COALESCE(m.leido, 0) = 0
                ORDER BY m.id ASC
            """, (self.ultimo_mensaje_notificado, usuario_id, usuario_id))
            nuevos = cursor.fetchall(); conn.close()
            for mid, contenido, remitente, conv_id in nuevos:
                self.ultimo_mensaje_notificado = max(self.ultimo_mensaje_notificado, mid)
                self.mostrar_notificacion_chat(remitente, contenido or "", conv_id)
        except Exception as e:
            print(f"Error monitorizando mensajes: {e}")
        try:
            self.root.after(2000, self.monitor_mensajes)
        except Exception:
            pass

    def mostrar_notificacion_chat(self, remitente, contenido, conversacion_id):
        self.notificaciones_chat += 1
        n = tk.Toplevel(self.root); n.title("Nuevo mensaje"); n.resizable(False, False); n.configure(bg="white"); n.attributes("-topmost", True)
        ancho, alto = 350, 155; n.update_idletasks()
        sw, sh = n.winfo_screenwidth(), n.winfo_screenheight()
        desplazamiento = max(self.notificaciones_chat - 1, 0) * 12
        n.geometry(f"{ancho}x{alto}+{sw-ancho-25}+{max(20, sh-alto-70-desplazamiento)}")
        tk.Label(n, text="💬  Nuevo mensaje", bg="white", fg=self.colores["sidebar"], font=("Arial", 13, "bold")).pack(anchor="w", padx=15, pady=(12,3))
        tk.Label(n, text=remitente, bg="white", fg="#222222", font=("Arial", 10, "bold")).pack(anchor="w", padx=15)
        texto = contenido[:55] + "..." if len(contenido) > 55 else contenido
        tk.Label(n, text=texto, bg="white", fg="#666666", font=("Arial", 9), wraplength=315, justify="left").pack(anchor="w", padx=15, pady=(2,7))
        tk.Button(n, text="Abrir chat", bg=self.colores["sidebar"], fg="white", activebackground="#4f0e1c", activeforeground="white", relief="flat", cursor="hand2", command=lambda: self.abrir_chat_desde_notificacion(n, remitente, conversacion_id)).pack(padx=15, pady=(0,10), anchor="e")
        n.after(7000, lambda: self.cerrar_notificacion(n))

    def cerrar_notificacion(self, notificacion):
        try:
            if notificacion.winfo_exists(): return
            notificacion.destroy()
            if self.notificaciones_chat > 0: self.notificaciones_chat -= 1
        except Exception:
            pass

    def abrir_chat_desde_notificacion(self, notificacion, remitente, conversacion_id):
        self.cerrar_notificacion(notificacion)
        try:
            self.abrir_chat()
            self.root.after(300, lambda: self.seleccionar_usuario_chat(remitente))
        except Exception as e:
            messagebox.showerror("Chat", f"No se pudo abrir la conversación.\n\n{e}")

    def seleccionar_usuario_chat(self, username):
        try:
            ventana = getattr(self, "ventana_chat", None)
            if not ventana: return
            ventana.cargar_usuarios()
            for indice, datos in ventana.mapa_usuarios.items():
                usuario_id, nombre = datos
                if nombre == username:
                    ventana.lista_usuarios.selection_clear(0, tk.END)
                    ventana.lista_usuarios.selection_set(indice)
                    ventana.lista_usuarios.see(indice)
                    ventana.usuario_seleccionado()
                    return
        except Exception as e:
            print(f"Error seleccionando chat desde notificación: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DashboardSGC(root)
    root.mainloop()