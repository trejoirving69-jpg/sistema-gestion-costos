import os
import traceback
import sys


base_prefix = getattr(sys, 'base_prefix', sys.prefix)
tcl_lib = os.path.join(base_prefix, 'tcl', 'tcl8.6')
tk_lib = os.path.join(base_prefix, 'tcl', 'tk8.6')
if os.path.isdir(tcl_lib):
    os.environ['TCL_LIBRARY'] = tcl_lib
if os.path.isdir(tk_lib):
    os.environ['TK_LIBRARY'] = tk_lib

import tkinter as tk


try:
    import customtkinter as ctxt
    from login import VentanaLogin
    from dashboard import DashboardSGC as DashboardPrincipal
    from conexion import inicializar_base_datos
except Exception as e:

    try:

        base_path = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
        err_path = os.path.join(base_path, "startup_import_error.log")
        with open(err_path, "a", encoding="utf-8") as ef:
            ef.write("IMPORT ERROR:\n")
            ef.write(str(e) + "\n")
            traceback.print_exc(file=ef)
    except Exception:
        pass
    raise

from conexion import obtener_app_path

LOG_PATH = os.path.join(obtener_app_path(), "startup.log")


def escribir_log(mensaje):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(mensaje + "\n")
    except Exception:
        
        try:
            fallback = os.path.join(os.path.dirname(os.path.abspath(__file__)), "startup.log")
            with open(fallback, "a", encoding="utf-8") as f:
                f.write(mensaje + "\n")
        except Exception:
            pass


def iniciar_sistema():
    escribir_log("iniciar_sistema: inicio")
    ventana_login = VentanaLogin()
    ventana_login.mainloop()
    
    if ventana_login.usuario_autenticado and ventana_login.rol_autenticado:
        escribir_log(f"iniciar_sistema: usuario autenticado {ventana_login.usuario_autenticado}")
        
        root = tk.Tk()
        # Modificación aquí: Primero instanciamos el Dashboard
        app = DashboardPrincipal(root)
        # Asignamos las propiedades de sesión
        app.usuario_autenticado = ventana_login.usuario_autenticado
        app.rol_autenticado = ventana_login.rol_autenticado
        app.password_autenticado = ventana_login.password_autenticado
        # Aplicar permisos UI según rol (por ejemplo ocultar historial si no es admin)
        try:
            app.aplicar_permisos()
        except Exception:
            pass

        try:
            app.actualizar_alertas_globales()
        except Exception:
            pass

        
        root.mainloop()
        
if __name__ == "__main__":
    escribir_log("main: inicializando base de datos")
    inicializar_base_datos()
    escribir_log("main: base de datos inicializada")
    iniciar_sistema()