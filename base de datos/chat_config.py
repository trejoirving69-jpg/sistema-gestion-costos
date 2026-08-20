import os

# ============================================================
# CONFIGURACIÓN DEL CHAT EN RED
# ============================================================

# Servidor de chat.
#
# En TU PC:
# 127.0.0.1 funciona solamente en esta computadora.
#
# Para que otra PC de tu misma red pueda conectarse:
# http://192.168.1.101:5000
#
CHAT_SERVER_URL = "https://irving8579.pythonanywhere.com"


# Clave interna del servidor.
#
# Por ahora NO necesitamos utilizarla para el login.
# El chat utilizará el token Bearer generado al iniciar sesión.
CHAT_API_KEY = os.getenv("CHAT_API_KEY", "").strip()
# Activar modo red
NETWORK_MODE = bool(CHAT_SERVER_URL.strip())