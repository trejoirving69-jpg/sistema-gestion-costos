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
CHAT_SERVER_URL = "http://127.0.0.1:5000"


# Clave interna del servidor.
#
# Por ahora NO necesitamos utilizarla para el login.
# El chat utilizará el token Bearer generado al iniciar sesión.
CHAT_API_KEY = "CAMBIAR_ESTA_CLAVE"


# Activar modo red
NETWORK_MODE = bool(CHAT_SERVER_URL.strip())