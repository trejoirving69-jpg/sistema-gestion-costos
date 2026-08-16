from flask import Flask, request, jsonify
from conexion import obtener_conexion, registrar_accion

app = Flask(__name__)

def add_solicitud(data):
    conn = obtener_conexion()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO solicitudes_web (cliente_potencial, servicio_interes, descripcion, correo, telefono, estado) VALUES (?, ?, ?, ?, ?, ?)",
        (
            data.get('cliente') or data.get('cliente_potencial') or 'No registrado',
            data.get('servicio') or data.get('servicio_interes') or 'No especificado',
            data.get('descripcion') or data.get('mensaje') or '',
            data.get('correo') or data.get('email') or '',
            data.get('telefono') or data.get('telefono_contacto') or '',
            'Pendiente'
        )
    )
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    # Registrar en bitácora
    try:
        registrar_accion('SistemaWeb', 'NUEVA_SOLICITUD_WEB', f"ID:{last_id} - Cliente:{data.get('cliente')}")
    except Exception:
        pass
    return last_id

@app.route('/api/solicitudes', methods=['POST'])
def recibir_solicitud():
    # Acepta JSON o datos de formulario
    data = request.get_json(silent=True) or request.form.to_dict()
    if not data:
        return jsonify({'ok': False, 'error': 'Sin datos'}), 400

    if not (data.get('cliente') or data.get('cliente_potencial')):
        return jsonify({'ok': False, 'error': 'Campo cliente es obligatorio'}), 400

    solicitud_id = add_solicitud(data)
    response = jsonify({'ok': True, 'id': solicitud_id})
    # Allow simple CORS for testing from a web page
    response.headers['Access-Control-Allow-Origin'] = '*'
    return response, 201

@app.route('/api/solicitudes', methods=['GET'])
def listar_solicitudes():
    conn = obtener_conexion()
    cur = conn.cursor()
    cur.execute("SELECT id, fecha_solicitud, cliente_potencial, servicio_interes, correo, telefono, estado FROM solicitudes_web ORDER BY id DESC")
    rows = cur.fetchall()
    conn.close()
    items = []
    for r in rows:
        items.append({
            'id': r[0],
            'fecha_solicitud': r[1],
            'cliente_potencial': r[2],
            'servicio_interes': r[3],
            'correo': r[4],
            'telefono': r[5],
            'estado': r[6]
        })
    response = jsonify({'ok': True, 'items': items})
    response.headers['Access-Control-Allow-Origin'] = '*'
    return response

if __name__ == '__main__':
    app.run(port=5001, debug=True)
