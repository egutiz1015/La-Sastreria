"""
La Sastrería - Sistema de Gestión
Fundador: Erick Gutierrez | Administradora: Keila Gutierrez
"""

from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file, flash
import sqlite3
import hashlib
import os
import json
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import functools

app = Flask(__name__)
app.secret_key = 'sastreria_erick_keila_2024_secret_key_secure'

DB_PATH = os.path.join(os.path.dirname(__file__), 'data', 'sastreria.db')
EXPORTS_PATH = os.path.join(os.path.dirname(__file__), 'exports')

# ─────────────────────────────────────────────
# USUARIOS PREDEFINIDOS (dueño y secretaria)
# ─────────────────────────────────────────────
USERS = {
    'erick': {
        'password': 'c98583b2cb9f8dcd129c0cda6913a8ed1db835a877842ee208569b50b97f46ee',
        'name': 'Erick Gutierrez',
        'role': 'owner',
        'display': 'Dueño & Fundador'
    },
    'keila': {
        'password': 'd6730f9706f88db370fef218422d05d16e3f94b9bbcf5dfa239cf21ff4ce8a11',
        'name': 'Keila Gutierrez',
        'role': 'admin',
        'display': 'Administradora'
    }
}

# ─────────────────────────────────────────────
# BASE DE DATOS
# ─────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_db()
    c = conn.cursor()

    # Tabla de clientes
    c.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            telefono TEXT,
            correo TEXT,
            nit TEXT,
            direccion TEXT,
            notas TEXT,
            fecha_registro TEXT NOT NULL,
            activo INTEGER DEFAULT 1
        )
    ''')

    # Tabla de órdenes de trabajo
    c.execute('''
        CREATE TABLE IF NOT EXISTS ordenes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_orden TEXT UNIQUE NOT NULL,
            cliente_id INTEGER NOT NULL,
            fecha_orden TEXT NOT NULL,
            fecha_entrega TEXT NOT NULL,
            estado TEXT DEFAULT 'pendiente',
            total REAL DEFAULT 0,
            notas_adicionales TEXT,
            usuario_registro TEXT,
            fecha_creacion TEXT NOT NULL,
            correo_enviado INTEGER DEFAULT 0,
            FOREIGN KEY (cliente_id) REFERENCES clientes(id)
        )
    ''')

    # Tabla de prendas/servicios por orden
    c.execute('''
        CREATE TABLE IF NOT EXISTS prendas_orden (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            orden_id INTEGER NOT NULL,
            tipo_prenda TEXT NOT NULL,
            descripcion_servicio TEXT NOT NULL,
            cantidad INTEGER DEFAULT 1,
            precio_tipo TEXT DEFAULT 'fijo',
            precio_unitario REAL NOT NULL,
            subtotal REAL NOT NULL,
            entregada INTEGER DEFAULT 0,
            fecha_entrega_real TEXT,
            FOREIGN KEY (orden_id) REFERENCES ordenes(id)
        )
    ''')

    # Catálogo de prendas/servicios
    c.execute('''
        CREATE TABLE IF NOT EXISTS catalogo_servicios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            descripcion TEXT,
            precio_base REAL,
            precio_tipo TEXT DEFAULT 'fijo',
            activo INTEGER DEFAULT 1
        )
    ''')

    # Configuración del sistema
    c.execute('''
        CREATE TABLE IF NOT EXISTS configuracion (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    ''')

    # Insertar configuración por defecto
    configs = [
        ('empresa_nombre', 'La Sastrería'),
        ('empresa_telefono', ''),
        ('empresa_correo', ''),
        ('smtp_servidor', 'smtp.gmail.com'),
        ('smtp_puerto', '587'),
        ('smtp_usuario', ''),
        ('smtp_password', ''),
        ('moneda', 'Q'),
    ]
    for clave, valor in configs:
        c.execute('INSERT OR IGNORE INTO configuracion VALUES (?, ?)', (clave, valor))

    # Servicios de catálogo por defecto
    servicios_default = [
        ('Pantalón - Arreglo de ruedo', 'Subir o bajar ruedo de pantalón', 25.00, 'fijo'),
        ('Pantalón - Entrada', 'Ajuste de cintura/piernas', 35.00, 'fijo'),
        ('Camisa - Arreglo manga', 'Acortar mangas de camisa', 30.00, 'fijo'),
        ('Vestido - Arreglo general', 'Arreglo de vestido', 0.00, 'variable'),
        ('Traje - Confección completa', 'Confección de traje a medida', 0.00, 'variable'),
        ('Falda - Arreglo de ruedo', 'Subir o bajar ruedo de falda', 25.00, 'fijo'),
        ('Chaqueta - Ajuste', 'Ajuste de chaqueta', 0.00, 'variable'),
        ('Bordado', 'Servicio de bordado en prenda', 0.00, 'variable'),
        ('Zipper - Cambio', 'Cambio de zipper/cremallera', 30.00, 'fijo'),
        ('Botones - Costura', 'Costura de botones', 10.00, 'fijo'),
    ]
    for nombre, desc, precio, tipo in servicios_default:
        c.execute('INSERT OR IGNORE INTO catalogo_servicios (nombre, descripcion, precio_base, precio_tipo) VALUES (?,?,?,?)',
                  (nombre, desc, precio, tipo))

    conn.commit()
    conn.close()

# ─────────────────────────────────────────────
# DECORADORES Y HELPERS
# ─────────────────────────────────────────────
def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if 'usuario' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def generate_order_number():
    conn = get_db()
    c = conn.cursor()
    today = date.today()
    prefix = f"ORD-{today.strftime('%Y%m')}-"
    c.execute("SELECT numero_orden FROM ordenes WHERE numero_orden LIKE ? ORDER BY id DESC LIMIT 1", (prefix + '%',))
    last = c.fetchone()
    conn.close()
    if last:
        last_num = int(last['numero_orden'].split('-')[-1])
        return f"{prefix}{last_num + 1:04d}"
    return f"{prefix}0001"

def get_config():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT clave, valor FROM configuracion")
    config = {row['clave']: row['valor'] for row in c.fetchall()}
    conn.close()
    return config

# ─────────────────────────────────────────────
# RUTAS DE AUTENTICACIÓN
# ─────────────────────────────────────────────
@app.route('/')
def index():
    if 'usuario' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        usuario = request.form.get('usuario', '').strip().lower()
        password = request.form.get('password', '')
        password_hash = hashlib.sha256(password.encode()).hexdigest()

        if usuario in USERS and USERS[usuario]['password'] == password_hash:
            session['usuario'] = usuario
            session['nombre'] = USERS[usuario]['name']
            session['rol'] = USERS[usuario]['role']
            session['display'] = USERS[usuario]['display']
            return redirect(url_for('dashboard'))
        else:
            error = 'Usuario o contraseña incorrectos'

    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ─────────────────────────────────────────────
# DASHBOARD
# ─────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    conn = get_db()
    c = conn.cursor()

    # Estadísticas generales
    c.execute("SELECT COUNT(*) as total FROM clientes WHERE activo=1")
    total_clientes = c.fetchone()['total']

    c.execute("SELECT COUNT(*) as total FROM ordenes WHERE estado='pendiente'")
    ordenes_pendientes = c.fetchone()['total']

    c.execute("SELECT COUNT(*) as total FROM ordenes WHERE estado='entregado'")
    ordenes_entregadas = c.fetchone()['total']

    c.execute("SELECT COALESCE(SUM(total),0) as total FROM ordenes WHERE estado='entregado'")
    ingresos_total = c.fetchone()['total']

    # Próximas entregas (7 días)
    hoy = date.today().isoformat()
    en_7_dias = (date.today() + timedelta(days=7)).isoformat()
    c.execute("""
        SELECT o.*, cl.nombre as cliente_nombre
        FROM ordenes o
        JOIN clientes cl ON o.cliente_id = cl.id
        WHERE o.fecha_entrega BETWEEN ? AND ? AND o.estado='pendiente'
        ORDER BY o.fecha_entrega ASC
        LIMIT 10
    """, (hoy, en_7_dias))
    proximas_entregas = c.fetchall()

    # Órdenes recientes
    c.execute("""
        SELECT o.*, cl.nombre as cliente_nombre
        FROM ordenes o
        JOIN clientes cl ON o.cliente_id = cl.id
        ORDER BY o.fecha_creacion DESC
        LIMIT 8
    """)
    ordenes_recientes = c.fetchall()

    # Entregas hoy
    c.execute("""
        SELECT COUNT(*) as total FROM ordenes
        WHERE fecha_entrega = ? AND estado='pendiente'
    """, (hoy,))
    entregas_hoy = c.fetchone()['total']

    conn.close()
    return render_template('dashboard.html',
        total_clientes=total_clientes,
        ordenes_pendientes=ordenes_pendientes,
        ordenes_entregadas=ordenes_entregadas,
        ingresos_total=ingresos_total,
        proximas_entregas=proximas_entregas,
        ordenes_recientes=ordenes_recientes,
        entregas_hoy=entregas_hoy,
        hoy=hoy
    )

# ─────────────────────────────────────────────
# CLIENTES
# ─────────────────────────────────────────────
@app.route('/clientes')
@login_required
def clientes():
    busqueda = request.args.get('q', '')
    conn = get_db()
    c = conn.cursor()
    if busqueda:
        c.execute("""
            SELECT cl.*, COUNT(o.id) as total_ordenes,
            MAX(o.fecha_orden) as ultima_orden
            FROM clientes cl
            LEFT JOIN ordenes o ON cl.id = o.cliente_id
            WHERE cl.activo=1 AND (cl.nombre LIKE ? OR cl.telefono LIKE ? OR cl.nit LIKE ? OR cl.correo LIKE ?)
            GROUP BY cl.id ORDER BY cl.nombre
        """, (f'%{busqueda}%',)*4)
    else:
        c.execute("""
            SELECT cl.*, COUNT(o.id) as total_ordenes,
            MAX(o.fecha_orden) as ultima_orden
            FROM clientes cl
            LEFT JOIN ordenes o ON cl.id = o.cliente_id
            WHERE cl.activo=1
            GROUP BY cl.id ORDER BY cl.nombre
        """)
    lista = c.fetchall()
    conn.close()
    return render_template('clientes.html', clientes=lista, busqueda=busqueda)

@app.route('/clientes/nuevo', methods=['GET', 'POST'])
@login_required
def nuevo_cliente():
    if request.method == 'POST':
        conn = get_db()
        c = conn.cursor()
        c.execute("""
            INSERT INTO clientes (nombre, telefono, correo, nit, direccion, notas, fecha_registro)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            request.form['nombre'],
            request.form.get('telefono', ''),
            request.form.get('correo', ''),
            request.form.get('nit', ''),
            request.form.get('direccion', ''),
            request.form.get('notas', ''),
            date.today().isoformat()
        ))
        cliente_id = c.lastrowid
        conn.commit()
        conn.close()
        flash('Cliente registrado exitosamente', 'success')
        return redirect(url_for('detalle_cliente', id=cliente_id))
    return render_template('nuevo_cliente.html')

@app.route('/clientes/<int:id>')
@login_required
def detalle_cliente(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM clientes WHERE id=?", (id,))
    cliente = c.fetchone()
    if not cliente:
        return redirect(url_for('clientes'))

    c.execute("""
        SELECT o.*, COUNT(po.id) as total_prendas
        FROM ordenes o
        LEFT JOIN prendas_orden po ON o.id = po.orden_id
        WHERE o.cliente_id=?
        GROUP BY o.id
        ORDER BY o.fecha_orden DESC
    """, (id,))
    ordenes = c.fetchall()
    conn.close()
    return render_template('detalle_cliente.html', cliente=cliente, ordenes=ordenes)

@app.route('/clientes/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def editar_cliente(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM clientes WHERE id=?", (id,))
    cliente = c.fetchone()

    if request.method == 'POST':
        c.execute("""
            UPDATE clientes SET nombre=?, telefono=?, correo=?, nit=?, direccion=?, notas=?
            WHERE id=?
        """, (
            request.form['nombre'],
            request.form.get('telefono', ''),
            request.form.get('correo', ''),
            request.form.get('nit', ''),
            request.form.get('direccion', ''),
            request.form.get('notas', ''),
            id
        ))
        conn.commit()
        conn.close()
        flash('Cliente actualizado', 'success')
        return redirect(url_for('detalle_cliente', id=id))

    conn.close()
    return render_template('editar_cliente.html', cliente=cliente)

# API: buscar clientes para autocompletar
@app.route('/api/clientes/buscar')
@login_required
def api_buscar_clientes():
    q = request.args.get('q', '')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id, nombre, telefono, correo, nit FROM clientes WHERE activo=1 AND nombre LIKE ? LIMIT 10", (f'%{q}%',))
    resultados = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify(resultados)

# ─────────────────────────────────────────────
# ÓRDENES
# ─────────────────────────────────────────────
@app.route('/ordenes')
@login_required
def ordenes():
    estado = request.args.get('estado', 'todos')
    busqueda = request.args.get('q', '')
    conn = get_db()
    c = conn.cursor()

    query = """
        SELECT o.*, cl.nombre as cliente_nombre, cl.telefono as cliente_tel,
        COUNT(po.id) as total_prendas
        FROM ordenes o
        JOIN clientes cl ON o.cliente_id = cl.id
        LEFT JOIN prendas_orden po ON o.id = po.orden_id
        WHERE 1=1
    """
    params = []

    if estado != 'todos':
        query += " AND o.estado=?"
        params.append(estado)

    if busqueda:
        query += " AND (cl.nombre LIKE ? OR o.numero_orden LIKE ?)"
        params.extend([f'%{busqueda}%', f'%{busqueda}%'])

    query += " GROUP BY o.id ORDER BY o.fecha_orden DESC"

    c.execute(query, params)
    lista = c.fetchall()
    conn.close()
    return render_template('ordenes.html', ordenes=lista, estado_filtro=estado, busqueda=busqueda)

@app.route('/ordenes/nueva', methods=['GET', 'POST'])
@login_required
def nueva_orden():
    if request.method == 'POST':
        data = request.get_json()
        conn = get_db()
        c = conn.cursor()

        numero_orden = generate_order_number()

        # Crear o usar cliente existente
        cliente_id = data.get('cliente_id')
        if not cliente_id:
            # Cliente nuevo
            c.execute("""
                INSERT INTO clientes (nombre, telefono, correo, nit, fecha_registro)
                VALUES (?, ?, ?, ?, ?)
            """, (data['cliente_nombre'], data.get('cliente_tel', ''),
                  data.get('cliente_correo', ''), data.get('cliente_nit', ''),
                  date.today().isoformat()))
            cliente_id = c.lastrowid

        # Crear orden
        c.execute("""
            INSERT INTO ordenes (numero_orden, cliente_id, fecha_orden, fecha_entrega,
            estado, total, notas_adicionales, usuario_registro, fecha_creacion)
            VALUES (?, ?, ?, ?, 'pendiente', ?, ?, ?, ?)
        """, (numero_orden, cliente_id, data['fecha_orden'], data['fecha_entrega'],
              data['total'], data.get('notas', ''), session['usuario'],
              datetime.now().isoformat()))

        orden_id = c.lastrowid

        # Insertar prendas
        for prenda in data['prendas']:
            c.execute("""
                INSERT INTO prendas_orden (orden_id, tipo_prenda, descripcion_servicio,
                cantidad, precio_tipo, precio_unitario, subtotal)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (orden_id, prenda['tipo'], prenda['descripcion'],
                  prenda['cantidad'], prenda['precio_tipo'],
                  prenda['precio_unitario'], prenda['subtotal']))

        conn.commit()

        # Obtener datos para el correo y Excel
        c.execute("SELECT * FROM ordenes WHERE id=?", (orden_id,))
        orden = dict(c.fetchone())
        c.execute("SELECT * FROM clientes WHERE id=?", (cliente_id,))
        cliente = dict(c.fetchone())
        c.execute("SELECT * FROM prendas_orden WHERE orden_id=?", (orden_id,))
        prendas = [dict(p) for p in c.fetchall()]

        conn.close()

        # Exportar a Excel
        exportar_orden_excel(orden, cliente, prendas)

        # Enviar correo si tiene email
        if cliente.get('correo'):
            enviar_correo_orden(orden, cliente, prendas)
            conn2 = get_db()
            conn2.execute("UPDATE ordenes SET correo_enviado=1 WHERE id=?", (orden_id,))
            conn2.commit()
            conn2.close()

        return jsonify({'success': True, 'numero_orden': numero_orden, 'orden_id': orden_id})

    # GET - Formulario
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM catalogo_servicios WHERE activo=1 ORDER BY nombre")
    catalogo = [dict(s) for s in c.fetchall()]
    conn.close()

    numero_previo = generate_order_number()
    return render_template('nueva_orden.html', catalogo=catalogo,
                           numero_previo=numero_previo,
                           fecha_hoy=date.today().isoformat())

@app.route('/ordenes/<int:id>')
@login_required
def detalle_orden(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT o.*, cl.nombre as cliente_nombre, cl.telefono as cliente_tel,
        cl.correo as cliente_correo, cl.nit as cliente_nit
        FROM ordenes o JOIN clientes cl ON o.cliente_id=cl.id
        WHERE o.id=?
    """, (id,))
    orden = c.fetchone()
    if not orden:
        return redirect(url_for('ordenes'))

    c.execute("SELECT * FROM prendas_orden WHERE orden_id=? ORDER BY id", (id,))
    prendas = c.fetchall()
    conn.close()
    return render_template('detalle_orden.html', orden=orden, prendas=prendas)

@app.route('/ordenes/<int:id>/editar', methods=['GET', 'POST'])
@login_required
def editar_orden(id):
    conn = get_db()
    c = conn.cursor()

    if request.method == 'POST':
        data = request.get_json()
        c.execute("""
            UPDATE ordenes SET fecha_entrega=?, notas_adicionales=?, total=?
            WHERE id=?
        """, (data['fecha_entrega'], data.get('notas', ''), data['total'], id))

        # Actualizar prendas existentes
        for prenda in data.get('prendas_actualizar', []):
            c.execute("""
                UPDATE prendas_orden SET tipo_prenda=?, descripcion_servicio=?,
                cantidad=?, precio_unitario=?, subtotal=?
                WHERE id=?
            """, (prenda['tipo'], prenda['descripcion'], prenda['cantidad'],
                  prenda['precio_unitario'], prenda['subtotal'], prenda['id']))

        # Agregar prendas nuevas
        for prenda in data.get('prendas_nuevas', []):
            c.execute("""
                INSERT INTO prendas_orden (orden_id, tipo_prenda, descripcion_servicio,
                cantidad, precio_tipo, precio_unitario, subtotal)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (id, prenda['tipo'], prenda['descripcion'], prenda['cantidad'],
                  prenda['precio_tipo'], prenda['precio_unitario'], prenda['subtotal']))

        # Eliminar prendas
        for prenda_id in data.get('prendas_eliminar', []):
            c.execute("DELETE FROM prendas_orden WHERE id=?", (prenda_id,))

        conn.commit()
        conn.close()
        return jsonify({'success': True})

    c.execute("""
        SELECT o.*, cl.nombre as cliente_nombre FROM ordenes o
        JOIN clientes cl ON o.cliente_id=cl.id WHERE o.id=?
    """, (id,))
    orden = c.fetchone()
    c.execute("SELECT * FROM prendas_orden WHERE orden_id=?", (id,))
    prendas = c.fetchall()
    c.execute("SELECT * FROM catalogo_servicios WHERE activo=1")
    catalogo = [dict(s) for s in c.fetchall()]
    conn.close()
    return render_template('editar_orden.html', orden=orden, prendas=prendas, catalogo=catalogo)

@app.route('/api/ordenes/<int:id>/estado', methods=['POST'])
@login_required
def cambiar_estado_orden(id):
    data = request.get_json()
    nuevo_estado = data.get('estado')
    conn = get_db()
    conn.execute("UPDATE ordenes SET estado=? WHERE id=?", (nuevo_estado, id))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/api/prendas/<int:id>/entregar', methods=['POST'])
@login_required
def entregar_prenda(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("UPDATE prendas_orden SET entregada=1, fecha_entrega_real=? WHERE id=?",
              (date.today().isoformat(), id))

    # Verificar si todas las prendas de la orden están entregadas
    c.execute("SELECT orden_id FROM prendas_orden WHERE id=?", (id,))
    orden_id = c.fetchone()['orden_id']
    c.execute("SELECT COUNT(*) as total, SUM(entregada) as entregadas FROM prendas_orden WHERE orden_id=?", (orden_id,))
    row = c.fetchone()
    if row['total'] == row['entregadas']:
        c.execute("UPDATE ordenes SET estado='entregado' WHERE id=?", (orden_id,))

    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ─────────────────────────────────────────────
# CALENDARIO
# ─────────────────────────────────────────────
@app.route('/calendario')
@login_required
def calendario():
    return render_template('calendario.html')

@app.route('/api/calendario/eventos')
@login_required
def api_eventos_calendario():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT o.id, o.numero_orden, o.fecha_entrega, o.estado, o.total,
        cl.nombre as cliente_nombre
        FROM ordenes o JOIN clientes cl ON o.cliente_id=cl.id
        ORDER BY o.fecha_entrega
    """)
    ordenes = c.fetchall()
    conn.close()

    eventos = []
    for o in ordenes:
        color = '#f59e0b' if o['estado'] == 'pendiente' else '#10b981' if o['estado'] == 'entregado' else '#6b7280'
        eventos.append({
            'id': o['id'],
            'title': f"{o['numero_orden']} - {o['cliente_nombre']}",
            'start': o['fecha_entrega'],
            'color': color,
            'estado': o['estado'],
            'total': o['total']
        })
    return jsonify(eventos)

# ─────────────────────────────────────────────
# CATÁLOGO DE SERVICIOS
# ─────────────────────────────────────────────
@app.route('/catalogo')
@login_required
def catalogo():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM catalogo_servicios ORDER BY nombre")
    servicios = [dict(s) for s in c.fetchall()]
    conn.close()
    return render_template('catalogo.html', servicios=servicios)

@app.route('/catalogo/nuevo', methods=['POST'])
@login_required
def nuevo_servicio():
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        INSERT INTO catalogo_servicios (nombre, descripcion, precio_base, precio_tipo)
        VALUES (?, ?, ?, ?)
    """, (data['nombre'], data.get('descripcion', ''),
          float(data.get('precio_base', 0)), data.get('precio_tipo', 'fijo')))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

@app.route('/catalogo/<int:id>/editar', methods=['POST'])
@login_required
def editar_servicio(id):
    data = request.get_json()
    conn = get_db()
    conn.execute("""
        UPDATE catalogo_servicios SET nombre=?, descripcion=?, precio_base=?, precio_tipo=?, activo=?
        WHERE id=?
    """, (data['nombre'], data.get('descripcion', ''), float(data.get('precio_base', 0)),
          data.get('precio_tipo', 'fijo'), data.get('activo', 1), id))
    conn.commit()
    conn.close()
    return jsonify({'success': True})

# ─────────────────────────────────────────────
# EXPORTAR / DESCARGAR
# ─────────────────────────────────────────────
@app.route('/ordenes/<int:id>/descargar')
@login_required
def descargar_orden(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT o.*, cl.nombre as cliente_nombre, cl.telefono as cliente_tel, cl.correo as cliente_correo, cl.nit as cliente_nit FROM ordenes o JOIN clientes cl ON o.cliente_id=cl.id WHERE o.id=?", (id,))
    orden = dict(c.fetchone())
    c.execute("SELECT * FROM clientes WHERE id=?", (orden['cliente_id'],))
    cliente = dict(c.fetchone())
    c.execute("SELECT * FROM prendas_orden WHERE orden_id=?", (id,))
    prendas = [dict(p) for p in c.fetchall()]
    conn.close()

    filepath = exportar_orden_excel(orden, cliente, prendas)
    return send_file(filepath, as_attachment=True,
                     download_name=f"Orden_{orden['numero_orden']}.xlsx")

@app.route('/exportar/clientes')
@login_required
def exportar_clientes():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT cl.*, COUNT(o.id) as total_ordenes, COALESCE(SUM(o.total),0) as total_gastado
        FROM clientes cl LEFT JOIN ordenes o ON cl.id=o.cliente_id
        WHERE cl.activo=1 GROUP BY cl.id ORDER BY cl.nombre
    """)
    clientes = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Clientes"

    headers = ['ID', 'Nombre', 'Teléfono', 'Correo', 'NIT', 'Dirección', 'Notas', 'Fecha Registro', 'Total Órdenes', 'Total Gastado (Q)']
    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')

    for row_num, cliente in enumerate(clientes, 2):
        ws.cell(row=row_num, column=1, value=cliente['id'])
        ws.cell(row=row_num, column=2, value=cliente['nombre'])
        ws.cell(row=row_num, column=3, value=cliente['telefono'])
        ws.cell(row=row_num, column=4, value=cliente['correo'])
        ws.cell(row=row_num, column=5, value=cliente['nit'])
        ws.cell(row=row_num, column=6, value=cliente['direccion'])
        ws.cell(row=row_num, column=7, value=cliente['notas'])
        ws.cell(row=row_num, column=8, value=cliente['fecha_registro'])
        ws.cell(row=row_num, column=9, value=cliente['total_ordenes'])
        ws.cell(row=row_num, column=10, value=round(cliente['total_gastado'], 2))

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    filepath = os.path.join(EXPORTS_PATH, f"Clientes_{date.today().isoformat()}.xlsx")
    wb.save(filepath)
    return send_file(filepath, as_attachment=True, download_name=f"Clientes_{date.today().isoformat()}.xlsx")

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
@app.route('/configuracion', methods=['GET', 'POST'])
@login_required
def configuracion():
    if session.get('rol') != 'owner':
        flash('Solo el dueño puede acceder a la configuración', 'error')
        return redirect(url_for('dashboard'))

    conn = get_db()
    config = get_config()

    if request.method == 'POST':
        for clave in ['empresa_telefono', 'empresa_correo', 'smtp_servidor',
                      'smtp_puerto', 'smtp_usuario', 'smtp_password']:
            conn.execute("UPDATE configuracion SET valor=? WHERE clave=?",
                         (request.form.get(clave, ''), clave))
        conn.commit()
        flash('Configuración guardada', 'success')
        config = get_config()

    conn.close()
    return render_template('configuracion.html', config=config)

# ─────────────────────────────────────────────
# FUNCIONES AUXILIARES: EXCEL + CORREO
# ─────────────────────────────────────────────
def exportar_orden_excel(orden, cliente, prendas):
    os.makedirs(EXPORTS_PATH, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orden de Trabajo"

    # Estilos
    titulo_font = Font(name='Arial', size=16, bold=True, color='1a1a2e')
    subtitulo_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    accent_fill = PatternFill(start_color="c8a96e", end_color="c8a96e", fill_type="solid")
    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    # Cabecera empresa
    ws.merge_cells('A1:G1')
    ws['A1'] = 'LA SASTRERÍA'
    ws['A1'].font = Font(name='Arial', size=18, bold=True, color='1a1a2e')
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:G2')
    ws['A2'] = 'Erick Gutierrez - Fundador'
    ws['A2'].alignment = Alignment(horizontal='center')
    ws['A2'].font = Font(name='Arial', size=10, color='666666')

    ws.merge_cells('A3:G3')
    ws['A3'] = f"ORDEN DE TRABAJO: {orden['numero_orden']}"
    ws['A3'].font = Font(name='Arial', size=13, bold=True, color='c8a96e')
    ws['A3'].alignment = Alignment(horizontal='center')

    # Datos cliente
    ws['A5'] = 'DATOS DEL CLIENTE'
    ws['A5'].font = subtitulo_font
    ws.merge_cells('A5:C5')
    ws['A5'].fill = header_fill

    ws['A6'] = 'Nombre:'
    ws['B6'] = cliente['nombre']
    ws['A7'] = 'Teléfono:'
    ws['B7'] = cliente.get('telefono', '')
    ws['A8'] = 'Correo:'
    ws['B8'] = cliente.get('correo', '')
    ws['A9'] = 'NIT:'
    ws['B9'] = cliente.get('nit', '')

    ws['E5'] = 'DATOS DE LA ORDEN'
    ws['E5'].font = subtitulo_font
    ws.merge_cells('E5:G5')
    ws['E5'].fill = header_fill

    ws['E6'] = 'Fecha Orden:'
    ws['F6'] = orden['fecha_orden']
    ws['E7'] = 'Fecha Entrega:'
    ws['F7'] = orden['fecha_entrega']
    ws['E8'] = 'Estado:'
    ws['F8'] = orden['estado'].upper()
    ws['E9'] = 'Registrado por:'
    ws['F9'] = orden.get('usuario_registro', '')

    # Tabla de prendas
    ws.row_dimensions[11].height = 20
    headers_prendas = ['Tipo de Prenda', 'Descripción del Servicio', 'Cantidad', 'Precio Tipo', 'Precio Unit. (Q)', 'Subtotal (Q)', 'Entregada']
    for col, h in enumerate(headers_prendas, 1):
        cell = ws.cell(row=11, column=col, value=h)
        cell.fill = header_fill
        cell.font = subtitulo_font
        cell.alignment = Alignment(horizontal='center')
        cell.border = border

    for i, prenda in enumerate(prendas, 12):
        ws.cell(row=i, column=1, value=prenda['tipo_prenda']).border = border
        ws.cell(row=i, column=2, value=prenda['descripcion_servicio']).border = border
        ws.cell(row=i, column=3, value=prenda['cantidad']).border = border
        ws.cell(row=i, column=4, value=prenda.get('precio_tipo', 'fijo').upper()).border = border
        ws.cell(row=i, column=5, value=round(prenda['precio_unitario'], 2)).border = border
        ws.cell(row=i, column=6, value=round(prenda['subtotal'], 2)).border = border
        ws.cell(row=i, column=7, value='Sí' if prenda.get('entregada') else 'No').border = border

    last_row = 12 + len(prendas)
    ws.merge_cells(f'A{last_row}:E{last_row}')
    ws[f'A{last_row}'] = 'TOTAL'
    ws[f'A{last_row}'].font = Font(bold=True, size=12)
    ws[f'A{last_row}'].fill = accent_fill
    ws[f'A{last_row}'].alignment = Alignment(horizontal='right')

    ws[f'F{last_row}'] = round(orden['total'], 2)
    ws[f'F{last_row}'].font = Font(bold=True, size=12)
    ws[f'F{last_row}'].fill = accent_fill

    if orden.get('notas_adicionales'):
        ws[f'A{last_row+2}'] = 'Notas adicionales:'
        ws[f'A{last_row+2}'].font = Font(bold=True)
        ws[f'B{last_row+2}'] = orden['notas_adicionales']

    # Ajustar columnas
    col_widths = [25, 35, 12, 14, 18, 14, 12]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    filepath = os.path.join(EXPORTS_PATH, f"Orden_{orden['numero_orden']}.xlsx")
    wb.save(filepath)
    return filepath

def enviar_correo_orden(orden, cliente, prendas):
    try:
        config = get_config()
        smtp_user = config.get('smtp_usuario', '')
        smtp_pass = config.get('smtp_password', '')
        smtp_server = config.get('smtp_servidor', 'smtp.gmail.com')
        smtp_port = int(config.get('smtp_puerto', 587))

        if not smtp_user or not smtp_pass:
            return False

        msg = MIMEMultipart()
        msg['From'] = smtp_user
        msg['To'] = cliente['correo']
        msg['Subject'] = f"La Sastrería - Orden {orden['numero_orden']}"

        prendas_html = ''.join([
            f"<tr><td>{p['tipo_prenda']}</td><td>{p['descripcion_servicio']}</td>"
            f"<td>{p['cantidad']}</td><td>Q{p['precio_unitario']:.2f}</td>"
            f"<td>Q{p['subtotal']:.2f}</td></tr>"
            for p in prendas
        ])

        body = f"""
        <html><body style="font-family: Arial; color: #333;">
        <div style="max-width:600px;margin:auto;border:1px solid #ddd;border-radius:8px;overflow:hidden;">
          <div style="background:#1a1a2e;padding:20px;text-align:center;">
            <h1 style="color:#c8a96e;margin:0;">La Sastrería</h1>
            <p style="color:#fff;margin:5px 0;">Erick Gutierrez - Fundador</p>
          </div>
          <div style="padding:20px;">
            <h2>Orden de Trabajo: {orden['numero_orden']}</h2>
            <p>Estimado/a <strong>{cliente['nombre']}</strong>,</p>
            <p>Su orden ha sido registrada. A continuación el resumen:</p>
            <table style="width:100%;border-collapse:collapse;">
              <tr style="background:#1a1a2e;color:#fff;">
                <th style="padding:8px;text-align:left;">Prenda</th>
                <th style="padding:8px;">Servicio</th>
                <th style="padding:8px;">Cant.</th>
                <th style="padding:8px;">Precio</th>
                <th style="padding:8px;">Subtotal</th>
              </tr>
              {prendas_html}
            </table>
            <div style="background:#f5f5f5;padding:15px;margin-top:15px;border-radius:4px;">
              <strong>Fecha de entrega estimada: {orden['fecha_entrega']}</strong><br>
              <strong style="font-size:18px;color:#1a1a2e;">TOTAL: Q{orden['total']:.2f}</strong>
            </div>
            {"<p><em>Notas: " + orden['notas_adicionales'] + "</em></p>" if orden.get('notas_adicionales') else ""}
          </div>
        </div>
        </body></html>
        """
        msg.attach(MIMEText(body, 'html'))

        # Adjuntar Excel
        filepath = os.path.join(EXPORTS_PATH, f"Orden_{orden['numero_orden']}.xlsx")
        if os.path.exists(filepath):
            with open(filepath, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename=Orden_{orden['numero_orden']}.xlsx")
                msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, cliente['correo'], msg.as_string())
        server.quit()
        return True
    except Exception as e:
        print(f"Error enviando correo: {e}")
        return False

# ─────────────────────────────────────────────
# REENVIAR CORREO
# ─────────────────────────────────────────────
@app.route('/ordenes/<int:id>/reenviar-correo', methods=['POST'])
@login_required
def reenviar_correo(id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT o.*, cl.nombre as cliente_nombre FROM ordenes o JOIN clientes cl ON o.cliente_id=cl.id WHERE o.id=?", (id,))
    orden = dict(c.fetchone())
    c.execute("SELECT * FROM clientes WHERE id=?", (orden['cliente_id'],))
    cliente = dict(c.fetchone())
    c.execute("SELECT * FROM prendas_orden WHERE orden_id=?", (id,))
    prendas = [dict(p) for p in c.fetchall()]
    conn.close()

    if not cliente.get('correo'):
        return jsonify({'success': False, 'error': 'El cliente no tiene correo registrado'})

    resultado = enviar_correo_orden(orden, cliente, prendas)
    return jsonify({'success': resultado})

# Inicializar DB siempre (local y Railway/gunicorn)
init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    is_production = os.environ.get("RAILWAY_ENVIRONMENT") is not None
    print("=" * 50)
    print("  LA SASTRERÍA - Sistema de Gestión")
    print("=" * 50)
    print("  Usuarios:")
    print("  • erick  / TEMP1  (Dueño)")
    print("  • keila  / TEMP2  (Admin)")
    print("=" * 50)
    print(f"  Acceder en: http://localhost:{port}")
    print("=" * 50)
    app.run(debug=not is_production, host="0.0.0.0", port=port)
