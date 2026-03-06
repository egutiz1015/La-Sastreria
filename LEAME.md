# ✂ La Sastrería — Sistema de Gestión
**Fundador:** Erick Gutierrez | **Administradora:** Keila Gutierrez

---

## 🚀 INSTALACIÓN Y EJECUCIÓN

### Windows
1. Instale **Python** desde https://www.python.org/downloads/
   - ✅ **MUY IMPORTANTE:** Marque la casilla **"Add Python to PATH"** al instalar
2. Copie la carpeta `La_Sastreria` a su computadora (ej: en el Escritorio)
3. Doble clic en **`INICIAR.bat`**
4. El navegador se abrirá en `http://localhost:5000`

### Mac / Linux
1. Instale Python si no lo tiene (en Mac viene preinstalado)
2. Abra Terminal en la carpeta del programa
3. Ejecute: `chmod +x iniciar_mac.sh && ./iniciar_mac.sh`

---

## 🔐 ACCESO AL SISTEMA

| Usuario | Contraseña | Rol |
|---------|-----------|-----|
| `erick` | `Sastreria2024!Erick` | Dueño & Fundador |
| `keila` | `Sastreria2024!Keila` | Administradora |

---

## 📁 ESTRUCTURA DE ARCHIVOS

```
La_Sastreria/
├── app.py              ← Programa principal
├── INICIAR.bat         ← Ejecutar en Windows
├── iniciar_mac.sh      ← Ejecutar en Mac/Linux
├── requirements.txt    ← Dependencias
├── data/
│   └── sastreria.db    ← Base de datos (se crea automáticamente)
├── exports/            ← Archivos Excel generados
└── templates/          ← Pantallas del sistema
```

---

## 📧 CONFIGURACIÓN DE CORREO (Gmail)

Para enviar correos automáticos:

1. En el sistema, ir a **⚙️ Configuración** (solo dueño)
2. Configurar Gmail:
   - Servidor: `smtp.gmail.com`
   - Puerto: `587`
   - Usuario: su correo de Gmail
   - Contraseña: **Contraseña de Aplicación** (NO su contraseña normal)

**Cómo obtener Contraseña de Aplicación en Gmail:**
1. Vaya a `myaccount.google.com`
2. Seguridad → Verificación en 2 pasos (debe estar activada)
3. Seguridad → Contraseñas de aplicaciones
4. Seleccione "Correo" y "Windows/Mac"
5. Copie la contraseña de 16 caracteres generada

---

## 💡 USO DEL SISTEMA

### Crear una Orden
1. Clic en **➕ Nueva Orden**
2. Busque cliente existente O ingrese datos nuevo cliente
3. Establezca fechas de orden y entrega
4. Agregue prendas del catálogo o manualmente
5. Para prendas con **precio fijo** (ej: arreglo ruedo = Q25) → tipo "Fijo"
6. Para prendas con **precio variable** (traje a medida) → tipo "Variable" e ingrese el precio
7. Clic **Guardar Orden** → se genera Excel automáticamente y se envía correo

### Marcar Entregas
- En el detalle de cada orden: clic **✓ Entregar** en cada prenda
- O cambiar estado completo de la orden desde el panel lateral
- El calendario muestra todas las fechas de entrega en colores

### Exportar
- Cada orden tiene botón **📥 Descargar Excel**
- En el menú: **📥 Exportar Clientes Excel** exporta toda la base de clientes

---

## 🔧 CAMBIAR CONTRASEÑAS

Abra `app.py` con el Bloc de Notas y busque la sección `USERS`.
Cambie los valores de contraseña en las líneas:
```python
'password': hashlib.sha256('SuNuevaContraseña'.encode()).hexdigest(),
```

---

## ☎ SOPORTE
El sistema guarda todos los datos en `data/sastreria.db`.
**Haga respaldo de esta carpeta regularmente.**
