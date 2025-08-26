# Notificador de Avances de Proyectos (Aktivgroup)

Este prototipo permite enviar correos de estado de proyecto a cada PM Cliente de forma masiva, a partir de un Excel.

## Requisitos
- **Python 3.11/3.12** instalado.
- Acceso SMTP (Gmail con **contraseña de aplicación** o Outlook/M365).
- VS Code (opcional, recomendado).

## Instalación rápida

### Windows (PowerShell)
```powershell
# 1) Crear y activar entorno virtual
py -3.12 -m venv .venv
.\.venv\Scripts\activate

# 2) Actualizar pip e instalar dependencias
python -m pip install --upgrade pip
pip install pandas openpyxl

# 3) (opcional) Guardar dependencias
pip freeze > requirements.txt
```

### macOS / Linux
```bash
# 1) Crear y activar entorno virtual
python3 -m venv .venv
source .venv/bin/activate

# 2) Actualizar pip e instalar dependencias
python -m pip install --upgrade pip
pip install pandas openpyxl

# 3) (opcional) Guardar dependencias
pip freeze > requirements.txt
```

## Cómo usar
1. Abre/edita `aktivgroup_notificador_proyectos.xlsx` y completa una fila por proyecto.  
2. Ejecuta la app:

   **Windows**
   ```powershell
   .\.venv\Scripts\activate
   python notificador_proyectos.py
   ```

   **macOS / Linux**
   ```bash
   source .venv/bin/activate
   python notificador_proyectos.py
   ```

3. En la ventana:
   - Selecciona el Excel.
   - Escribe tu correo **remitente** y la **contraseña de aplicación** (Gmail/Outlook).
   - Ajusta **Asunto** y **Cuerpo (HTML)** con los placeholders:
     `{Cliente}`, `{Proyecto}`, `{pct}`, `{hitos}`, `{riesgos}`, `{proximos}`, `{fecha_entrega}`, `{pm_cliente_nombre}`, `{pm_aktiv_nombre}`
   - Usa **Vista previa** y luego **Enviar correos**.

> Recomendación: usa **contraseñas de aplicación** (no tu contraseña normal) y, si es posible, una cuenta de servicio.

## Configuración SMTP (referencia)
- **Gmail**: host `smtp.gmail.com`, puerto `587`, STARTTLS, *contraseña de aplicación*.  
- **Outlook / Microsoft 365**: host `smtp.office365.com`, puerto `587`, STARTTLS.

## Columnas requeridas del Excel
- Cliente  
- Proyecto  
- % Avance  
- Hitos cumplidos (última semana)  
- Bloqueos / Riesgos  
- Próximos pasos  
- Fecha próxima entrega  
- PM Cliente (Nombre)  
- Correo PM Cliente  
- PM Aktivgroup (Nombre)  
- Correo PM Aktivgroup  

## Librerías
- `pandas`, `openpyxl`  
- `tkinter` (incluida con Python en Windows/macOS estándar)  
- `smtplib`, `email` (estándar de Python)  
- (opcional) `reportlab` para PDF

## (Opcional) Generar ejecutable
```powershell
# Instalar PyInstaller
pip install pyinstaller

# Crear ejecutable (carpeta dist/)
pyinstaller --onefile notificador_proyectos.py
```

## Licencia
MIT
