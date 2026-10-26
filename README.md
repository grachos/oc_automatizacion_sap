# Creación automática de OCs de terceros en SAP

Aplicación de escritorio (Windows, Python + tkinter) que crea en SAP (ME21N) las órdenes de compra de
transporte de terceros a partir de un Excel, usando SAP GUI Scripting. Al final valida que el valor de
las OCs coincida con el Excel. Instructivo de usuario: [`docs/Instructivo_Creacion_OCs_SAP.pdf`](docs/Instructivo_Creacion_OCs_SAP.pdf).

## Usuarios finales
Usar `CrearOCsSAP.exe` (se publica como *release*, no está en el código). Las credenciales se configuran
desde **Menú > Configuración** y la plantilla se baja desde **Menú > Descargar plantilla de Excel**.

## Requisitos
- SAP GUI con scripting habilitado: servidor (RZ11) `sapgui/user_scripting = TRUE`; cliente: SAP Logon >
  Opciones > Accesibilidad y scripting > Habilitar scripting.
- Para ejecutar desde código: Python 3.10+ y `pip install -r requirements.txt`, luego `run.bat`.

## Generar el .exe
`build.bat` (usa PyInstaller) deja `dist\CrearOCsSAP.exe`. El `.env` se guarda junto al .exe.

## Excel de entrada
Columnas: `NIT PROVEEDOR`, `VALOR OC`, `NUMERO LIQUIDACION` (ver `plantilla_ocs.xlsx`).

## Qué hace por fila (ME21N)
Proveedor → Dat.org. (DA00 / 025 / DA25) → "Pago en" 7 días con Cond. pago en blanco → posición (cat. F,
imput. 1, TRANSPORTE, 1 UP, centro, grupo TRANSPORT, Fe.entrega = hoy + 15) → servicio 4220 por el valor,
IVA VC → cuenta de mayor 2815052000 → texto de cabecera (digitado en la app) → texto de posición
(= liquidación) → Grabar. Los parámetros se editan en **Menú > Configuración**.

## Interrupciones
Cada fila se registra en `<excel>_avance.csv` (`EN_PROCESO` / `OK` / `ERROR`) y el `_resultado.xlsx` se
actualiza por fila. Al cargar de nuevo el mismo Excel se omiten las filas ya creadas y las interrumpidas
quedan en **REVISAR**.

## Estructura
`app.py` UI · `po_creator.py` flujo ME21N · `sap_gui.py` utilidades COM · `config.py` / `settings_ui.py`
configuración · `excel_io.py` lectura/plantilla/resultado · `journal.py` bitácora · `docs/` instructivo y
grabación de referencia (`grabacion_me21n.vbs`, hecha en QAS).

No subir nunca el `.env` (credenciales).
