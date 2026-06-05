# Generador de AR con IA · Manual de instalación

## ¿Qué hace esto?
El técnico entra a la página, pega el procedimiento (PE) o carga las fases a mano,
confirma los riesgos detectados por la IA y descarga el Análisis de Riesgo completo
en Excel (.xlsm con tu plantilla oficial) y en PDF listo para firmar.

---

## Lo que necesitás instalar en el servidor (una sola vez)

### 1. Python 3.10+
https://www.python.org/downloads/

### 2. LibreOffice (para recalcular fórmulas y exportar PDF)
- Windows: https://www.libreoffice.org/download/libreoffice/
- Linux:   sudo apt install libreoffice

### 3. Dependencias Python
Abrí una terminal en esta carpeta y ejecutá:
```
pip install -r requirements.txt
```

---

## Configuración

### 1. Colocá tu plantilla
Reemplazá el archivo `plantilla.xlsm` con tu plantilla oficial (la que venía de ejemplo).
Asegurate de que tenga las mismas hojas: INSTALACIÓN, FASES DE TAREA, RIESGOS, IPER, IMPRIMIR.

### 2. Clave de IA (Anthropic)
Conseguís la clave en https://console.anthropic.com → API Keys → Create key.
Luego la configurás así (en la terminal, antes de arrancar el servidor):

- Windows:   set ANTHROPIC_API_KEY=sk-ant-TUCLAVEAQUI
- Linux/Mac: export ANTHROPIC_API_KEY=sk-ant-TUCLAVEAQUI

---

## Arrancar el servidor

```
python app.py
```

Abrís el navegador en: http://localhost:8000

Si querés que otros en la red accedan (técnicos desde el celular en obra):
El servidor ya escucha en 0.0.0.0, así que alcanza con que usen
la IP de tu computadora: http://192.168.X.X:8000

---

## Estructura de carpetas

```
backend/
  app.py               ← el servidor principal
  engine_runner.py     ← conecta el motor con el servidor
  plantilla.xlsm       ← TU PLANTILLA OFICIAL (reemplazala)
  requirements.txt
  engine/
    generar_ar.py      ← motor que rellena la plantilla
    extraer_catalogo.py
    catalogo_riesgos.json   ← barreras y puntajes de Fine por riesgo
    riesgos_validos.json    ← los 28 tipos de riesgo del IPER
  frontend/
    index.html         ← la interfaz que ve el técnico
  scripts/
    recalc.py          ← recalcula fórmulas con LibreOffice
  outputs/             ← acá se guardan los AR generados (se crea automáticamente)
```

---

## Para llevarlo a internet (venderlo a empresas)

Con esto tenés la base. Para publicarlo online necesitás:
1. Un servidor Linux (ej. DigitalOcean, Vultr o AWS Light · desde USD 6/mes)
2. Un dominio (ej. generarar.com.ar)
3. SSL (HTTPS) → gratis con Let's Encrypt

Eso es un paso aparte cuando estés listo para comercializarlo.
El código no cambia; solo se configura el entorno del servidor.

---

## Importante · Responsabilidad legal

El sistema genera un **borrador** automático.
Todo AR debe ser revisado y firmado por el técnico de HYST habilitado antes de usarse.
El sistema acelera la tarea; la responsabilidad técnica sigue siendo del profesional.
