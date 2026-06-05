@echo off
chcp 65001 >nul
cls
echo.
echo  ██████████████████████████████████████████████████
echo  ██                                              ██
echo  ██      GENERADOR DE AR  -  Asistente HYST      ██
echo  ██                                              ██
echo  ██████████████████████████████████████████████████
echo.

:: ── 1. Verificar Python ──────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python no esta instalado.
    echo.
    echo  Instala Python desde:  https://www.python.org/downloads/
    echo  Cuando lo instales, TILDA la opcion "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
echo  [OK] Python encontrado.

:: ── 2. Instalar dependencias ─────────────────────────
echo  [..] Instalando librerias necesarias...
pip install -r requirements.txt -q --disable-pip-version-check
if errorlevel 1 (
    echo  [ERROR] No se pudieron instalar las librerias.
    echo  Asegurate de tener conexion a internet e intentá de nuevo.
    pause
    exit /b 1
)
echo  [OK] Librerias instaladas.

:: ── 3. Verificar plantilla ───────────────────────────
if not exist plantilla.xlsm (
    echo.
    echo  [AVISO] No encontre la plantilla oficial ^(plantilla.xlsm^).
    echo  Copiala en esta carpeta antes de continuar.
    echo.
    pause
)

:: ── 4. Clave de IA ───────────────────────────────────
if "%ANTHROPIC_API_KEY%"=="" (
    echo.
    echo  ────────────────────────────────────────────────────
    echo   Para usar la IA necesitas una clave personal.
    echo   La obtienes gratis en: https://console.anthropic.com
    echo   (Hace falta crear una cuenta, es rapido)
    echo  ────────────────────────────────────────────────────
    echo.
    set /p ANTHROPIC_API_KEY="  Pega tu clave aqui y presiona Enter: "
    echo.
)

:: ── 5. Arrancar el servidor ──────────────────────────
echo  [OK] Todo listo. Abriendo el sistema...
echo.
echo  ► El sistema esta corriendo en:  http://localhost:8000
echo  ► Para cerrarlo: cerrá esta ventana o presiona Ctrl+C
echo.
echo  ────────────────────────────────────────────────────────
echo   Comparti esta direccion con los tecnicos en tu red:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /R "IPv4"') do (
    set IP=%%a
    goto :mostrar_ip
)
:mostrar_ip
echo   http://%IP: =% :8000
echo  ────────────────────────────────────────────────────────
echo.

timeout /t 2 >nul
start http://localhost:8000
python app.py

pause
