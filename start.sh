#!/bin/bash
# Iniciar display virtual para LibreOffice
Xvfb :0 -screen 0 1024x768x24 &
export DISPLAY=:0

# Dar un segundo para que Xvfb arranque
sleep 1

# Iniciar el servidor web con gunicorn (modo producción)
exec gunicorn app:app \
  --bind 0.0.0.0:${PORT:-8000} \
  --workers 2 \
  --timeout 300 \
  --keep-alive 5 \
  --log-level info
