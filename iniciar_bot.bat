@echo off
title Gerador de Relatorios TRO
echo ===================================================
echo     Preparando o Ambiente do Bot (Streamlit)...
echo ===================================================

echo Instalando bibliotecas (se necessario)...
pip install -r requirements.txt -q

echo.
echo Iniciando o sistema no seu navegador...
streamlit run app.py

pause
