@echo off
title Gerador de Relatorios TRO
echo ===================================================
echo     Preparando o Ambiente do Bot (Streamlit)...
echo ===================================================

:: Instala as dependencias caso alguem ainda nao tenha instalado
echo Verificando atualizacoes e bibliotecas...
pip install -r requirements.txt -q

:: Inicia o aplicativo web
echo.
echo Iniciando o sistema no seu navegador...
streamlit run app.py

pause
