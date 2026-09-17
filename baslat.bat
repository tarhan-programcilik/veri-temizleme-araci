@echo off
chcp 65001 >nul
title Veri Temizleme ve Özetleme Aracı
cd /d "%~dp0"
echo ========================================================
echo   Veri Temizleme ve Özetleme Aracı Başlatılıyor...
echo ========================================================
streamlit run app.py
pause
