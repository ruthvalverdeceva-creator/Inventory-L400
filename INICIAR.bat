@echo off
REM Cambia esta ruta a tu carpeta de OneDrive sincronizada (ahi se guarda el Excel actualizado)
set OUT="C:\Users\TuUsuario\OneDrive - Comunidad UPN\CEVA_Picking_Salida.xlsx"
pip install openpyxl
python server.py
pause
