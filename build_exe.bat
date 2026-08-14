@echo off
echo ============================================
echo   Qarz Nazorat Tizimi - .exe yig'ish
echo ============================================

echo [0/3] Eski build/dist papkalari tozalanmoqda...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q QarzNazorat.spec 2>nul

echo [1/3] Kerakli kutubxonalar o'rnatilmoqda...
python -m pip install -r requirements.txt

echo [2/3] .exe yig'ilmoqda (bir necha daqiqa vaqt olishi mumkin)...
pyinstaller --noconfirm --onefile --windowed ^
    --name "QarzNazorat" ^
    --add-data "templates;templates" ^
    --collect-all "pyxlsb" ^
    --collect-all "openpyxl" ^
    main.py

echo [3/3] Tayyor!
echo Natija: dist\QarzNazorat.exe
pause
