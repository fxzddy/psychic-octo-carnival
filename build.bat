@echo off
echo Installing...
pip install -r requirements.txt

echo Cleaning...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist __pycache__ rmdir /s /q __pycache__
del /q *.spec 2>nul

echo Packeting...
pyinstaller --onefile --windowed ^
    --name="Face" ^
    --icon=icon.ico ^
    --add-data="requirements.txt;." ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=PyQt5 ^
    --hidden-import=matplotlib ^
    --hidden-import=matplotlib.backends.backend_qt5agg ^
    --hidden-import=numpy ^
    --hidden-import=xlrd ^
    --exclude-module=tkinter ^
    --exclude-module=PyQt4 ^
    --exclude-module=PySide ^
    --clean ^
    main.py

echo Sucess!
echo File: So.exe
pause