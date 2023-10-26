call .\venv\Scripts\activate.bat

python3 -m PyInstaller -w -c -F -i "icon.ico" --paths venv/Lib/site-package --copy-metadata pikepdf --copy-metadata ocrmypdf --collect-submodules ocrmypdf --collect-datas ocrmypdf.data Extracao_Lights.py -n Extracao_Lights

pause