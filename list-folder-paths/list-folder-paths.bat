
@echo off
set /p folder=Zadej cestu ke slozce: 
powershell -command "Get-ChildItem -Path '%folder%' -Recurse | Select-Object -ExpandProperty FullName | Out-File '%USERPROFILE%\Desktop\vypis.txt' -Encoding utf8"
echo Hotovo
pause
