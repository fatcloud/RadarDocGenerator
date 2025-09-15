@echo off
CALL vars.bat

REM -- Activate Conda environment and run the script
CALL "%CondaPath%\Scripts\activate.bat" %envpath%
python.exe .\diagnosis.py "%PolicyPath%"
pause
