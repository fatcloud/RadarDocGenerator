SET CondaPath=C:\tools\miniconda3
SET PolicyPath=.
SET envpath=".\.env"
CALL %CondaPath%\Scripts\activate %envpath% && python .\gendoc.py %PolicyPath% && pause