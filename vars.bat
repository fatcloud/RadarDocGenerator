@echo off
SET CondaPath=C:\tools\miniconda3
SET PolicyPath=aaaa
SET envpath="%~dp0.env"
if exist %envpath% (
    CALL %CondaPath%\Scripts\activate %envpath%
) else (
    CALL %CondaPath%\Scripts\activate
)
