:: Setup and switch to the virtual environment
SET CondaPath=C:\tools\miniconda3
SET envpath=".\.env"
CALL %CondaPath%\Scripts\activate
CALL conda env create -f environment.yml -p %envpath% --force
CALL conda clean --all -y
PAUSE
