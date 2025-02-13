@echo off
echo Setting up the project...
python -m venv .venv
echo .venv created


call .venv\Scripts\activate.bat
echo .venv activated.

python.exe -m pip install --upgrade pip
echo pip upgrade -- done

echo Installing dependencies...
pip install -r requirements.txt
echo requirements installed


echo Creating .gitignore file...

echo .venv/ > .gitignore
:: Create .gitignore and add the .venv directory
echo __pycache__/ >> .gitignore
echo *.pyc >> .gitignore
echo *.pyo >> .gitignore
echo *.log >> .gitignore
::echo db.sqlite3 >> .gitignore
echo .env >> .gitignore
echo .gitignore file created.

echo __pycache__/ >> .gitignore
echo *.pyc >> .gitignore
echo *.pyo >> .gitignore
echo *.log >> .gitignore
echo db.sqlite3 >> .gitignore
echo .env >> .gitignore




echo.
echo Setup complete! Now you can run the server using run_server.bat

pause
