@echo off
python -m pip install python-docx openpyxl psutil
set /p FilePath=Please enter the file path(if you don't enter, the default path is the current python file directory): 
python ./XunFeiWord2AnkiFiles.py %FilePath%
pause