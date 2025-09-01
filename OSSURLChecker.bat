@echo off
chcp 65001 > nul
echo.

::banner
python3 banner.py
echo.

::demo url
python3 KeyExtract.py
echo.

::extract the Hosts of Excel
python3 ExtractHost.py
echo.

:delete 
del /q "*result*.xlsx"
echo.

::Checking
python3 OSSURLChecker.py
echo.
pause