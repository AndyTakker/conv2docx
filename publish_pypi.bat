@echo off
chcp 1251 >nul
REM Публикация на PyPi
pip install twine
twine upload dist/*
pause
