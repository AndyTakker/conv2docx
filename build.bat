@echo off
chcp 1251 >nul
REM Сборка дистрибутива
rmdir /s /q dist
rmdir /s /q conv2docx.egg-info
python -m build 
pause
