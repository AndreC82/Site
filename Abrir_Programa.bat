@echo off
cd /d "%~dp0"
echo Iniciando o programa de levantamento...
echo (O navegador vai abrir sozinho em alguns segundos.)
echo.
echo NAO FECHE esta janela enquanto estiver usando o programa.
echo.
python -m pdf_takeoff.webapp
echo.
echo O programa foi encerrado. Se apareceu algum erro acima, tire um print e mande.
pause
