@echo off
REM Batch-Datei zum Starten des PDF-Konverter-Skripts

REM Pfad zum Python-Interpreter
set PYTHON_PATH=python

REM Pfad zum Skript
set SCRIPT_PATH=pdf_converter.py

REM Ausführen des Skripts
%PYTHON_PATH% %SCRIPT_PATH%

REM Optional: Warten, bevor das Fenster geschlossen wird
pause
