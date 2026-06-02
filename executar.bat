@echo off

cd /d "C:\Users\ext.matheusmm\Documents\Envio de Emails"

call ".venv\Scripts\activate.bat"

python -u "run_pipeline.py"

cmd /k