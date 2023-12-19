@echo off

echo ===========================================
echo +                                         +
echo +   PDFs to PNGs, PNGs be inserted Docx   +
echo +                                         +
echo ===========================================
echo:
echo:
echo:
python -m venv .env
call .\.env\Scripts\activate.bat
echo PLEASE WAIT! PROCESSING ...
python -m pip install -r requirements.txt
python process.py
echo:
echo:
echo:
echo ===========================================
echo +                                         +
echo +        Your process is finished         +
echo +        Check data into doc_here         +
echo +                                         +
echo ===========================================
call .\.env\Scripts\deactivate.bat

pause