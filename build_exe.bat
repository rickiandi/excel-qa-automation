@echo off
python -m pip install -r requirements.txt
python -m pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name ExcelQA_DataEntry excel_qa_dataentry.py
pause
