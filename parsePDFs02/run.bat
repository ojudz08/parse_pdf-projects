@echo off
echo Install requirements...
python -m pip install -r requirements.txt

echo Create pdf_parser.exe...
pyinstaller main.py --clean --onefile --name pdf_parser -y

python scripts/file_transfer.py