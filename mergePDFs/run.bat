@echo off
echo Install requirements...
python -m pip install -r requirements.txt

echo Create pdf_merger.exe...
pyinstaller main.py --clean --onefile --name pdf_merger -y

python scripts/file_transfer.py