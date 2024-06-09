@echo off
echo Install requirements...
python -m pip install -r requirements.txt

echo Create pdf_merger.exe...
pyinstaller main.py --noconsole --clean --onefile --name pdf_splitter -y

python scripts/file_transfer.py