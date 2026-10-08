Set-Location $PSScriptRoot
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --log-level WARN --onefile --windowed --name MergeExcel --specpath build app.py
exit $LASTEXITCODE
