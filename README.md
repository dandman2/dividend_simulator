# dividend_simulator
Dividend Simulator App for windows that generates excel


# Create Virtual Environment
python -m venv .venv
.venv\Scripts\activate
python.exe -m pip install --upgrade pip
pip install -U pip
pip install pyinstaller yfinance pandas openpyxl


# Run the app
python.exe .\dividend_sim_ui.py


# Build EXE file
cd dividend_simulator
.venv\Scripts\activate
pyinstaller --onefile --icon=icon.ico --noconsole --name DividendSimulator dividend_sim_ui.py
