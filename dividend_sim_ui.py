import tkinter as tk
from tkinter import messagebox
import re
from datetime import datetime
from pathlib import Path
import sys

# השתקת print כשנבנה עם --noconsole (מונע שגיאת flush)
class _Silent:
    def write(self, *a, **k): pass
    def flush(self): pass
if getattr(sys, "stdout", None) is None: sys.stdout = _Silent()
if getattr(sys, "stderr", None) is None: sys.stderr = _Silent()

# תיקיית הריצה: אם EXE → תיקיית ה-EXE, אחרת תיקיית הקובץ
def get_run_dir():
    return (Path(sys.executable).parent if getattr(sys, "frozen", False)
            else Path(__file__).parent.resolve())

# ייבוא ישיר של פונקציית הייצוא
from dividend_sim import generate_dividend_excel

class DividendSimulatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Dividend Simulator")
        self.root.geometry("400x600")
        self.root.configure(bg="#F3F4F6")

        self.main_frame = tk.Frame(root, bg="white", bd=2, relief="flat")
        self.main_frame.place(relx=0.5, rely=0.5, anchor="center", width=360, height=560)
        self.main_frame.configure(highlightbackground="#D1D5DB", highlightthickness=2)

        tk.Label(self.main_frame, text="Dividend Simulator",
                 font=("Helvetica", 16, "bold"), bg="white", fg="#1F2937").pack(pady=20)

        self.form_frame = tk.Frame(self.main_frame, bg="white")
        self.form_frame.pack(padx=20, fill="x")

        # Ticker
        tk.Label(self.form_frame, text="Ticker Symbol", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.ticker_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                     relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.ticker_entry.insert(0, "e.g., AAPL")
        self.ticker_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.ticker_entry, "e.g., AAPL"))
        self.ticker_entry.pack(fill="x", pady=5)

        # Start date
        tk.Label(self.form_frame, text="Start Date (DD.MM.YYYY)", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.start_date_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                         relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.start_date_entry.insert(0, "e.g., 01.01.2024")
        self.start_date_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.start_date_entry, "e.g., 01.01.2024"))
        self.start_date_entry.pack(fill="x", pady=5)

        # End date
        tk.Label(self.form_frame, text="End Date (DD.MM.YYYY)", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.end_date_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                       relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.end_date_entry.insert(0, "e.g., 31.12.2025")
        self.end_date_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.end_date_entry, "e.g., 31.12.2025"))
        self.end_date_entry.pack(fill="x", pady=5)

        # Shares
        tk.Label(self.form_frame, text="Shares (default: 1000)", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.shares_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                     relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.shares_entry.insert(0, "e.g., 1000")
        self.shares_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.shares_entry, "e.g., 1000"))
        self.shares_entry.pack(fill="x", pady=5)

        # Exchange rate
        tk.Label(self.form_frame, text="Exchange Rate (default: 3.69)", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.exchange_rate_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                            relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.exchange_rate_entry.insert(0, "e.g., 3.69")
        self.exchange_rate_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.exchange_rate_entry, "e.g., 3.69"))
        self.exchange_rate_entry.pack(fill="x", pady=5)

        # Tax rate
        tk.Label(self.form_frame, text="Tax Rate (default: 0.25)", font=("Helvetica", 10), bg="white", fg="#374151").pack(anchor="w")
        self.tax_rate_entry = tk.Entry(self.form_frame, font=("Helvetica", 10), bg="#F9FAFB",
                                       relief="flat", bd=1, highlightthickness=1, highlightbackground="#D1D5DB")
        self.tax_rate_entry.insert(0, "e.g., 0.25")
        self.tax_rate_entry.bind("<FocusIn>", lambda e: self._clear_placeholder(self.tax_rate_entry, "e.g., 0.25"))
        self.tax_rate_entry.pack(fill="x", pady=5)

        self.generate_button = tk.Button(self.form_frame, text="Generate Excel Report",
                                         font=("Helvetica", 10, "bold"),
                                         bg="#4F46E5", fg="white", activebackground="#4338CA",
                                         relief="flat", bd=0, command=self.generate_report)
        self.generate_button.pack(fill="x", pady=20)

        self.result_label = tk.Label(self.main_frame, text="", font=("Helvetica", 10), bg="white", fg="#374151", wraplength=320)
        self.result_label.pack(pady=10)

    def _clear_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)

    def generate_report(self):
        ticker = self.ticker_entry.get().strip().upper()
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
        shares = self.shares_entry.get().strip()
        exchange_rate = self.exchange_rate_entry.get().strip()
        tax_rate = self.tax_rate_entry.get().strip()

        date_regex = r"^\d{2}\.\d{2}\.\d{4}$"
        if not ticker or ticker == "e.g., AAPL":
            self.result_label.config(text="Please enter a ticker symbol.", fg="#DC2626")
            return
        if not re.match(date_regex, start_date) or not re.match(date_regex, end_date):
            self.result_label.config(text="Invalid date format. Use DD.MM.YYYY (e.g., 01.01.2024).", fg="#DC2626")
            return
        try:
            if shares and shares != "e.g., 1000": int(shares)
            if exchange_rate and exchange_rate != "e.g., 3.69": float(exchange_rate)
            if tax_rate and tax_rate != "e.g., 0.25": float(tax_rate)
        except ValueError:
            self.result_label.config(text="Shares must be int; Exchange/Tax must be numbers.", fg="#DC2626")
            return

        self.result_label.config(text="Generating report...", fg="#4B5563")
        try:
            shares_i = int(shares) if shares and shares != "e.g., 1000" else 1000
            ex_i = float(exchange_rate) if exchange_rate and exchange_rate != "e.g., 3.69" else 3.69
            tax_i = float(tax_rate) if tax_rate and tax_rate != "e.g., 0.25" else 0.25

            start_iso = datetime.strptime(start_date, "%d.%m.%Y").strftime("%Y-%m-%d")
            end_iso   = datetime.strptime(end_date,   "%d.%m.%Y").strftime("%Y-%m-%d")

            # כאן נקבע נתיב היעד – תיקיית ה-EXE/הקובץ
            run_dir = get_run_dir()
            output_path = str(run_dir / f"{ticker}_Sim.xlsx")

            generate_dividend_excel(
                ticker=ticker,
                start_date=start_iso,
                end_date=end_iso,
                output_file=output_path,
                shares=shares_i,
                exchange_rate=ex_i,
                tax_rate=tax_i
            )
            self.result_label.config(text=f"Report generated:\n{output_path}", fg="#16A34A")
        except Exception as e:
            self.result_label.config(text=f"Error: {str(e)}", fg="#DC2626")

if __name__ == "__main__":
    root = tk.Tk()
    app = DividendSimulatorApp(root)
    root.mainloop()
