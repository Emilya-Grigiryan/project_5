import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import pandas as pd
import requests
import xlsxwriter  # Ensure xlsxwriter is imported

class CryptoInfoApp:
    """
    A simple application to generate Excel files with cryptocurrency information.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Crypto Information App")
        self.root.geometry("400x300")

        self.file_path = ""
        self.output_directory = os.path.expanduser("~") + "/Downloads/"

        self.label_info = tk.Label(self.root, text="Upload a file containing cryptocurrency symbols (.txt):")
        self.label_info.pack(pady=10)

        self.btn_browse = tk.Button(self.root, text="Browse", command=self.browse_file)
        self.btn_browse.pack(pady=5)

        self.btn_generate = tk.Button(self.root, text="Generate Excel", command=self.generate_excel)
        self.btn_generate.pack(pady=10)

        self.label_directory = tk.Label(self.root, text="Select output directory (default: Downloads):")
        self.label_directory.pack(pady=5)

        self.btn_directory = tk.Button(self.root, text="Choose Directory", command=self.choose_directory)
        self.btn_directory.pack(pady=5)

    def browse_file(self):
        """Open file dialog to select a text file containing cryptocurrency symbols."""
        self.file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if self.file_path:
            messagebox.showinfo("File Selected", f"Selected File: {self.file_path}")

    def choose_directory(self):
        """Open directory dialog to select output directory."""
        self.output_directory = filedialog.askdirectory()
        if self.output_directory:
            messagebox.showinfo("Directory Selected", f"Output Directory: {self.output_directory}")

    def generate_excel(self):
        """Generate Excel file with cryptocurrency information."""
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return

        try:
            with open(self.file_path, 'r', encoding='utf-8') as file:
                symbols = [line.strip() for line in file.readlines() if line.strip()]

            if not symbols:
                messagebox.showerror("Error", "File is empty or does not contain valid symbols.")
                return

            data = self.fetch_crypto_data(symbols)

            if not data:
                messagebox.showerror("Error", "Failed to fetch data from the API.")
                return

            df = pd.DataFrame(data)

            excel_filename = filedialog.asksaveasfilename(
                initialdir=self.output_directory,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

            if excel_filename:
                self.save_to_excel(df, excel_filename)
        except FileNotFoundError:
            messagebox.showerror("Error", "File not found. Please check the file path.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def fetch_crypto_data(self, symbols):
        """Fetch cryptocurrency data from an API."""
        api_key = "your_api_key_here"  # Replace with your API key
        base_url = "https://api.coincap.io/v2/assets"

        data = []
        for symbol in symbols:
            try:
                response = requests.get(f"{base_url}/{symbol}", params={"key": api_key})
                if response.status_code == 200:
                    crypto_data = response.json()["data"]
                    data.append({
                        "Name": crypto_data["name"],
                        "Symbol": crypto_data["symbol"],
                        "Current Price": float(crypto_data["priceUsd"]),
                        "Market Cap": float(crypto_data["marketCapUsd"]),
                        "Total Volume": float(crypto_data["volumeUsd24Hr"]),
                        "Price Change (24h)": float(crypto_data["changePercent24Hr"]),
                        "Last Updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })
                else:
                    messagebox.showwarning("Warning", f"Failed to fetch data for symbol: {symbol}")
            except requests.RequestException as e:
                messagebox.showerror("Error", f"Network error while fetching data for {symbol}: {str(e)}")
            except KeyError:
                messagebox.showerror("Error", f"Unexpected data format for symbol {symbol}.")
            except Exception as e:
                messagebox.showerror("Error", f"Error fetching data for symbol {symbol}: {str(e)}")

        return data

    def save_to_excel(self, df, filename):
        """Save DataFrame to Excel file using xlsxwriter engine."""
        try:
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']  # Replace 'Sheet1' with your sheet name if different
                header_format = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE'})
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(0, len(df.columns) - 1, 18)  # Set column width for all columns
                messagebox.showinfo("Success", f"Excel file saved successfully: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving Excel file: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CryptoInfoApp(root)
    root.mainloop()

