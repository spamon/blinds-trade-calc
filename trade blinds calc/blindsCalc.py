import pandas as pd
import tkinter as tk
from tkinter import ttk

def load_excel_data(file_name):
    return pd.ExcelFile(file_name)

def find_price(width, drop, sheet_data):
    # Drop the first row and first two columns to get only the price data
    prices = sheet_data.iloc[1:, 2:]
    
    # Convert the 'Drop' and width columns to numeric for comparison
    drops = pd.to_numeric(sheet_data['mm'][1:])
    widths = pd.to_numeric(sheet_data.columns[2:])

    # Find the closest drop and width categories
    closest_drop = drops[drops >= drop].min()
    closest_width = widths[widths >= width].min()

    # If there is no higher or equal category available, return a message
    if pd.isna(closest_drop) or pd.isna(closest_width):
        return "The provided measurements are out of the available range."

    # Find the price based on the closest drop and width
    price = prices.loc[sheet_data['mm'] == closest_drop, closest_width].values[0]
    return price

def populate_price_bands(selected_blind):
    selected_file = blind_types[selected_blind]
    excel_data = load_excel_data(selected_file)
    sheet_names = excel_data.sheet_names
    selected_band_combobox['values'] = sheet_names

def get_price():
    selected_blind = selected_blind_var.get()
    selected_band = selected_band_var.get()
    width = int(width_entry.get())
    drop = int(drop_entry.get())

    try:
        excel_data = load_excel_data(blind_types[selected_blind])
        sheet_data = excel_data.parse(selected_band)
        price = find_price(width, drop, sheet_data)
        result_label.config(text=f"The price for the selected measurements is: {price}")
    except Exception as e:
        result_label.config(text="Error: " + str(e))

def clear_fields():
    selected_blind_var.set("")
    selected_band_var.set("")
    width_entry.delete(0, "end")
    drop_entry.delete(0, "end")
    result_label.config(text="")

blind_types = {
    "Vertical Blinds": "Vertical Blinds.xlsx",
    "Roller Blinds": "roller blinds.xlsx",
    "Open Cassette Roller Blinds": "open cassette roller blinds.xlsx",
    "Mirage Blinds": "mirage blinds.xlsx",
    "Vision Blinds": "vision blinds.xlsx",
    "Senses Roller Blinds": "senses roller blinds.xlsx",
    "Allusion Blinds": "allusion blinds.xlsx",
    "Perfect Fit Roller Blinds": "perfect fit roller blinds.xlsx",
    "Perfect Fit Pleated Blinds": "perfect fit pleated blinds.xlsx",
    "Roman Blinds": "roman blinds.xlsx",
    "PVC Shutters": "pvc shutters.xlsx",
    "Perfect Fit Shutters": "perfect fit shutters.xlsx"
}

root = tk.Tk()
root.title("Blind Price Calculator")

selected_blind_var = tk.StringVar()
selected_band_var = tk.StringVar()

selected_blind_label = ttk.Label(root, text="Select Blind Type:")
selected_blind_label.grid(row=0, column=0, padx=10, pady=5)

selected_blind_combobox = ttk.Combobox(root, textvariable=selected_blind_var, values=list(blind_types.keys()))
selected_blind_combobox.grid(row=0, column=1, padx=10, pady=5)
selected_blind_combobox.bind("<<ComboboxSelected>>", lambda event: populate_price_bands(selected_blind_var.get()))

selected_band_label = ttk.Label(root, text="Select Price Band:")
selected_band_label.grid(row=1, column=0, padx=10, pady=5)

selected_band_combobox = ttk.Combobox(root, textvariable=selected_band_var)
selected_band_combobox.grid(row=1, column=1, padx=10, pady=5)

width_label = ttk.Label(root, text="Enter Width (mm):")
width_label.grid(row=2, column=0, padx=10, pady=5)

width_entry = ttk.Entry(root)
width_entry.grid(row=2, column=1, padx=10, pady=5)

drop_label = ttk.Label(root, text="Enter Drop (mm):")
drop_label.grid(row=3, column=0, padx=10, pady=5)

drop_entry = ttk.Entry(root)
drop_entry.grid(row=3, column=1, padx=10, pady=5)

calculate_button = ttk.Button(root, text="Calculate Price", command=get_price)
calculate_button.grid(row=4, column=0, columnspan=2, padx=10, pady=5)

result_label = ttk.Label(root, text="")
result_label.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

clear_button = ttk.Button(root, text="Clear Fields", command=clear_fields)
clear_button.grid(row=6, column=0, columnspan=2, padx=10, pady=5)

root.mainloop()
