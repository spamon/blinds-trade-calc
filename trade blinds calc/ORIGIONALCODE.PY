import pandas as pd

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

def get_price(excel_data):
    sheet_names = excel_data.sheet_names
    
    # Show available pricing bands (sheet names)
    print("Available Pricing Bands:")
    for i, sheet_name in enumerate(sheet_names, start=1):
        print(f"{i}. {sheet_name}")
    
    # Ask the user to select a pricing band
    while True:
        try:
            choice = int(input("Select a pricing band (enter the number): "))
            if 1 <= choice <= len(sheet_names):
                selected_sheet = sheet_names[choice - 1]
                break
            else:
                print("Invalid selection. Please enter a number from the list above.")
        except ValueError:
            print("Invalid input. Please enter a number.")
    
    # Load the data from the selected sheet
    sheet_data = excel_data.parse(selected_sheet)
    
    # Ask the user to enter the width and drop measurements
    while True:
        try:
            width = int(input("Enter the width in mm: "))
            drop = int(input("Enter the drop in mm: "))
            break
        except ValueError:
            print("Invalid input. Please enter a valid number.")
    
    # Find and display the price
    price = find_price(width, drop, sheet_data)
    print(f"The price for the selected measurements is: {price}")

def select_blind_type():
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
    
    print("Available Blind Types:")
    for i, blind_type in enumerate(blind_types.keys(), start=1):
        print(f"{i}. {blind_type}")
    
    while True:
        try:
            choice = int(input("Select a blind type (enter the number): "))
            if 1 <= choice <= len(blind_types):
                selected_blind = list(blind_types.keys())[choice - 1]
                return blind_types[selected_blind]
            else:
                print("Invalid selection. Please enter a number from the list above.")
        except ValueError:
            print("Invalid input. Please enter a number.")

if __name__ == "__main__":
    blind_file_name = select_blind_type()
    excel_data = load_excel_data(blind_file_name)
    get_price(excel_data)