import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import os

def select_save_location():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        save_location_entry.delete(0, tk.END)
        save_location_entry.insert(0, folder_selected)

def dec_to_bin(value, length):
    return bin(int(value))[2:].zfill(length)

def bin_to_hex(binary_str):
    hex_str = hex(int(binary_str, 2))[2:].upper()
    return hex_str.zfill(len(binary_str) // 4)

def generate_epc(upc, serial_number):
    # Extract GS1 Company Prefix, Item Reference Number, and Check Digit
    gs1_company_prefix = "0" + upc[:6]
    item_reference_number = upc[6:11]

    # Form GTIN-14
    gtin14 = "0" + gs1_company_prefix + item_reference_number

    # Binary SGTIN-96 parts
    header = "00110000"
    filter_value = "001"
    partition = "101"
    gs1_binary = dec_to_bin(gs1_company_prefix, 24)
    item_reference_binary = dec_to_bin(item_reference_number, 20)
    serial_binary = dec_to_bin(serial_number, 38)

    # Concatenate all parts to form the full binary SGTIN-96
    epc_binary = header + filter_value + partition + gs1_binary + item_reference_binary + serial_binary

    # Convert binary EPC to hexadecimal
    epc_hex = bin_to_hex(epc_binary)
    return epc_hex

def generate_file():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()
    end_serial = serial_end_entry.get().strip()
    save_location = save_location_entry.get().strip()

    if not upc or not start_serial or not end_serial or not save_location:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    try:
        start_serial = int(start_serial)
        end_serial = int(end_serial)
    except ValueError:
        messagebox.showerror("Input Error", "Serial numbers must be integers.")
        return

    if start_serial > end_serial:
        messagebox.showerror("Input Error", "Starting serial number must be less than or equal to the ending serial number.")
        return

    num_serials = end_serial - start_serial + 1
    file_name = f"{upc}_{num_serials}.xlsx"
    file_path = os.path.join(save_location, file_name)

    try:
        # Load the baseline conversion file
        df = pd.read_excel('UPC2EPC.xlsx', sheet_name='Sheet1')
        
        # Create new data for the specified range
        serial_numbers = list(range(start_serial, end_serial + 1))
        epc_values = [generate_epc(upc, sn) for sn in serial_numbers]
        
        # Populate the DataFrame with the new data
        df = pd.DataFrame({
            'UPC': [upc] * num_serials,
            'Serial #': serial_numbers,
            'EPC': epc_values
        })

        # Save the updated file
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"File saved successfully: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Set up the GUI
root = tk.Tk()
root.title("UPC to EPC Conversion")

tk.Label(root, text="UPC:").grid(row=0, column=0, padx=10, pady=10)
upc_entry = tk.Entry(root)
upc_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Starting Serial #:").grid(row=1, column=0, padx=10, pady=10)
serial_start_entry = tk.Entry(root)
serial_start_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="Ending Serial #:").grid(row=2, column=0, padx=10, pady=10)
serial_end_entry = tk.Entry(root)
serial_end_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Save Location:").grid(row=3, column=0, padx=10, pady=10)
save_location_entry = tk.Entry(root, width=40)
save_location_entry.grid(row=3, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=select_save_location).grid(row=3, column=2, padx=10, pady=10)

tk.Button(root, text="Generate File", command=generate_file).grid(row=4, column=0, columnspan=3, pady=20)

root.mainloop()
