import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import os
import math

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
    gs1_company_prefix = "0" + upc[:6]
    item_reference_number = upc[6:11]
    gtin14 = "0" + gs1_company_prefix + item_reference_number
    header = "00110000"
    filter_value = "001"
    partition = "101"
    gs1_binary = dec_to_bin(gs1_company_prefix, 24)
    item_reference_binary = dec_to_bin(item_reference_number, 20)
    serial_binary = dec_to_bin(serial_number, 38)
    epc_binary = header + filter_value + partition + gs1_binary + item_reference_binary + serial_binary
    epc_hex = bin_to_hex(epc_binary)
    return epc_hex

def open_roll_tracker(upc, start_serial, end_serial, lpr, total_qty, qty_db):
    try:
        roll_tracker_path = os.path.join("C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor", 'Roll Tracker v.3.xlsx')
        if not os.path.exists(roll_tracker_path):
            messagebox.showerror("File Error", f"Roll Tracker file not found at: {roll_tracker_path}")
            return
        wb = openpyxl.load_workbook(roll_tracker_path)
        input_sheet = wb['Input']
        input_sheet['D3'] = lpr
        input_sheet['D4'] = total_qty
        input_sheet['D5'] = start_serial
        input_sheet['D8'] = end_serial
        input_sheet['D10'] = qty_db
        temp_path = os.path.join(os.path.dirname(__file__), 'temp_Roll_Tracker.xlsx')
        wb.save(temp_path)
        os.startfile(temp_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def generate_file():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()
    lpr = lpr_entry.get().strip()
    total_qty = total_qty_entry.get().strip()
    qty_db = qty_db_entry.get().strip()
    save_location = save_location_entry.get().strip()

    if not upc or not start_serial or not lpr or not total_qty or not qty_db or not save_location:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    try:
        start_serial = int(start_serial)
        lpr = int(lpr)
        total_qty = int(total_qty)
        qty_db = int(qty_db)
    except ValueError:
        messagebox.showerror("Input Error", "Serial numbers and quantities must be integers.")
        return

    end_serial = start_serial + total_qty - 1

    num_serials = end_serial - start_serial + 1
    num_dbs = math.ceil(num_serials / qty_db)
    
    try:
        for db_index in range(num_dbs):
            chunk_start = start_serial + db_index * qty_db
            chunk_end = min(chunk_start + qty_db - 1, end_serial)
            chunk_serial_numbers = list(range(chunk_start, chunk_end + 1))
            epc_values = [generate_epc(upc, sn) for sn in chunk_serial_numbers]

            df = pd.DataFrame({
                'UPC': [upc] * len(chunk_serial_numbers),
                'Serial #': chunk_serial_numbers,
                'EPC': epc_values
            })

            start_range = (chunk_start) // 1000
            end_range = (chunk_end + 1) // 1000
            file_name = f"{upc}.DB{db_index + 1}.{start_range}K-{end_range}K.xlsx"
            file_path = os.path.join(save_location, file_name)
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']
                worksheet.column_dimensions['C'].width = 40

        open_roll_tracker(upc, start_serial, end_serial, lpr, total_qty, qty_db)
        messagebox.showinfo("Success", f"Files saved successfully in: {save_location}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

root = tk.Tk()
root.title("UPC to EPC Conversion")
icon_path = "C:\\Users\\Jason\\OneDrive\\Documents\\UPC2EPC Convertor\\download.png"
root.iconphoto(False, tk.PhotoImage(file=icon_path))
font_style = ("Helvetica", 12)
padding = {'padx': 10, 'pady': 10}

tk.Label(root, text="UPC:", font=font_style).grid(row=0, column=0, **padding)
upc_entry = tk.Entry(root, font=font_style)
upc_entry.grid(row=0, column=1, **padding)

tk.Label(root, text="Starting Serial #:", font=font_style).grid(row=1, column=0, **padding)
serial_start_entry = tk.Entry(root, font=font_style)
serial_start_entry.grid(row=1, column=1, **padding)

tk.Label(root, text="Labels per Roll (LPR):", font=font_style).grid(row=2, column=0, **padding)
lpr_entry = tk.Entry(root, font=font_style)
lpr_entry.grid(row=2, column=1, **padding)

tk.Label(root, text="Total Quantity:", font=font_style).grid(row=3, column=0, **padding)
total_qty_entry = tk.Entry(root, font=font_style)
total_qty_entry.grid(row=3, column=1, **padding)

tk.Label(root, text="Qty/DB:", font=font_style).grid(row=4, column=0, **padding)
qty_db_entry = tk.Entry(root, font=font_style)
qty_db_entry.grid(row=4, column=1, **padding)

tk.Label(root, text="Save Location:", font=font_style).grid(row=5, column=0, **padding)
save_location_entry = tk.Entry(root, font=font_style, width=40)
save_location_entry.grid(row=5, column=1, **padding)
tk.Button(root, text="Browse...", command=select_save_location, font=font_style).grid(row=5, column=2, **padding)

tk.Button(root, text="Generate File", command=generate_file, font=font_style, bg="#4CAF50", fg="white").grid(row=6, column=0, columnspan=3, pady=20)

root.mainloop()
