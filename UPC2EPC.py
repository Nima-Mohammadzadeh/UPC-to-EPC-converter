import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
import os
import math
import webbrowser
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

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
        roll_tracker_path = os.path.join(os.path.dirname(__file__), 'Roll Tracker v.3.xlsx')
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
        hex_sheet = wb['HEX']
        wb.active = wb.index(hex_sheet)
        temp_path = os.path.join(os.path.dirname(__file__), 'temp_Roll_Tracker.xlsx')
        wb.save(temp_path)
        os.startfile(temp_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def validate_upc(upc):
    if (len(upc) != 12) or (not upc.isdigit()):
        messagebox.showerror("Input Error", "UPC must be exactly 12 digits.")
        return False
    return True

def generate_file():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()
    total_qty = total_qty_entry.get().strip()
    lpr = lpr_entry.get().strip()
    qty_db = qty_db_entry.get().strip()
    save_location = save_location_entry.get().strip()

    if not upc or not start_serial or not lpr or not total_qty or not qty_db or not save_location:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    if not validate_upc(upc):
        return

    try:
        start_serial = int(start_serial)
        total_qty = int(total_qty)
        lpr = int(lpr)
        qty_db = int(qty_db)
    except ValueError:
        messagebox.showerror("Input Error", "Serial numbers and quantities must be integers.")
        return

    end_serial = start_serial + total_qty - 1
    num_serials = end_serial - start_serial + 1
    num_dbs = math.ceil(num_serials / qty_db)
    
    try:
        progress_bar['maximum'] = num_dbs
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

            start_range = (chunk_start // 1000) + 1 if chunk_start % 1000 == 0 else (chunk_start // 1000)
            end_range = ((chunk_end + 1) // 1000)
            file_name = f"{upc}.DB{db_index + 1}.{start_range}K-{end_range}K.xlsx"
            file_path = os.path.join(save_location, file_name)
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']
                worksheet.column_dimensions['C'].width = 40

            progress_bar['value'] = db_index + 1
            root.update_idletasks()

        open_roll_tracker(upc, start_serial, end_serial, lpr, total_qty, qty_db)
        messagebox.showinfo("Success", f"Files saved successfully in: {save_location}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def preview_file():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()
    total_qty = total_qty_entry.get().strip()
    lpr = lpr_entry.get().strip()
    qty_db = qty_db_entry.get().strip()

    if not upc or not start_serial or not lpr or not total_qty or not qty_db:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    if not validate_upc(upc):
        return

    try:
        start_serial = int(start_serial)
        total_qty = int(total_qty)
        lpr = int(lpr)
        qty_db = int(qty_db)
    except ValueError:
        messagebox.showerror("Input Error", "Serial numbers and quantities must be integers.")
        return

    end_serial = start_serial + total_qty - 1
    chunk_serial_numbers = list(range(start_serial, min(start_serial + 10, end_serial + 1)))
    epc_values = [generate_epc(upc, sn) for sn in chunk_serial_numbers]

    df = pd.DataFrame({
        'UPC': [upc] * len(chunk_serial_numbers),
        'Serial #': chunk_serial_numbers,
        'EPC': epc_values
    })

    preview_window = tk.Toplevel(root)
    preview_window.title("Preview Data")
    preview_window.geometry("600x400")

    preview_table = ttk.Treeview(preview_window, columns=("UPC", "Serial #", "EPC"), show="headings")
    preview_table.heading("UPC", text="UPC")
    preview_table.heading("Serial #", text="Serial #")
    preview_table.heading("EPC", text="EPC")

    for index, row in df.iterrows():
        preview_table.insert("", "end", values=(row["UPC"], row["Serial #"], row["EPC"]))

    preview_table.pack(expand=True, fill="both")

def verify_epc():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()

    if not upc or not start_serial:
        messagebox.showerror("Input Error", "UPC and Starting Serial # are required for verification.")
        return

    if not validate_upc(upc):
        return

    try:
        start_serial = int(start_serial)
    except ValueError:
        messagebox.showerror("Input Error", "Serial numbers must be integers.")
        return

    epc = generate_epc(upc, start_serial)
    epc_url = "https://www.gs1.org/services/epc-encoderdecoder"

    try:
        # Initialize the WebDriver (assuming ChromeDriver is in your PATH)
        driver = webdriver.Chrome()
        driver.get(epc_url)

        # Wait for the page to load
        driver.implicitly_wait(10)

        # Locate the input field by its identifier and fill in the EPC value
        epc_input_field = driver.find_element(By.XPATH, '//*[@id="epcContainer"]/table/tbody/tr/td/div/div[5]/input')
        epc_input_field.send_keys(epc)

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while verifying the EPC: {str(e)}")
    finally:
        # Optionally, close the WebDriver after some time or leave it open for manual interaction
        # driver.quit()
        pass

def clear_fields():
    upc_entry.delete(0, tk.END)
    serial_start_entry.delete(0, tk.END)
    lpr_entry.delete(0, tk.END)
    total_qty_entry.delete(0, tk.END)
    qty_db_entry.delete(0, tk.END)
    save_location_entry.delete(0, tk.END)
    progress_bar['value'] = 0

root = tk.Tk()
root.title("Database Generator")

# Set icon path relative to script location
icon_path = resource_path('download.png')
root.iconphoto(False, tk.PhotoImage(file=icon_path))

font_style = ("Helvetica", 12)
padding = {'padx': 10, 'pady': 10}

header_frame = tk.Frame(root, bg="#004B87")
header_frame.grid(row=0, column=0, columnspan=3, sticky="ew")

tk.Label(header_frame, text="Database Generator", font=("Helvetica", 16, "bold"), bg="#004B87", fg="white").pack(pady=10)

input_frame = tk.Frame(root)
input_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

tk.Label(input_frame, text="UPC:", font=font_style).grid(row=0, column=0, sticky="e", **padding)
upc_entry = tk.Entry(input_frame, font=font_style)
upc_entry.grid(row=0, column=1, sticky="ew", **padding)
input_frame.columnconfigure(1, weight=1)

tk.Label(input_frame, text="Starting Serial #:", font=font_style).grid(row=1, column=0, sticky="e", **padding)
serial_start_entry = tk.Entry(input_frame, font=font_style)
serial_start_entry.grid(row=1, column=1, sticky="ew", **padding)

tk.Label(input_frame, text="Labels per Roll (LPR):", font=font_style).grid(row=2, column=0, sticky="e", **padding)
lpr_entry = tk.Entry(input_frame, font=font_style)
lpr_entry.grid(row=2, column=1, sticky="ew", **padding)

tk.Label(input_frame, text="Total Quantity:", font=font_style).grid(row=3, column=0, sticky="e", **padding)
total_qty_entry = tk.Entry(input_frame, font=font_style)
total_qty_entry.grid(row=3, column=1, sticky="ew", **padding)

tk.Label(input_frame, text="Qty/DB:", font=font_style).grid(row=4, column=0, sticky="e", **padding)
qty_db_entry = tk.Entry(input_frame, font=font_style)
qty_db_entry.grid(row=4, column=1, sticky="ew", **padding)

tk.Label(input_frame, text="Save Location:", font=font_style).grid(row=5, column=0, sticky="e", **padding)
save_location_entry = tk.Entry(input_frame, font=font_style)
save_location_entry.grid(row=5, column=1, sticky="ew", **padding)
tk.Button(input_frame, text="Browse...", command=select_save_location, font=font_style, bg="#004B87", fg="white").grid(row=5, column=2, **padding)

button_frame = tk.Frame(root)
button_frame.grid(row=6, column=0, columnspan=3, pady=20)

tk.Button(button_frame, text="Generate File", command=generate_file, font=font_style, bg="#4CAF50", fg="white").grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Clear", command=clear_fields, font=font_style, bg="#E60000", fg="white").grid(row=0, column=1, padx=10)
tk.Button(button_frame, text="Preview", command=preview_file, font=font_style, bg="#FFC107", fg="black").grid(row=0, column=2, padx=10)
tk.Button(button_frame, text="Verify", command=verify_epc, font=font_style, bg="#2196F3", fg="white").grid(row=0, column=3, padx=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=7, column=0, columnspan=3, pady=10, sticky="ew")

footer_frame = tk.Frame(root, bg="#004B87")
footer_frame.grid(row=8, column=0, columnspan=3, sticky="ew")
tk.Label(footer_frame, text="Starport Technologies - Converting RFID into the Future", font=("Helvetica", 10), bg="#004B87", fg="white").pack(pady=10)

root.columnconfigure(1, weight=1)
root.mainloop()
