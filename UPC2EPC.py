import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
import pandas as pd
import openpyxl
import os
import math
import webbrowser
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
from datetime import datetime

# Base path for customer folders
CUSTOMER_BASE_PATH = "Z:/3 Encoding and Printing Files/Customers Encoding Files"

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

def open_roll_tracker(upc, start_serial, end_serial, lpr, total_qty, qty_db, first_epc):
    try:
        roll_tracker_path = os.path.join(os.path.dirname(__file__), 'Roll Tracker v.3.xlsx')
        if not os.path.exists(roll_tracker_path):
            messagebox.showerror("Error", "Roll Tracker v.3.xlsx not found!")
            return
        wb = openpyxl.load_workbook(roll_tracker_path)
        ws = wb['HEX']
        ws['B5'] = upc
        ws['B6'] = start_serial
        ws['B7'] = end_serial
        ws['B8'] = lpr
        ws['B9'] = total_qty
        ws['B10'] = qty_db
        ws['B11'] = first_epc
        wb.save(roll_tracker_path)
        os.startfile(roll_tracker_path)
    except Exception as e:
        messagebox.showerror("Error", str(e))

def verify_epc():
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)
        driver.get("https://verify.gs1.org")
        epc = first_epc_var.get()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "query"))
        ).send_keys(epc + Keys.RETURN)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "gs1-result"))
        )
        driver.quit()
    except Exception as e:
        messagebox.showerror("Error", str(e))

def create_job_ticket_folder(base_path, customer, label_size, po_number, job_ticket_number, upcs):
    current_date = datetime.now().strftime("%y.%m.%d")
    folder_name = f"{current_date} - {po_number} - {job_ticket_number}"
    job_ticket_folder = os.path.join(base_path, customer, label_size, folder_name)

    for upc in upcs:
        upc_folder = os.path.join(job_ticket_folder, upc)
        data_folder = os.path.join(upc_folder, 'Data')
        print_folder = os.path.join(upc_folder, 'Print')

        os.makedirs(data_folder, exist_ok=True)
        os.makedirs(print_folder, exist_ok=True)

    return job_ticket_folder

def generate_file():
    try:
        upc = upc_entry.get()
        start_serial = int(start_serial_entry.get())
        lpr = int(lpr_entry.get())
        total_qty = int(total_qty_entry.get())
        qty_db = int(qty_db_entry.get())
        save_location = save_location_entry.get()
        first_epc = generate_epc(upc, start_serial)
        db_entries = []

        for i in range(total_qty):
            serial = start_serial + i
            epc = generate_epc(upc, serial)
            db_entries.append([upc, serial, epc])

        num_files = math.ceil(total_qty / qty_db)
        progress_bar["maximum"] = num_files
        progress_bar["value"] = 0

        for i in range(num_files):
            start_idx = i * qty_db
            end_idx = min(start_idx + qty_db, total_qty)
            db_chunk = db_entries[start_idx:end_idx]

            df = pd.DataFrame(db_chunk, columns=['UPC', 'Serial', 'EPC'])
            file_name = f"UPC.DB1.1-{len(db_chunk)}K.xlsx"
            file_path = os.path.join(save_location, file_name)
            df.to_excel(file_path, index=False)

            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            ws.column_dimensions['C'].width = 40
            wb.save(file_path)

            progress_bar["value"] += 1
            root.update_idletasks()

        open_roll_tracker(upc, start_serial, start_serial + total_qty - 1, lpr, total_qty, qty_db, first_epc)
        messagebox.showinfo("Success", "File generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def clear_fields():
    upc_entry.delete(0, tk.END)
    start_serial_entry.delete(0, tk.END)
    lpr_entry.delete(0, tk.END)
    total_qty_entry.delete(0, tk.END)
    qty_db_entry.delete(0, tk.END)
    save_location_entry.delete(0, tk.END)
    first_epc_var.set("")

def preview_file():
    messagebox.showinfo("Preview", "Preview functionality not implemented yet.")

def load_sizes(customer):
    customer_path = os.path.join(CUSTOMER_BASE_PATH, customer)
    if not os.path.exists(customer_path):
        messagebox.showerror("Error", f"Customer path '{customer_path}' does not exist.")
        return []

    sizes = [f.name for f in os.scandir(customer_path) if f.is_dir()]
    return sizes

def update_sizes(*args):
    customer = customer_var.get()
    if customer:
        sizes = load_sizes(customer)
        size_dropdown['values'] = sizes
        size_var.set('')  # Clear the current selection

def add_upc_entry_job():
    entry = tk.Entry(job_frame, font=font_style)
    entry.grid(row=len(upc_entries_job) + 6, column=1, sticky="ew", **padding)
    upc_entries_job.append(entry)
    update_job_frame_layout()

def update_job_frame_layout():
    add_upc_entry_job_button.grid(row=len(upc_entries_job) + 6, column=2, padx=5)
    create_job_folder_button.grid(row=len(upc_entries_job) + 7, column=0, columnspan=3, pady=20)

def toggle_multiple_upcs_job():
    if multiple_upc_job_var.get():
        add_upc_entry_job_button.grid(row=5, column=2, padx=5)
        add_upc_entry_job()
    else:
        for entry in upc_entries_job:
            entry.destroy()
        upc_entries_job.clear()
        add_upc_entry_job_button.grid_forget()
        update_job_frame_layout()

def open_job_folder_creator():
    customer = customer_var.get()
    size = size_var.get()
    po_number = po_entry.get()
    job_ticket_number = ticket_entry.get()
    upcs = [upc_entry_job.get()]
    if multiple_upc_job_var.get():
        for entry in upc_entries_job:
            upcs.append(entry.get())
    
    if not customer or not size or not po_number or not job_ticket_number or not all(upcs):
        messagebox.showerror("Error", "Please fill all required fields.")
        return
    
    job_ticket_folder = create_job_ticket_folder(CUSTOMER_BASE_PATH, customer, size, po_number, job_ticket_number, upcs)
    messagebox.showinfo("Success", f"Job folder created at: {job_ticket_folder}")

def load_customers():
    if not os.path.exists(CUSTOMER_BASE_PATH):
        messagebox.showerror("Error", f"Base path '{CUSTOMER_BASE_PATH}' does not exist.")
        return []

    customers = [f.name for f in os.scandir(CUSTOMER_BASE_PATH) if f.is_dir()]
    return customers

root = tk.Tk()
root.title("Database Generator and Job Folder Creator")

font_style = ("Helvetica", 12)
dropdown_font_style = ("Helvetica", 14)
padding = {'padx': 10, 'pady': 5}

style = ttk.Style()
style.configure("TCombobox", font=dropdown_font_style)

# Lists to hold dynamically created UPC entry fields
upc_entries_job = []

# Left Frame for Database Generation
db_frame = tk.Frame(root)
db_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

tk.Label(db_frame, text="Database Generator", font=("Helvetica", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

tk.Label(db_frame, text="UPC:", font=font_style).grid(row=1, column=0, sticky="e", **padding)
upc_entry = tk.Entry(db_frame, font=font_style)
upc_entry.grid(row=1, column=1, sticky="ew", **padding)

tk.Label(db_frame, text="Start Serial:", font=font_style).grid(row=2, column=0, sticky="e", **padding)
start_serial_entry = tk.Entry(db_frame, font=font_style)
start_serial_entry.grid(row=2, column=1, sticky="ew", **padding)

tk.Label(db_frame, text="Labels per Roll:", font=font_style).grid(row=3, column=0, sticky="e", **padding)
lpr_entry = tk.Entry(db_frame, font=font_style)
lpr_entry.grid(row=3, column=1, sticky="ew", **padding)

tk.Label(db_frame, text="Total Quantity:", font=font_style).grid(row=4, column=0, sticky="e", **padding)
total_qty_entry = tk.Entry(db_frame, font=font_style)
total_qty_entry.grid(row=4, column=1, sticky="ew", **padding)

tk.Label(db_frame, text="Qty/DB:", font=font_style).grid(row=5, column=0, sticky="e", **padding)
qty_db_entry = tk.Entry(db_frame, font=font_style)
qty_db_entry.grid(row=5, column=1, sticky="ew", **padding)

tk.Label(db_frame, text="Save Location:", font=font_style).grid(row=6, column=0, sticky="e", **padding)
save_location_entry = tk.Entry(db_frame, font=font_style)
save_location_entry.grid(row=6, column=1, sticky="ew", **padding)
tk.Button(db_frame, text="Browse...", command=select_save_location, font=font_style, bg="#004B87", fg="white").grid(row=6, column=2, **padding)

button_frame = tk.Frame(db_frame)
button_frame.grid(row=7, column=0, columnspan=3, pady=20)

tk.Button(button_frame, text="Generate File", command=generate_file, font=font_style, bg="#4CAF50", fg="white").grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Clear", command=clear_fields, font=font_style, bg="#E60000", fg="white").grid(row=0, column=1, padx=10)
tk.Button(button_frame, text="Preview", command=preview_file, font=font_style, bg="#FFC107", fg="black").grid(row=0, column=2, padx=10)
tk.Button(button_frame, text="Verify", command=verify_epc, font=font_style, bg="#2196F3", fg="white").grid(row=0, column=3, padx=10)

progress_bar = ttk.Progressbar(db_frame, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=8, column=0, columnspan=3, pady=10, sticky="ew")

footer_frame = tk.Frame(db_frame, bg="#004B87")
footer_frame.grid(row=9, column=0, columnspan=3, sticky="ew")
tk.Label(footer_frame, text="Starport Technologies - Converting RFID into the Future", font=("Helvetica", 10), bg="#004B87", fg="white").pack(pady=10)

# Right Frame for Job Folder Creation
job_frame = tk.Frame(root)
job_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

tk.Label(job_frame, text="Job Folder Creator", font=("Helvetica", 16, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

tk.Label(job_frame, text="Customer:", font=font_style).grid(row=1, column=0, sticky="e", **padding)
customer_var = tk.StringVar()
customer_dropdown = ttk.Combobox(job_frame, textvariable=customer_var, font=font_style, height=20)
customer_dropdown['values'] = load_customers()
customer_dropdown.grid(row=1, column=1, sticky="ew", **padding)
customer_dropdown.configure(font=dropdown_font_style)

tk.Label(job_frame, text="Size:", font=font_style).grid(row=2, column=0, sticky="e", **padding)
size_var = tk.StringVar()
size_dropdown = ttk.Combobox(job_frame, textvariable=size_var, font=font_style, height=20)
size_dropdown.grid(row=2, column=1, sticky="ew", **padding)
size_dropdown.configure(font=dropdown_font_style)

tk.Label(job_frame, text="PO Number:", font=font_style).grid(row=3, column=0, sticky="e", **padding)
po_entry = tk.Entry(job_frame, font=font_style)
po_entry.grid(row=3, column=1, sticky="ew", **padding)

tk.Label(job_frame, text="Ticket Number:", font=font_style).grid(row=4, column=0, sticky="e", **padding)
ticket_entry = tk.Entry(job_frame, font=font_style)
ticket_entry.grid(row=4, column=1, sticky="ew", **padding)

tk.Label(job_frame, text="UPC:", font=font_style).grid(row=5, column=0, sticky="e", **padding)
upc_entry_job = tk.Entry(job_frame, font=font_style)
upc_entry_job.grid(row=5, column=1, sticky="ew", **padding)

multiple_upc_job_var = tk.IntVar()
multiple_upc_job_check = tk.Checkbutton(job_frame, text="Multiple UPC's?", font=font_style, variable=multiple_upc_job_var, command=toggle_multiple_upcs_job)
multiple_upc_job_check.grid(row=5, column=2, sticky="w", **padding)

add_upc_entry_job_button = tk.Button(job_frame, text="Add UPC", command=add_upc_entry_job, font=font_style, bg="#004B87", fg="white")

customer_var.trace('w', update_sizes)  # Update sizes when customer changes

create_job_folder_button = tk.Button(job_frame, text="Create Job Folder", command=open_job_folder_creator, font=font_style, bg="#4CAF50", fg="white")
create_job_folder_button.grid(row=6, column=0, columnspan=3, pady=20)

root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
root.mainloop()
