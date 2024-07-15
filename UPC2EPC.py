import shutil
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openpyxl
import math
import webbrowser
from selenium import webdriver
from selenium.webdriver.common.by import By
import sys

# Define global paths
base_path = r'Z:\3 Encoding and Printing Files\Customers Encoding Files'
template_base_path = r'Z:\3 Encoding and Printing Files\Templates'

# Global variable to store the path of the created job folder
job_data_folder_path = None

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def select_save_location():
    global job_data_folder_path
    folder_selected = job_data_folder_path if job_data_folder_path else filedialog.askdirectory()
    if folder_selected:
        save_location_entry.delete(0, tk.END)
        save_location_entry.insert(0, folder_selected)

def select_template():
    customer = customer_var.get().strip()
    label_size = label_size_var.get().strip()
    initial_dir = os.path.join(template_base_path, customer, label_size)
    
    if not os.path.exists(initial_dir):
        messagebox.showerror("Directory Error", f"Directory does not exist: {initial_dir}")
        return
    
    file_selected = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("BarTender Template Files", "*.btw")])
    if file_selected:
        template_entry.delete(0, tk.END)
        template_entry.insert(0, file_selected)

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

def calculate_total_quantity():
    try:
        total_qty = int(total_qty_entry.get())
        if var_2_percent.get():
            total_qty += total_qty * 0.02
        if var_7_percent.get():
            total_qty += total_qty * 0.07
        return int(total_qty)
    except ValueError:
        return 0

def generate_file():
    upc = upc_entry.get().strip()
    start_serial = serial_start_entry.get().strip()
    lpr = lpr_entry.get().strip()
    qty_db = qty_db_entry.get().strip()
    save_location = save_location_entry.get().strip()
    total_qty = calculate_total_quantity()

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

def on_checkbox_change():
    total_qty = calculate_total_quantity()
    total_quantity_label.config(text=f"Updated Total Quantity: {total_qty}")

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
    customer_menu.set('')
    label_size_menu.set('')
    ticket_number_entry.delete(0, tk.END)
    po_number_entry.delete(0, tk.END)
    upc_entry_job.delete(0, tk.END)

def populate_customer_dropdown():
    customer_dir = base_path
    if os.path.exists(customer_dir):
        customers = os.listdir(customer_dir)
        customers = [customer for customer in customers if os.path.isdir(os.path.join(customer_dir, customer))]
        customer_var.set(customers[0] if customers else '')
        customer_menu['values'] = customers
        update_label_size_dropdown(customers[0])
    else:
        messagebox.showerror("Directory Error", f"Customer directory not found: {customer_dir}")

def update_label_size_dropdown(customer):
    label_size_dir = os.path.join(base_path, customer)
    if os.path.exists(label_size_dir):
        label_sizes = os.listdir(label_size_dir)
        label_sizes = [size for size in label_sizes if os.path.isdir(os.path.join(label_size_dir, size))]
        label_size_var.set(label_sizes[0] if label_sizes else '')
        label_size_menu['values'] = label_sizes
    else:
        messagebox.showerror("Directory Error", f"Label size directory not found for customer: {customer}")

def on_customer_select(event):
    selected_customer = customer_var.get()
    update_label_size_dropdown(selected_customer)

def create_job_folder():
    global job_data_folder_path
    customer = customer_var.get().strip()
    label_size = label_size_var.get().strip()
    ticket_number = ticket_number_entry.get().strip()
    po_number = po_number_entry.get().strip()
    upc = upc_entry_job.get().strip()

    if not customer or not label_size or not ticket_number or not po_number or not upc:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    # Construct the template path
    template_path = os.path.join(template_base_path, customer, label_size, f"Template {label_size}.btw")

    if not os.path.exists(template_path):
        messagebox.showerror("Template Error", f"Template not found at {template_path}")
        return

    today_date = datetime.datetime.now().strftime("%m.%d.%y")
    folder_name = f"{today_date} - {po_number} - {ticket_number}"
    job_folder_path = os.path.join(base_path, customer, label_size, folder_name)
    upc_folder_path = os.path.join(job_folder_path, upc)
    job_data_folder_path = os.path.join(upc_folder_path, "data")

    try:
        os.makedirs(upc_folder_path, exist_ok=True)
        os.makedirs(os.path.join(upc_folder_path, "print"), exist_ok=True)
        os.makedirs(job_data_folder_path, exist_ok=True)
        
        # Copy the constructed template path to the print folder and rename it to the UPC
        shutil.copy(template_path, os.path.join(upc_folder_path, "print", f"{upc}.btw"))
        print(f"Template copied to {os.path.join(upc_folder_path, 'print', f'{upc}.btw')}")
        messagebox.showinfo("Success", f"Folder created successfully at: {upc_folder_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def create_database_generator_tab(tab):
    header_frame = tk.Frame(tab, bg="#004B87")
    header_frame.grid(row=0, column=0, columnspan=3, sticky="ew")

    tk.Label(header_frame, text="Database Generator", font=("Helvetica", 16, "bold"), bg="#004B87", fg="white").pack(pady=10)

    input_frame = tk.Frame(tab)
    input_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    tk.Label(input_frame, text="UPC:", font=("Helvetica", 12)).grid(row=0, column=0, sticky="e", padx=10, pady=10)
    global upc_entry
    upc_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    upc_entry.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
    input_frame.columnconfigure(1, weight=1)

    tk.Label(input_frame, text="Starting Serial #:", font=("Helvetica", 12)).grid(row=1, column=0, sticky="e", padx=10, pady=10)
    global serial_start_entry
    serial_start_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    serial_start_entry.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="Labels per Roll (LPR):", font=("Helvetica", 12)).grid(row=2, column=0, sticky="e", padx=10, pady=10)
    global lpr_entry
    lpr_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    lpr_entry.grid(row=2, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="Total Quantity:", font=("Helvetica", 12)).grid(row=3, column=0, sticky="e", padx=10, pady=10)
    global total_qty_entry
    total_qty_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    total_qty_entry.grid(row=3, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="Qty/DB:", font=("Helvetica", 12)).grid(row=4, column=0, sticky="e", padx=10, pady=10)
    global qty_db_entry
    qty_db_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    qty_db_entry.grid(row=4, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="Save Location:", font=("Helvetica", 12)).grid(row=5, column=0, sticky="e", padx=10, pady=10)
    global save_location_entry
    save_location_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    save_location_entry.grid(row=5, column=1, sticky="ew", padx=10, pady=10)
    tk.Button(input_frame, text="Browse...", command=select_save_location, font=("Helvetica", 12), bg="#004B87", fg="white").grid(row=5, column=2, padx=10, pady=10)

    # Checkboxes for 2% and 7%
    global var_2_percent, var_7_percent
    var_2_percent = tk.BooleanVar()
    var_7_percent = tk.BooleanVar()
    checkbox_2_percent = tk.Checkbutton(input_frame, text="2%", variable=var_2_percent, command=on_checkbox_change)
    checkbox_7_percent = tk.Checkbutton(input_frame, text="7%", variable=var_7_percent, command=on_checkbox_change)
    checkbox_2_percent.grid(row=6, column=0, padx=10, pady=5)
    checkbox_7_percent.grid(row=6, column=1, padx=10, pady=5)

    # Label to show updated total quantity
    global total_quantity_label
    total_quantity_label = tk.Label(input_frame, text="Updated Total Quantity: 0", font=("Helvetica", 12))
    total_quantity_label.grid(row=7, column=0, columnspan=2, padx=10, pady=5)

    button_frame = tk.Frame(tab)
    button_frame.grid(row=8, column=0, columnspan=3, pady=20)

    tk.Button(button_frame, text="Generate File", command=generate_file, font=("Helvetica", 12), bg="#4CAF50", fg="white").grid(row=0, column=0, padx=10)
    tk.Button(button_frame, text="Clear", command=clear_fields, font=("Helvetica", 12), bg="#E60000", fg="white").grid(row=0, column=1, padx=10)
    tk.Button(button_frame, text="Preview", command=preview_file, font=("Helvetica", 12), bg="#FFC107", fg="black").grid(row=0, column=2, padx=10)
    tk.Button(button_frame, text="Verify", command=verify_epc, font=("Helvetica", 12), bg="#2196F3", fg="white").grid(row=0, column=3, padx=10)

    global progress_bar
    progress_bar = ttk.Progressbar(tab, orient="horizontal", length=400, mode="determinate")
    progress_bar.grid(row=9, column=0, columnspan=3, pady=10, sticky="ew")

    footer_frame = tk.Frame(tab, bg="#004B87")
    footer_frame.grid(row=10, column=0, columnspan=3, sticky="ew")
    tk.Label(footer_frame, text="Starport Technologies - Converting RFID into the Future", font=("Helvetica", 10), bg="#004B87", fg="white").pack(pady=10)

def create_job_creator_tab(tab):
    header_frame = tk.Frame(tab, bg="#004B87")
    header_frame.grid(row=0, column=0, columnspan=3, sticky="ew")

    tk.Label(header_frame, text="Job Creator", font=("Helvetica", 16, "bold"), bg="#004B87", fg="white").pack(pady=10)

    input_frame = tk.Frame(tab)
    input_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

    tk.Label(input_frame, text="Customer Name:", font=("Helvetica", 12)).grid(row=0, column=0, sticky="e", padx=10, pady=10)
    global customer_var
    customer_var = tk.StringVar()
    global customer_menu
    customer_menu = ttk.Combobox(input_frame, textvariable=customer_var, font=("Helvetica", 12), height=15)
    customer_menu.grid(row=0, column=1, sticky="ew", padx=10, pady=10)
    customer_menu.bind("<<ComboboxSelected>>", on_customer_select)

    tk.Label(input_frame, text="Label Size:", font=("Helvetica", 12)).grid(row=1, column=0, sticky="e", padx=10, pady=10)
    global label_size_var
    label_size_var = tk.StringVar()
    global label_size_menu
    label_size_menu = ttk.Combobox(input_frame, textvariable=label_size_var, font=("Helvetica", 12), height=15)
    label_size_menu.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="Ticket Number:", font=("Helvetica", 12)).grid(row=2, column=0, sticky="e", padx=10, pady=10)
    global ticket_number_entry
    ticket_number_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    ticket_number_entry.grid(row=2, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="PO Number:", font=("Helvetica", 12)).grid(row=3, column=0, sticky="e", padx=10, pady=10)
    global po_number_entry
    po_number_entry = tk.Entry(input_frame, font=("Helvetica", 12))
    po_number_entry.grid(row=3, column=1, sticky="ew", padx=10, pady=10)

    tk.Label(input_frame, text="UPC:", font=("Helvetica", 12)).grid(row=4, column=0, sticky="e", padx=10, pady=10)
    global upc_entry_job
    upc_entry_job = tk.Entry(input_frame, font=("Helvetica", 12))
    upc_entry_job.grid(row=4, column=1, sticky="ew", padx=10, pady=10)

    button_frame = tk.Frame(tab)
    button_frame.grid(row=6, column=0, columnspan=3, pady=20)

    tk.Button(button_frame, text="Create Job Folder", command=create_job_folder, font=("Helvetica", 12), bg="#4CAF50", fg="white").grid(row=0, column=0, padx=10)
    tk.Button(button_frame, text="Clear", command=clear_fields, font=("Helvetica", 12), bg="#E60000", fg="white").grid(row=0, column=1, padx=10)

    footer_frame = tk.Frame(tab, bg="#004B87")
    footer_frame.grid(row=7, column=0, columnspan=3, sticky="ew")
    tk.Label(footer_frame, text="Starport Technologies - Converting RFID into the Future", font=("Helvetica", 10), bg="#004B87", fg="white").pack(pady=10)

    tab.rowconfigure(1, weight=1)
    tab.columnconfigure(0, weight=1)

    populate_customer_dropdown()


# Main function to initialize the GUI
def initialize_gui():
    notebook = ttk.Notebook(root)

    job_creator_tab = ttk.Frame(notebook)
    database_generator_tab = ttk.Frame(notebook)

    notebook.add(job_creator_tab, text="Job Creator")
    notebook.add(database_generator_tab, text="Database Generator")

    notebook.pack(expand=True, fill='both')

    create_job_creator_tab(job_creator_tab)
    create_database_generator_tab(database_generator_tab)

root = tk.Tk()
root.title("Job Creator and Database Generator")

# Set icon path relative to script location
icon_path = resource_path('download.png')
root.iconphoto(False, tk.PhotoImage(file=icon_path))

root.resizable(False, False)

initialize_gui()    

root.mainloop()
