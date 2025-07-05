import os
import pandas as pd
from datetime import datetime, timedelta
import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from fpdf import FPDF
from tkinter import messagebox

# === Constants ===
DATA_FILE = "invoice-data.xlsx"
INVOICE_FOLDER = "invoices"
os.makedirs(INVOICE_FOLDER, exist_ok=True)

# === Load or Create Excel File ===
if os.path.exists(DATA_FILE):
    df = pd.read_excel(DATA_FILE)
else:
    df = pd.DataFrame(columns=[
        "Invoice ID", "Client Name", "Service", "Rate (Rs.)", "Quantity",
        "Tax (%)", "Total (Rs.)", "Invoice Date", "Due Date", "Status"
    ])
    df.to_excel(DATA_FILE, index=False)

# === Helper Functions ===
def generate_invoice_number():
    if df.empty:
        return "INV001"
    last_id = df["Invoice ID"].dropna().iloc[-1]
    number = int(last_id.replace("INV", "")) + 1
    return f"INV{number:03d}"

def generate_invoice_pdf(entry):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # Title
    pdf.set_font("Helvetica", 'B', 16)
    pdf.cell(200, 10, text=f"Invoice - {entry['Invoice ID']}", ln=True, align='C')
    pdf.ln(5)

    # Body content
    pdf.set_font("Helvetica", size=12)
    lines = [
        f"Client Name: {entry['Client Name']}",
        f"Service Provided: {entry['Service']}",
        f"Rate: Rs. {entry['Rate (Rs.)']}",
        f"Quantity: {entry['Quantity']}",
        f"Tax: {entry['Tax (%)']}%",
        f"Total Amount: Rs. {entry['Total (Rs.)']}",
        f"Invoice Date: {entry['Invoice Date']}",
        f"Due Date: {entry['Due Date']}",
        f"Status: {entry['Status']}"
    ]

    for line in lines:
        pdf.cell(200, 10, text=line, ln=True)

    pdf.output(f"{INVOICE_FOLDER}/{entry['Invoice ID']}.pdf")

def create_invoice():
    name = name_entry.get()
    service = service_entry.get()
    try:
        rate = float(rate_entry.get())
        quantity = int(quantity_entry.get())
    except ValueError:
        messagebox.showerror("Input Error", "Rate and Quantity must be numbers.")
        return

    if not name or not service:
        messagebox.showerror("Input Error", "Client Name and Service are required.")
        return

    tax = 18
    subtotal = rate * quantity
    total = round(subtotal * (1 + tax / 100), 2)
    invoice_id = generate_invoice_number()
    today = datetime.today()
    invoice_date = today.strftime("%d-%m-%Y")
    due_date = (today + timedelta(days=7)).strftime("%d-%m-%Y")

    new_entry = {
        "Invoice ID": invoice_id,
        "Client Name": name,
        "Service": service,
        "Rate (Rs.)": rate,
        "Quantity": quantity,
        "Tax (%)": tax,
        "Total (Rs.)": total,
        "Invoice Date": invoice_date,
        "Due Date": due_date,
        "Status": "Sent"
    }

    global df
    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
    df.to_excel(DATA_FILE, index=False)

    generate_invoice_pdf(new_entry)
    messagebox.showinfo("Success", f"Invoice {invoice_id} created and saved!")

    # Refresh dropdown
    invoice_ids = df["Invoice ID"].dropna().tolist()
    invoice_combo["values"] = invoice_ids
    invoice_combo.set("")

def open_invoice():
    selected_id = invoice_combo.get()
    if not selected_id:
        messagebox.showwarning("Select ID", "Please select an invoice ID.")
        return

    file_path = os.path.abspath(f"{INVOICE_FOLDER}/{selected_id}.pdf")
    if os.path.exists(file_path):
        try:
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open the PDF:\n{e}")
    else:
        messagebox.showerror("Not Found", f"Invoice {selected_id} not found.")

# === UI Setup ===
app = tb.Window(themename="flatly")
app.title("Invoice Manager Pro")
app.geometry("550x450")

notebook = tb.Notebook(app)
notebook.pack(fill="both", expand=True, padx=10, pady=10)

# === Create Invoice Tab ===
create_tab = tb.Frame(notebook)
notebook.add(create_tab, text="➕ Create Invoice")

tb.Label(create_tab, text="Client Name").grid(row=0, column=0, sticky="w", pady=5)
name_entry = tb.Entry(create_tab)
name_entry.grid(row=0, column=1, pady=5, sticky="ew")

tb.Label(create_tab, text="Service").grid(row=1, column=0, sticky="w", pady=5)
service_entry = tb.Entry(create_tab)
service_entry.grid(row=1, column=1, pady=5, sticky="ew")

tb.Label(create_tab, text="Rate (Rs.)").grid(row=2, column=0, sticky="w", pady=5)
rate_entry = tb.Entry(create_tab)
rate_entry.grid(row=2, column=1, pady=5, sticky="ew")

tb.Label(create_tab, text="Quantity").grid(row=3, column=0, sticky="w", pady=5)
quantity_entry = tb.Entry(create_tab)
quantity_entry.grid(row=3, column=1, pady=5, sticky="ew")

create_tab.columnconfigure(1, weight=1)

tb.Button(create_tab, text="Generate Invoice", bootstyle=SUCCESS, command=create_invoice)\
    .grid(row=5, columnspan=2, pady=15)

# === Open Invoice Tab ===
open_tab = tb.Frame(notebook)
notebook.add(open_tab, text="📂 Open Invoice")

tb.Label(open_tab, text="Select Invoice ID").pack(pady=10)
invoice_ids = df["Invoice ID"].dropna().tolist()
invoice_combo = tb.Combobox(open_tab, values=invoice_ids, font=("Segoe UI", 11))
invoice_combo.pack(pady=5)

tb.Button(open_tab, text="Open PDF", bootstyle=INFO, command=open_invoice).pack(pady=10)

# === Start ===
app.mainloop()
