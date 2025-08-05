import tkinter as tk
from tkinter import messagebox, ttk
import csv
import os
import openpyxl

csv_file = 'contacts.csv'
excel_file = 'contacts.xlsx'

#CSV headers
if not os.path.exists(csv_file):
    with open(csv_file, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["Name", "Email", "Phone", "Address"])

# New contact
def save_contact():
    name = name_entry.get().strip()
    email = email_entry.get().strip()
    phone = phone_entry.get().strip()
    address = address_entry.get("1.0", tk.END).strip()

    if not (name and email and phone and address):
        messagebox.showwarning("Validation Error", "All fields are required.")
        return

    with open(csv_file, mode='a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([name, email, phone, address])

    messagebox.showinfo("Success", "Contact saved successfully!")
    clear_fields()

# Clear form inputs
def clear_fields():
    name_entry.delete(0, tk.END)
    email_entry.delete(0, tk.END)
    phone_entry.delete(0, tk.END)
    address_entry.delete("1.0", tk.END)

# Export CSV to Excel
def export_to_excel():
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Contacts"

        with open(csv_file, mode='r') as file:
            reader = csv.reader(file)
            for row in reader:
                ws.append(row)

        wb.save(excel_file)
        messagebox.showinfo("Exported", f"Contacts exported to {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# View contacts in new window
def view_contacts():
    view_win = tk.Toplevel(window)
    view_win.title("Saved Contacts")
    view_win.geometry("600x300")

    tree = ttk.Treeview(view_win, columns=("Name", "Email", "Phone", "Address"), show="headings")
    tree.heading("Name", text="Name")
    tree.heading("Email", text="Email")
    tree.heading("Phone", text="Phone")
    tree.heading("Address", text="Address")
    tree.pack(fill=tk.BOTH, expand=True)

    with open(csv_file, mode='r') as file:
        reader = csv.reader(file)
        next(reader)  # skip header
        for row in reader:
            tree.insert("", tk.END, values=row)

# GUI Layout
window = tk.Tk()
window.title("Contact Form")
window.geometry("450x500")
window.resizable(False, False)

tk.Label(window, text="Name:").pack(pady=(10, 0))
name_entry = tk.Entry(window, width=40)
name_entry.pack()

tk.Label(window, text="Email:").pack(pady=(10, 0))
email_entry = tk.Entry(window, width=40)
email_entry.pack()

tk.Label(window, text="Phone:").pack(pady=(10, 0))
phone_entry = tk.Entry(window, width=40)
phone_entry.pack()

tk.Label(window, text="Address:").pack(pady=(10, 0))
address_entry = tk.Text(window, width=30, height=4)
address_entry.pack()

tk.Button(window, text="Submit", command=save_contact, bg="#4CAF50", fg="white").pack(pady=10)
tk.Button(window, text="View Contacts", command=view_contacts, bg="#2196F3", fg="white").pack(pady=5)
tk.Button(window, text="Export to Excel", command=export_to_excel, bg="#FF9800", fg="white").pack(pady=5)

window.mainloop()