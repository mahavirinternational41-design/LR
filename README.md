import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

# Function to save data to Excel
def save_data():
    data = {
        "Party Name": entry_party.get(),
        "Invoice No": entry_invoice.get(),
        "LR No": entry_lr.get(),
        "Date": entry_date.get(),
        "Stock Received": entry_stock.get(),
        "Issued By": entry_issued.get()
    }
    
    # Check if any field is empty
    if "" in data.values():
        messagebox.showwarning("Error", "All fields are required!")
        return

    df = pd.DataFrame([data])
    file_name = "logistics_data.xlsx"

    # If file exists, append; otherwise, create new
    if os.path.isfile(file_name):
        with pd.ExcelWriter(file_name, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
            existing_df = pd.read_excel(file_name)
            combined_df = pd.concat([existing_df, df], ignore_index=True)
            combined_df.to_excel(file_name, index=False)
    else:
        df.to_excel(file_name, index=False)

    messagebox.showinfo("Success", "Data saved to Excel!")
    # Clear entries
    for entry in entries:
        entry.delete(0, tk.END)

# GUI Setup
root = tk.Tk()
root.title("Logistics Data Entry")
root.geometry("400x400")

labels = ["Party Name", "Invoice No", "LR No", "Date", "Stock Received", "Issued By"]
entries = []

for label_text in labels:
    label = tk.Label(root, text=label_text)
    label.pack(pady=2)
    entry = tk.Entry(root, width=40)
    entry.pack(pady=5)
    entries.append(entry)

# Assigning entries to variables for the save function
entry_party, entry_invoice, entry_lr, entry_date, entry_stock, entry_issued = entries

btn_save = tk.Button(root, text="Save to Excel", command=save_data, bg="green", fg="white")
btn_save.pack(pady=20)

root.mainloop()
