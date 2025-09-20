import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to validate shift dates and find missing ones
def validate_shift_dates(row, shift_col, FG):
    enterprise_id = row['Enterprise id']
    raw_data = row[shift_col]

    if pd.isna(raw_data) or not str(raw_data).strip():
        return "True (empty)", []

    shift_days = [d.strip().lstrip('0') for d in str(raw_data).replace('.', ',').split(',') if d.strip()]
    matching_entries = FG[FG['Email'].str.contains(enterprise_id, case=False, na=False)]
    matching_days = matching_entries['Time Entry Day'].dropna().tolist()

    missing_days = [day for day in shift_days if day not in matching_days]
    is_valid = len(missing_days) == 0

    return is_valid, missing_days

# Function to process the Excel file
def process_file(input_path, output_path):
    try:
        # Load the Excel sheets
        Shift_Data = pd.read_excel(input_path, sheet_name='Shift_Data')
        FG = pd.read_excel(input_path, sheet_name='FG')

        # Normalize Time Entry Date to extract day as string
        FG['Time Entry Day'] = pd.to_datetime(FG['Time Entry Date'], errors='coerce').dt.day.astype('Int64').astype(str)

        # Apply validation
        Shift_Data['Shift B Valid'], Shift_Data['Missing Shift B Dates'] = zip(*Shift_Data.apply(lambda row: validate_shift_dates(row, 'Shift B dates', FG), axis=1))
        Shift_Data['Shift C Valid'], Shift_Data['Missing Shift C Dates'] = zip(*Shift_Data.apply(lambda row: validate_shift_dates(row, 'Shift C dates', FG), axis=1))

        # Calculate total shift days per user with NaN handling
        Shift_Data['Total Shift B Days'] = Shift_Data['Shift B dates'].apply(
            lambda x: 0 if pd.isna(x) else len([d.strip() for d in str(x).replace('.', ',').split(',') if d.strip()])
        )
        Shift_Data['Total Shift C Days'] = Shift_Data['Shift C dates'].apply(
            lambda x: 0 if pd.isna(x) else len([d.strip() for d in str(x).replace('.', ',').split(',') if d.strip()])
        )
        Shift_Data['Total Shift Days'] = Shift_Data['Total Shift B Days'] + Shift_Data['Total Shift C Days']

        # Calculate overall totals
        total_resources = Shift_Data['Enterprise id'].nunique()
        total_shift_b_days = Shift_Data['Total Shift B Days'].sum()
        total_shift_c_days = Shift_Data['Total Shift C Days'].sum()

        # Create a summary DataFrame
        summary_data = {
            'Total number of resources': [total_resources],
            'Total number of Shift B days': [total_shift_b_days],
            'Total number of Shift C days': [total_shift_c_days]
        }
        summary_df = pd.DataFrame(summary_data)

        # Write to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            Shift_Data.to_excel(writer, sheet_name='Shift Data', index=False)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        messagebox.showinfo("Success", f"The output has been written to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# GUI Functions
def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xlsx;*.xls")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, file_path)

def start_processing():
    input_path = input_entry.get()
    output_path = output_entry.get()
    if not input_path or not output_path:
        messagebox.showwarning("Input Required", "Please select both input and output file paths.")
        return
    process_file(input_path, output_path)

# Create the main window
root = tk.Tk()
root.title("Shift_Allowance_Checker_Tool")

# Layout
tk.Label(root, text="Select Input Excel File:").grid(row=0, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_input_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Select Output Excel File:").grid(row=1, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_output_file).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Start Processing", command=start_processing).grid(row=2, column=0, columnspan=3, pady=20)

# Run the GUI
root.mainloop()
