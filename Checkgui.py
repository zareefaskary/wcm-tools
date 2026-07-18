import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
import pandas as pd

class CreateToolTip:
    """
    Create a tooltip for a given widget
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)  # Remove window decorations
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify=tk.LEFT,
            background="#ffffe0", relief=tk.SOLID, borderwidth=1,
            font=("tahoma", "8", "normal")
        )
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

def update_tooltip(entry_widget):
    path = entry_widget.get()
    if path:
        CreateToolTip(entry_widget, path)

def check_plant_codes(master_file, ysam_folder, master_sheet, company_col, plant_col, year, posting_period, output_box):
    output_box.insert(tk.END, f"\nReading master file: {master_file}\n")
    master_df = pd.read_excel(master_file, sheet_name=master_sheet, dtype=str)
    master_df[company_col] = master_df[company_col].str.strip()
    master_df[plant_col] = master_df[plant_col].str.strip()

    grouped = master_df.groupby(company_col)[plant_col].apply(list).to_dict()
    total_companies = len(grouped)

    missing_files = []
    problems_found = False  # <-- track if any issues are found

    for idx, (company_code, expected_plants) in enumerate(grouped.items(), start=1):
        company_str = str(company_code).zfill(4)
        possible_filenames = [
            f"YSAM_{year}0{posting_period}_{company_str}.xlsx",
            f"YSAM_{year}0{posting_period}_{company_str}.XLSX",
            f"YSAM_{year}0{posting_period}_{int(company_code)}.xlsx",
            f"YSAM_{year}0{posting_period}_{int(company_code)}.XLSX"
        ]

        ysam_file = None
        for fname in possible_filenames:
            fpath = Path(ysam_folder) / fname
            if fpath.exists():
                ysam_file = fpath
                break

        if ysam_file is None:
            output_box.insert(tk.END, f"Company {company_code} → FILE NOT FOUND\n")
            missing_files.append(company_code)
            problems_found = True
            continue

        try:
            ysam_df = pd.read_excel(ysam_file, dtype=str)

            # Check Plant column
            if 'Plant' not in ysam_df.columns:
                output_box.insert(tk.END, f"Company {company_code} → Plant column missing in file {ysam_file.name}\n")
                problems_found = True
                continue

            actual_plants = ysam_df['Plant'].str.strip().unique().tolist()
            expected_set = set(expected_plants)
            actual_set = set(actual_plants)

            missing_plants = sorted(expected_set - actual_set)
            extra_plants = sorted(actual_set - expected_set)

            # Check FYr and Posting Period
            fy_col, month_col = 'FYr previous period', 'Month prev. period'
            fy_mismatch = False
            month_mismatch = False

            if fy_col in ysam_df.columns:
                fy_values = ysam_df[fy_col].dropna().unique()
                fy_mismatch = not all(str(v).strip() == year for v in fy_values)
            else:
                fy_mismatch = True

            if month_col in ysam_df.columns:
                month_values = ysam_df[month_col].dropna().unique()
                month_mismatch = not all(str(v).strip().zfill(2) == posting_period.zfill(2) for v in month_values)
            else:
                month_mismatch = True

            # Only output problems
            if missing_plants or extra_plants or fy_mismatch or month_mismatch:
                output_box.insert(tk.END, f"\nCompany {company_code} → File: {ysam_file.name}\n")
                if missing_plants:
                    output_box.insert(tk.END, f"  Missing plants: {missing_plants}\n")
                if extra_plants:
                    output_box.insert(tk.END, f"  Extra plants: {extra_plants}\n")
                if fy_mismatch:
                    output_box.insert(tk.END, f"  FYr mismatch (expected {year})\n")
                if month_mismatch:
                    output_box.insert(tk.END, f"  Posting period mismatch (expected {posting_period})\n")
                problems_found = True

        except Exception as e:
            output_box.insert(tk.END, f"Company {company_code} → ERROR reading file {ysam_file.name}: {e}\n")
            problems_found = True

    if missing_files:
        output_box.insert(tk.END, f"\nCompanies with missing files: {missing_files}\n")

    if not problems_found:  # <-- if no issues found
        output_box.insert(tk.END, "\nAll files are correctly exported and no data mismatch.\n")

    output_box.insert(tk.END, "\n=== Validation Complete ===\n")


# GUI Section
def browse_master():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    master_entry.delete(0, tk.END)
    master_entry.insert(0, path)
    update_tooltip(master_entry)

def browse_ysam():
    path = filedialog.askdirectory()
    ysam_entry.delete(0, tk.END)
    ysam_entry.insert(0, path)
    update_tooltip(ysam_entry)

def run_check():
    master_file = master_entry.get()
    ysam_folder = ysam_entry.get()
    master_sheet = sheet_entry.get()
    company_col = company_col_entry.get()
    plant_col = plant_col_entry.get()
    year = year_entry.get()
    posting_period = posting_period_combobox.get()

    if not all([master_file, ysam_folder, master_sheet, company_col, plant_col, year, posting_period]):
        messagebox.showerror("Missing input", "Please fill in all fields.")
        return

    output_box.delete(1.0, tk.END)
    check_plant_codes(master_file, ysam_folder, master_sheet, company_col, plant_col, year, posting_period, output_box)

def clear_all_fields():
    master_entry.delete(0, tk.END)
    ysam_entry.delete(0, tk.END)
    sheet_entry.delete(0, tk.END)  # reset to default
    company_col_entry.delete(0, tk.END)
    plant_col_entry.delete(0, tk.END)
    year_entry.delete(0, tk.END)
    posting_period_combobox.set("")
    

# Tkinter Window
root = tk.Tk()
root.title("YSAM Plant Code Checker")
root.geometry("800x640")
root.resizable(False, False)

tk.Label(root, text="Master Excel File:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
master_entry = tk.Entry(root, width=60)
master_entry.grid(row=0, sticky="w", column=1, pady=5)
tk.Button(root, text="Browse", command=browse_master).grid(row=0, sticky="w", column=2, padx=5)

tk.Label(root, text="YSAM Folder:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
ysam_entry = tk.Entry(root, width=60)
ysam_entry.grid(row=1, sticky="w", column=1, pady=5)
tk.Button(root, text="Browse", command=browse_ysam).grid(row=1, sticky="w", column=2, padx=5)

tk.Label(root, text="Sheet Name:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
sheet_entry = tk.Entry(root, width=20)
sheet_entry.insert(0, "Plant list_All")
sheet_entry.grid(row=2, column=1, sticky="w", pady=5)

tk.Label(root, text="Company Code Column:").grid(row=3, column=0, sticky="w", padx=10, pady=5)
company_col_entry = tk.Entry(root, width=20)
company_col_entry.insert(0, "Company Code")
company_col_entry.grid(row=3, column=1, sticky="w", pady=5)

tk.Label(root, text="Plant Code Column:").grid(row=4, column=0, sticky="w", padx=10, pady=5)
plant_col_entry = tk.Entry(root, width=20)
plant_col_entry.insert(0, "Plant Code")
plant_col_entry.grid(row=4, column=1, sticky="w", pady=5)

tk.Label(root, text="Year:").grid(row=5, column=0, sticky="w", padx=10, pady=5)
year_entry = tk.Entry(root, width=10)
year_entry.grid(row=5, column=1, sticky="w", pady=5)

tk.Label(root, text="Posting Period:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
posting_period_combobox = ttk.Combobox(root, values=["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], width=58, state="readonly")
posting_period_combobox.grid(row=6, column=1, sticky="w", pady=5)

tk.Button(root, text="Run Check", command=run_check, bg="#0078D7", fg="white", width=20).grid(row=7, column=1, pady=10)
tk.Button(root, text="Clear All", command=clear_all_fields, bg="#D9534F", fg="white", width=20).grid(row=7, column=0, pady=10, padx=1)

credit_label = tk.Label(root, text="Created by Zareef Askary", fg="#555555", font=("Arial", 9, "italic"))
credit_label.grid(row=11, column=0, columnspan=3, sticky="e", padx=10, pady=(0, 10))

# Scrollable Text Box for Output
output_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=95, height=20)
output_box.grid(row=8, column=0, columnspan=3, padx=10, pady=10)

root.mainloop()
