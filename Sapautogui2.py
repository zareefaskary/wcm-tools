import os
import time
import threading
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import tkinter.ttk as ttk  # Required for Combobox

# === Global stop flag ===
stop_flag = False

class DynamicToolTip:
    def __init__(self, widget):
        self.widget = widget
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        if isinstance(self.widget, (tk.Entry, ttk.Combobox)):
            text = self.widget.get()
            if not text:
                return

        if self.tooltip_window:
            return

        x = self.widget.winfo_rootx() + 25
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = tk.Label(tw, text=text, justify=tk.LEFT,
                        background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                        font=("Arial", 9, "normal"), padx=5, pady=3,
                        wraplength=500)
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def write_log(message):
    log_text.config(state="normal")
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.config(state="disabled")

def browse_file(entry):
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def browse_folder(entry):
    folder = filedialog.askdirectory()
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)

def stop_script():
    global stop_flag
    stop_flag = True
    messagebox.showinfo("Stopping", "🛑 Script stop requested. It will stop after the current company finishes.")

def handle_export_permission(session):
    """
    Blind handler for the SAP Security/Save popup.
    If no popup exists, it exits silently in 0.5 seconds.
    """
    try:
        # Give the system a moment to realize a popup is needed
        time.sleep(1) 
        
        # Check if a new window (wnd[1]) appeared AFTER we clicked Save
        # In SAP, session.Children.Count tells you how many windows are open.
        if session.Children.Count > 1: 
            # Look for common 'Allow', 'Yes', or 'Grant' buttons
            # 1. Standard SAP Security 'Allow' button
            # 2. Standard 'Yes' button
            # 3. Enter Key (sendVKey 0)
            
            btn_ids = [
                "wnd[1]/usr/btnSPOP-OPTION1", 
                "wnd[1]/usr/btnSPOP-VAR_SAVE", 
                "wnd[1]/tbar[0]/btn[0]",
                "wnd[1]/usr/btnSPOP-VARIANT_SAVE"        
            ]
            
            for bid in btn_ids:
                btn = session.findById(bid, False)
                if btn:
                    btn.press()
                    write_log(f"✅ Security popup handled using: {bid}")
                    return

            # If no button found, try pressing 'Enter' as a last resort
            session.findById("wnd[1]").sendVKey(0)
            write_log("✅ Security popup dismissed via Enter key.")
            
    except Exception:
        # If anything fails, we assume there was no popup and continue
        pass

def run_script_thread():
    thread = threading.Thread(target=run_script)
    thread.start()

def run_script():
    global stop_flag
    stop_flag = False

    try:
        EXCEL_FILE = excel_file_entry.get()
        SHEET_NAME = sheet_name_entry.get()
        OUTPUT_FOLDER = output_folder_entry.get()
        SAP_SYSTEM = sap_system_entry.get()
        POSTING_PERIOD = posting_period_combobox.get()  # Get value from Combobox
        YEAR = year_entry.get()
        LAYOUT = layout_entry.get()
        COMPANY_COL = company_col_entry.get()
        PLANT_COL = plant_col_entry.get()

        if not all([EXCEL_FILE, SHEET_NAME, OUTPUT_FOLDER, SAP_SYSTEM, POSTING_PERIOD, YEAR, LAYOUT, COMPANY_COL, PLANT_COL]):
            messagebox.showerror("Missing Input", "Please fill in all fields before running.")
            return

        write_log("📂 Reading Excel...")
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str)
        df[COMPANY_COL] = df[COMPANY_COL].str.zfill(4)
        df[PLANT_COL] = df[PLANT_COL].str.zfill(4)
        company_to_plants = df.groupby(COMPANY_COL)[PLANT_COL].apply(list).to_dict()

        write_log("🔌 Connecting to SAP...")
        SapGuiAuto = win32.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.OpenConnection(SAP_SYSTEM, True)
        connection = application.Children(0)
        session = connection.Children(0)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        session.findById("wnd[0]/tbar[0]/okcd").text = "ysam"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        try:
            session.findById("wnd[0]/usr/lbl[5,4]").setFocus
            session.findById("wnd[0]/usr/lbl[5,4]").caretPosition = 1
            session.findById("wnd[0]").sendVKey(2)
            session.findById("wnd[0]/usr/lbl[9,16]").setFocus
            session.findById("wnd[0]/usr/lbl[9,16]").caretPosition = 0
            session.findById("wnd[0]").sendVKey(2)
            session.findById("wnd[0]/usr/lbl[82,31]").setFocus
            session.findById("wnd[0]/usr/lbl[82,31]").caretPosition = 0
            session.findById("wnd[0]").sendVKey(2)
            session.findById("wnd[1]/usr/btnWEITER").press()
        except:
            write_log("⚠️ Navigation screens not found, continuing...")

        write_log("✅ Inside YSAM input screen — starting company loop...\n")

        for company, plants in company_to_plants.items():
            if stop_flag:
                write_log("🛑 Script stopped by user.")
                messagebox.showinfo("Stopped", "🛑 Script stopped by user.")
                return

            try:
                session.findById("wnd[0]/usr/ctxtPBUKRS").text = company
                session.findById("wnd[0]/usr/txtPGJAHR").text = YEAR
                session.findById("wnd[0]/usr/ctxtPGMON").text = POSTING_PERIOD

                if len(plants) > 0:
                    session.findById("wnd[0]/usr/ctxtSWERKS-LOW").setFocus
                    session.findById("wnd[0]/usr/ctxtSWERKS-LOW").caretPosition = 0
                    session.findById("wnd[0]/usr/btn%_SWERKS_%_APP_%-VALU_PUSH").press()
                    time.sleep(0.5)

                    try:
                        session.findById("wnd[1]/tbar[0]/btn[16]").press()
                        time.sleep(0.3)
                    except:
                        pass

                    for idx, plant in enumerate(plants):
                        if stop_flag:
                            write_log("🛑 Script stopped during plant entry.")
                            messagebox.showinfo("Stopped", "🛑 Script stopped by user.")
                            return

                        field = (
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                            "ssubSCREEN_HEADER:SAPLALDB:3010/"
                            "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]"
                        )

                        session.findById(field).text = plant
                        session.findById(field).caretPosition = len(plant)

                        if idx < len(plants) - 1:
                            session.findById("wnd[1]/tbar[0]/btn[13]").press()
                            time.sleep(0.2)

                    session.findById("wnd[1]/tbar[0]/btn[8]").press()

                session.findById("wnd[0]/usr/ctxtP_VARI").text = LAYOUT
                session.findById("wnd[0]").sendVKey(8)

                session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
                filename = f"YSAM_{YEAR}0{POSTING_PERIOD}_{company}.XLSX"
                session.findById("wnd[1]/usr/ctxtDY_PATH").text = OUTPUT_FOLDER
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                session.findById("wnd[1]/tbar[0]/btn[0]").press()

                handle_export_permission(session)

                session.findById("wnd[0]").sendVKey(3)
                time.sleep(1)
                write_log(f"✅ Exported company {company} successfully!")

            except Exception as e:
                write_log(f"⚠️ Error processing company {company}: {e}")
                time.sleep(1)

        messagebox.showinfo("Done", "🎉 All exports completed successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def clear_all_fields():
    excel_file_entry.delete(0, tk.END)
    sheet_name_entry.delete(0, tk.END)
    output_folder_entry.delete(0, tk.END)
    sap_system_entry.delete(0, tk.END)
    posting_period_combobox.set("")  # Clear Combobox
    year_entry.delete(0, tk.END)
    layout_entry.delete(0, tk.END)
    company_col_entry.delete(0, tk.END)
    plant_col_entry.delete(0, tk.END)

# === GUI Setup ===
root = tk.Tk()
root.title("YSAM Export Automation")
root.geometry("800x760")
root.resizable(False, False)

# Excel File
tk.Label(root, text="Excel File:", anchor="w", width=20).grid(row=0, column=0, padx=10, pady=6, sticky="w")
excel_file_entry = tk.Entry(root, width=60)
excel_file_entry.grid(row=0, column=1, padx=5, pady=6)
DynamicToolTip(excel_file_entry)
tk.Button(root, text="Browse", command=lambda: browse_file(excel_file_entry)).grid(row=0, column=2, padx=0, pady=6, sticky="w")

# Sheet Name
tk.Label(root, text="Sheet Name:", anchor="w", width=20).grid(row=1, column=0, padx=10, pady=6, sticky="w")
sheet_name_entry = tk.Entry(root, width=60)
sheet_name_entry.grid(row=1, column=1, padx=5, pady=6)
DynamicToolTip(sheet_name_entry)

# Output Folder
tk.Label(root, text="Output Folder:", anchor="w", width=20).grid(row=2, column=0, padx=10, pady=6, sticky="w")
output_folder_entry = tk.Entry(root, width=60)
output_folder_entry.grid(row=2, column=1, padx=5, pady=6)
DynamicToolTip(output_folder_entry)
tk.Button(root, text="Browse", command=lambda: browse_folder(output_folder_entry)).grid(row=2, column=2, padx=0, pady=6, sticky="w")

# SAP System
tk.Label(root, text="SAP System:", anchor="w", width=20).grid(row=3, column=0, padx=10, pady=6, sticky="w")
sap_system_entry = tk.Entry(root, width=60)
sap_system_entry.grid(row=3, column=1, padx=5, pady=6)
DynamicToolTip(sap_system_entry)

# Posting Period (Combobox)
tk.Label(root, text="Posting Period:", anchor="w", width=20).grid(row=4, column=0, padx=10, pady=6, sticky="w")
posting_period_combobox = ttk.Combobox(root, values=["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"], width=58, state="readonly")
posting_period_combobox.grid(row=4, column=1, padx=5, pady=6)
DynamicToolTip(posting_period_combobox)

# Year
tk.Label(root, text="Year:", anchor="w", width=20).grid(row=5, column=0, padx=10, pady=6, sticky="w")
year_entry = tk.Entry(root, width=60)
year_entry.grid(row=5, column=1, padx=5, pady=6)
DynamicToolTip(year_entry)

# Layout
tk.Label(root, text="Layout:", anchor="w", width=20).grid(row=6, column=0, padx=10, pady=6, sticky="w")
layout_entry = tk.Entry(root, width=60)
layout_entry.grid(row=6, column=1, padx=5, pady=6)
DynamicToolTip(layout_entry)

# Company Code Column
tk.Label(root, text="Company Code Column:", anchor="w", width=20).grid(row=7, column=0, padx=10, pady=6, sticky="w")
company_col_entry = tk.Entry(root, width=60)
company_col_entry.grid(row=7, column=1, padx=5, pady=6)
DynamicToolTip(company_col_entry)

# Plant Code Column
tk.Label(root, text="Plant Code Column:", anchor="w", width=20).grid(row=8, column=0, padx=10, pady=6, sticky="w")
plant_col_entry = tk.Entry(root, width=60)
plant_col_entry.grid(row=8, column=1, padx=5, pady=6)
DynamicToolTip(plant_col_entry)

# Default values
sheet_name_entry.insert(0, "Plant list_All")
company_col_entry.insert(0, "Company Code")
plant_col_entry.insert(0, "Plant Code")

# Buttons
button_frame = tk.Frame(root)
button_frame.grid(row=9, column=1, pady=20)

tk.Button(button_frame, text="Run Script", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=10, command=run_script_thread).grid(row=0, column=0, padx=10)
tk.Button(button_frame, text="Stop Script", bg="#E53935", fg="white", font=("Arial", 12, "bold"), width=10, command=stop_script).grid(row=0, column=1, padx=10)
tk.Button(button_frame, text="Clear All", bg="#FFA500", fg="white", font=("Arial", 12, "bold"), width=10, command=clear_all_fields).grid(row=0, column=2, padx=10)

# Log Text Widget
log_text = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20, width=95)
log_text.grid(row=10, column=0, columnspan=3, padx=10, pady=10)

credit_label = tk.Label(root, text="Created by Zareef Askary", fg="#555555", font=("Arial", 9, "italic"))
credit_label.grid(row=11, column=0, columnspan=3, sticky="e", padx=10, pady=(0, 10))

root.mainloop()
