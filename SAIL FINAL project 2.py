import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os

def apply_styles(widget):
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview",
                    background="#f0f0f0",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="#f0f0f0",
                    font=("Segoe UI", 10))
    style.map('Treeview', background=[('selected', '#0078D7')])
    widget.configure(bg="#f9f9f9")

def load_file(var, filetypes):
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        var.set(path)

def update_progress(percent_text):
    progress_label.config(text=percent_text)
    root.update_idletasks()

def run_comparison():
    master_path = master_file_path.get()
    changes_path = changes_file_path.get()

    if not os.path.exists(master_path) or not os.path.exists(changes_path):
        messagebox.showerror("Error", "One or both file paths are invalid.")
        return

    try:
        update_progress("10% - Loading files...")
        master_df = pd.read_csv(master_path)
        changes_df = pd.read_excel(changes_path) if changes_path.endswith(".xlsx") else pd.read_csv(changes_path)

        update_progress("25% - Cleaning columns...")
        master_df.columns = master_df.columns.str.upper()
        changes_df.columns = changes_df.columns.str.upper()
        master_df.drop(columns=[col for col in master_df.columns if "YYYYMM" in col], inplace=True, errors='ignore')
        changes_df.drop(columns=[col for col in changes_df.columns if "YYYYMM" in col], inplace=True, errors='ignore')

        update_progress("40% - Merging data...")
        merged_df = pd.merge(master_df, changes_df, on="UNIT_PERNO", suffixes=('_OLD', '_NEW'))

        update_progress("60% - Comparing fields...")
        changes = []
        for col in master_df.columns:
            if col == "UNIT_PERNO":
                continue
            old_col = f"{col}_OLD"
            new_col = f"{col}_NEW"
            if old_col in merged_df.columns and new_col in merged_df.columns:
                changed_rows = merged_df[merged_df[old_col] != merged_df[new_col]]
                for _, row in changed_rows.iterrows():
                    changes.append({
                        "Unit Per no": row["UNIT_PERNO"],
                        "Field": col,
                        "Old Value": row[old_col],
                        "New Value": row[new_col]
                    })

        df_changes = pd.DataFrame(changes)
        df_changes.to_csv("Changes_New.csv", index=False)

        update_progress("75% - Extracting new joinees...")
        new_joiners = changes_df[~changes_df["UNIT_PERNO"].isin(master_df["UNIT_PERNO"])]
        fields = ["UNIT_PERNO", "SAIL_PERNO", "PAN", "IFSC_CD", "BANK_ACNO", "UNIT_JOIN_DT", "DOJ_SAIL"]
        new_joiners = new_joiners[[col for col in fields if col in new_joiners.columns]]
        new_joiners.to_csv("New_Joinees.csv", index=False)

        update_progress("90% - Saving reports...")
        count_report = df_changes.groupby("Field").size().reset_index(name="Count")
        count_report = pd.concat([
            count_report,
            pd.DataFrame([{"Field": "new_joinees", "Count": new_joiners.shape[0]}])
        ], ignore_index=True)
        count_report.to_csv("Count.csv", index=False)

        display_data(df_changes)
        update_progress("✅ Done")
        messagebox.showinfo("Success", "Comparison complete. Reports generated.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def display_data(df):
    tree.delete(*tree.get_children())
    tree.configure(columns=list(df.columns))
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=120)
    for _, row in df.iterrows():
        tree.insert("", "end", values=list(row))

def load_report(filename):
    if not os.path.exists(filename):
        messagebox.showerror("Error", f"{filename} not found.")
        return
    df = pd.read_csv(filename)
    display_data(df)

def clear_table():
    tree.delete(*tree.get_children())

# GUI Setup
root = tk.Tk()
root.title("Employee Comparison Tool (BSL_SAIL)")
root.geometry("1080x700")
apply_styles(root)

master_file_path = tk.StringVar()
changes_file_path = tk.StringVar()

frame = tk.Frame(root, bg="#ffffff", bd=2)
frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

header = tk.Label(frame, text="Employee Comparison Tool (BSL_SAIL)", font=("Segoe UI", 23, "bold"), bg="#ffffff", fg="#333333")
header.grid(row=0, column=0, columnspan=3, pady=10)

entry_opts = {'font': ("Segoe UI", 10), 'bg': "#f9f9f9", 'fg': "#333", 'relief': tk.GROOVE}

tk.Label(frame, text="Master CSV File", font=("Segoe UI", 10), bg="#ffffff").grid(row=1, column=0, sticky="w", padx=10, pady=5)
tk.Entry(frame, textvariable=master_file_path, width=75, **entry_opts).grid(row=1, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: load_file(master_file_path, [("CSV files", "*.csv")]), bg="#0078D7", fg="white").grid(row=1, column=2)

tk.Label(frame, text="Changes File (CSV/XLSX)", font=("Segoe UI", 10), bg="#ffffff").grid(row=2, column=0, sticky="w", padx=10, pady=5)
tk.Entry(frame, textvariable=changes_file_path, width=75, **entry_opts).grid(row=2, column=1, padx=5)
tk.Button(frame, text="Browse", command=lambda: load_file(changes_file_path, [("CSV", ".csv"), ("Excel", ".xlsx")]), bg="#0078D7", fg="white").grid(row=2, column=2)

tk.Button(frame, text="Run Comparison", bg="#28a745", fg="white", command=run_comparison).grid(row=3, column=1, pady=10)

progress_label = tk.Label(frame, text="", fg="#0078D7", font=("Segoe UI", 10, "bold"), bg="#ffffff")
progress_label.grid(row=4, column=1)

btn_frame = tk.Frame(frame, bg="#ffffff")
btn_frame.grid(row=5, column=1, pady=10)
tk.Button(btn_frame, text="View Changes Report", command=lambda: load_report("Changes_New.csv"), bg="#6c757d", fg="white").grid(row=0, column=0, padx=5)
tk.Button(btn_frame, text="View Count Report", command=lambda: load_report("Count.csv"), bg="#6c757d", fg="white").grid(row=0, column=1, padx=5)
tk.Button(btn_frame, text="View New Joiners", command=lambda: load_report("New_Joinees.csv"), bg="#6c757d", fg="white").grid(row=0, column=2, padx=5)
tk.Button(btn_frame, text="Clear Table", bg="#dc3545", fg="white", command=clear_table).grid(row=0, column=3, padx=5)

tree = ttk.Treeview(frame, show="headings")
tree.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar_y.set)
scrollbar_y.grid(row=6, column=3, sticky='ns')

frame.grid_rowconfigure(6, weight=1)
frame.grid_columnconfigure(1, weight=1)

root.mainloop()