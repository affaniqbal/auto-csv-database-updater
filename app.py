

import sys 
import subprocess

# List of required modules
required_modules = ['pandas', 'fuzzywuzzy', 'tkinter', 'tkinterdnd2']

# Check and install missing modules
for module in required_modules:
    try:
        __import__(module)
    except ImportError:
        print(f"Module {module} is not installed. Installing now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", module])

import pandas as pd
from fuzzywuzzy import fuzz
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

def select_file(entry):
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    entry.delete(0, tk.END)
    entry.insert(tk.END, file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    output_file_entry.delete(0, tk.END)
    output_file_entry.insert(tk.END, file_path)

def drop_file(event):
    file_path = event.data
    event.widget.delete(0, tk.END)
    event.widget.insert(tk.END, file_path)

def process_files():
    core_file = core_file_entry.get()
    source_file = source_file_entry.get()
    output_file = output_file_entry.get()

    if not core_file or not source_file or not output_file:
        messagebox.showwarning("Warning", "Please select all required files.")
        return

    min_ratio = float(min_ratio_entry.get())

    core = pd.read_csv(core_file)
    source = pd.read_csv(source_file)

    # Get the first column of each DataFrame
    core_col = core.iloc[:, 0].astype(str)
    source_col = source.iloc[:, 0].astype(str)

    # Find potential matches between source and core
    potential_matches = []
    for value in source_col:
        best_match = None
        max_ratio = 0
        for core_value in core_col:
            ratio = fuzz.partial_ratio(value, core_value)
            if ratio >= min_ratio and ratio > max_ratio:
                best_match = core_value
                max_ratio = ratio
        potential_matches.append((value, best_match))

    # Display the potential matches in a treeview
    match_window = tk.Toplevel(window)
    match_window.title("Potential Matches")

    treeview = ttk.Treeview(match_window, columns=("Source", "Match", "Action"), show="headings")
    treeview.heading("Source", text="Source Value")
    treeview.heading("Match", text="Suggested Match")
    treeview.heading("Action", text="Action")
    treeview.pack(fill=tk.BOTH, expand=True)

    for value, match in potential_matches:
        if match is not None:
            treeview.insert("", tk.END, values=(value, match, "Keep Source"))
        else:
            treeview.insert("", tk.END, values=(value, "No Match", "Keep Source"))

    def update_action(event):
        selected_item = treeview.focus()
        current_action = treeview.item(selected_item, "values")[2]
        if current_action == "Keep Source":
            treeview.set(selected_item, "Action", "Keep Match")
        elif current_action == "Keep Match":
            treeview.set(selected_item, "Action", "Keep Both")
        else:
            treeview.set(selected_item, "Action", "Keep Source")

    treeview.bind("<Double-1>", update_action)



    def confirm_matches():
        confirmed_matches = []
        for item in treeview.get_children():
            values = treeview.item(item, "values")
            source_value = values[0]
            match_value = values[1]
            action = values[2]
            if action == "Keep Source":
                confirmed_matches.append((source_value, source_value))
            elif action == "Keep Match":
                if match_value != "No Match":
                    confirmed_matches.append((source_value, match_value))
            else:  # Keep Both
                confirmed_matches.append((source_value, source_value))
                if match_value != "No Match":
                    confirmed_matches.append((match_value, match_value))

        source_dict = {value[0]: value[1] for value in confirmed_matches}
        source['Matched'] = source.iloc[:, 0].map(source_dict)

        core_dict = {value[1]: value[0] for value in confirmed_matches if value[0] != value[1]}
        core['Matched'] = core.iloc[:, 0].map(core_dict)

        updated_core = pd.concat([core, source[~source['Matched'].isin(core.iloc[:, 0])]], ignore_index=True)
        updated_core = updated_core.iloc[:, :2]  # Select only the first two columns
        updated_core.to_csv(output_file, index=False)
        
        messagebox.showinfo("Success", "Matching process completed. Results saved to the output file.")
        match_window.destroy()


    confirm_button = ttk.Button(match_window, text="Confirm Matches", command=confirm_matches)
    confirm_button.pack()

# Create the main window
window = TkinterDnD.Tk()
window.title("CSV Matching")

# Create a frame for file selection
file_frame = ttk.Frame(window)
file_frame.pack(pady=10)

# Core file selection
core_file_label = ttk.Label(file_frame, text="Core File:")
core_file_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
core_file_entry = ttk.Entry(file_frame, width=50)
core_file_entry.grid(row=0, column=1, padx=5, pady=5)
core_file_entry.drop_target_register(DND_FILES)
core_file_entry.dnd_bind("<<Drop>>", drop_file)
core_file_button = ttk.Button(file_frame, text="Select", command=lambda: select_file(core_file_entry))
core_file_button.grid(row=0, column=2, padx=5, pady=5)

# Source file selection
source_file_label = ttk.Label(file_frame, text="Source File:")
source_file_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
source_file_entry = ttk.Entry(file_frame, width=50)
source_file_entry.grid(row=1, column=1, padx=5, pady=5)
source_file_entry.drop_target_register(DND_FILES)
source_file_entry.dnd_bind("<<Drop>>", drop_file)
source_file_button = ttk.Button(file_frame, text="Select", command=lambda: select_file(source_file_entry))
source_file_button.grid(row=1, column=2, padx=5, pady=5)

# Output file selection
output_file_label = ttk.Label(file_frame, text="Output File:")
output_file_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
output_file_entry = ttk.Entry(file_frame, width=50)
output_file_entry.grid(row=2, column=1, padx=5, pady=5)
output_file_entry.drop_target_register(DND_FILES)
output_file_entry.dnd_bind("<<Drop>>", drop_file)
output_file_button = ttk.Button(file_frame, text="Select", command=select_output_file)
output_file_button.grid(row=2, column=2, padx=5, pady=5)

# Minimum ratio selection
min_ratio_label = ttk.Label(file_frame, text="Minimum Ratio:")
min_ratio_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
min_ratio_entry = ttk.Entry(file_frame, width=10)
min_ratio_entry.insert(0, "80")
min_ratio_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
min_ratio_info_label = ttk.Label(file_frame, text="Lower value: more matches, Higher value: fewer matches")
min_ratio_info_label.grid(row=3, column=1, padx=5, pady=5, sticky=tk.E)

# Create a button to start the matching process
process_button = ttk.Button(window, text="Start Matching", command=process_files)
process_button.pack(pady=10)

# Run the main event loop
window.mainloop()