import sys
import subprocess

# List of required modules
required_modules = ['pandas', 'fuzzywuzzy', 'tkinter', 'tkinterdnd2', 'python-Levenshtein']

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

def select_file(text_widget):
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    text_widget.delete('1.0', tk.END)
    text_widget.insert(tk.END, file_path)

def select_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    output_file_text.delete('1.0', tk.END)
    output_file_text.insert(tk.END, file_path)

def drop_file(event):
    file_path = event.data
    event.widget.delete('1.0', tk.END)
    event.widget.insert(tk.END, file_path)

def select_columns():
    core_file = core_file_text.get('1.0', tk.END).strip()
    source_file = source_file_text.get('1.0', tk.END).strip()

    if not core_file or not source_file:
        messagebox.showwarning("Warning", "Please select the core and source files.")
        return

    core = pd.read_csv(core_file)
    source = pd.read_csv(source_file)

    column_window = tk.Toplevel(window)
    column_window.title("Select Columns")

    core_label = ttk.Label(column_window, text="Core File:")
    core_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    core_columns = ttk.Combobox(column_window, values=list(core.columns), state="readonly")
    core_columns.grid(row=0, column=1, padx=5, pady=5)
    core_columns.current(0)

    source_label = ttk.Label(column_window, text="Source File:")
    source_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
    source_columns = ttk.Combobox(column_window, values=list(source.columns), state="readonly")
    source_columns.grid(row=1, column=1, padx=5, pady=5)
    source_columns.current(0)

    core_preview_label = ttk.Label(column_window, text="Core File Preview:")
    core_preview_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
    core_preview_text = tk.Text(column_window, wrap=tk.WORD, width=60, height=5)
    core_preview_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)
    core_preview_text.insert(tk.END, core.head().to_string(index=False))

    source_preview_label = ttk.Label(column_window, text="Source File Preview:")
    source_preview_label.grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
    source_preview_text = tk.Text(column_window, wrap=tk.WORD, width=60, height=5)
    source_preview_text.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
    source_preview_text.insert(tk.END, source.head().to_string(index=False))

    def confirm_columns():
        core_column = core_columns.get()
        source_column = source_columns.get()
        column_window.destroy()
        process_files(core_column, source_column)

    confirm_button = ttk.Button(column_window, text="Confirm", command=confirm_columns)
    confirm_button.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

def process_files(core_column, source_column):
    core_file = core_file_text.get('1.0', tk.END).strip()
    source_file = source_file_text.get('1.0', tk.END).strip()
    output_file = output_file_text.get('1.0', tk.END).strip()

    if not core_file or not source_file or not output_file:
        messagebox.showwarning("Warning", "Please select all required files.")
        return

    min_ratio = float(min_ratio_entry.get())

    core = pd.read_csv(core_file)
    source = pd.read_csv(source_file)

    core_col = core[core_column].astype(str)
    source_col = source[source_column].astype(str)

    # Find potential matches between source and core
    potential_matches = []
    for value in source_col:
        best_match = None
        max_ratio = 0
        for core_value in core_col:
            ratio = fuzz.ratio(value, core_value)
            if ratio >= min_ratio and ratio > max_ratio:
                best_match = core_value
                max_ratio = ratio
        potential_matches.append((value, best_match, max_ratio))

    # Display the potential matches in a treeview
    match_window = tk.Toplevel(window)
    match_window.title("Potential Matches")

    treeview = ttk.Treeview(match_window, columns=("Source", "Match", "Ratio", "Action"), show="headings")
    treeview.heading("Source", text="Source Value")
    treeview.heading("Match", text="Suggested Match")
    treeview.heading("Ratio", text="Match Ratio")
    treeview.heading("Action", text="Action")
    treeview.pack(fill=tk.BOTH, expand=True)

    def update_action(item, action):
        source_value = treeview.item(item, "values")[0]
        match_value = treeview.item(item, "values")[1]

        if action == "Keep Source":
            treeview.set(item, "Match", source_value)
        elif action == "Keep Match":
            treeview.set(item, "Match", match_value)
        else:  # Keep Both
            treeview.insert("", tk.END, values=(match_value, match_value, 100, ""))

    for value, match, ratio in potential_matches:
        if match is not None:
            item = treeview.insert("", tk.END, values=(value, match, ratio, ""))
        else:
            item = treeview.insert("", tk.END, values=(value, "No Match", 0, ""))

        button_frame = ttk.Frame(match_window)
        treeview.set(item, "Action", "")
        treeview.item(item, tags=("button_frame",))
        button_frame.grid(row=treeview.index(item), column=3, sticky="nsew")

        keep_source_button = ttk.Button(button_frame, text="Keep Source", command=lambda i=item: update_action(i, "Keep Source"))
        keep_source_button.pack(side=tk.LEFT, padx=5)

        keep_match_button = ttk.Button(button_frame, text="Keep Match", command=lambda i=item: update_action(i, "Keep Match"))
        keep_match_button.pack(side=tk.LEFT, padx=5)

        keep_both_button = ttk.Button(button_frame, text="Keep Both", command=lambda i=item: update_action(i, "Keep Both"))
        keep_both_button.pack(side=tk.LEFT, padx=5)

    def confirm_matches():
        confirmed_matches = []
        for item in treeview.get_children():
            if treeview.tag_has("button_frame", item):
                values = treeview.item(item, "values")
                source_value = values[0]
                match_value = values[1]
                if match_value != "No Match":
                    confirmed_matches.append((source_value, match_value))

        source_dict = {value[0]: value[1] for value in confirmed_matches}
        source['Matched'] = source[source_column].map(source_dict)

        core_dict = {value[1]: value[0] for value in confirmed_matches if value[0] != value[1]}
        core['Matched'] = core[core_column].map(core_dict)

        updated_core = pd.concat([core, source[~source['Matched'].isin(core[core_column])]], ignore_index=True)
        updated_core = updated_core[['Matched', core_column]]
        updated_core.to_csv(output_file, index=False)

        messagebox.showinfo("Success", "Matching process completed. Results saved to the output file.")
        match_window.destroy()

    confirm_button = ttk.Button(match_window, text="Confirm Matches", command=confirm_matches)
    confirm_button.pack(pady=10)

# Create the main window
window = TkinterDnD.Tk()
window.title("CSV Matching")

# Create a frame for file selection
file_frame = ttk.Frame(window)
file_frame.pack(pady=10)

# Core file selection
core_file_label = ttk.Label(file_frame, text="Core File:")
core_file_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
core_file_text = tk.Text(file_frame, wrap=tk.WORD, width=60, height=2)
core_file_text.grid(row=0, column=1, padx=5, pady=5)
core_file_text.drop_target_register(DND_FILES)
core_file_text.dnd_bind("<<Drop>>", drop_file)
core_file_button = ttk.Button(file_frame, text="Select", command=lambda: select_file(core_file_text))
core_file_button.grid(row=0, column=2, padx=5, pady=5)

# Source file selection
source_file_label = ttk.Label(file_frame, text="Source File:")
source_file_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
source_file_text = tk.Text(file_frame, wrap=tk.WORD, width=60, height=2)
source_file_text.grid(row=1, column=1, padx=5, pady=5)
source_file_text.drop_target_register(DND_FILES)
source_file_text.dnd_bind("<<Drop>>", drop_file)
source_file_button = ttk.Button(file_frame, text="Select", command=lambda: select_file(source_file_text))
source_file_button.grid(row=1, column=2, padx=5, pady=5)

# Output file selection
output_file_label = ttk.Label(file_frame, text="Output File:")
output_file_label.grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
output_file_text = tk.Text(file_frame, wrap=tk.WORD, width=60, height=2)
output_file_text.grid(row=2, column=1, padx=5, pady=5)
output_file_text.drop_target_register(DND_FILES)
output_file_text.dnd_bind("<<Drop>>", drop_file)
output_file_button = ttk.Button(file_frame, text="Select", command=select_output_file)
output_file_button.grid(row=2, column=2, padx=5, pady=5)

# Minimum ratio selection
min_ratio_label = ttk.Label(file_frame, text="Minimum Ratio:")
min_ratio_label.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
min_ratio_entry = ttk.Entry(file_frame, width=10)
min_ratio_entry.insert(0, "80")
min_ratio_entry.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
min_ratio_info_label = ttk.Label(file_frame, text="Lower value: more matches, Higher value: fewer matches")
min_ratio_info_label.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)

window.minsize(800, 200)  # Set the minimum window size to 800x200 pixels

# Create a button to proceed to column selection
proceed_button = ttk.Button(window, text="Proceed to Column Selection", command=select_columns)
proceed_button.pack(pady=10)

# Run the main event loop
window.mainloop()
