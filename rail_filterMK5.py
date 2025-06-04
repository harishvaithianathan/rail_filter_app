import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

toggle_button_state = False
checkboxes = []
cost_columns = []
df = pd.DataFrame()
tree = None

def load_data(file_path):
    df = pd.read_excel(file_path, engine='openpyxl', header=16)  # Header row is index 16 (row 17)
    df.columns = df.columns.str.strip()
    return df

def filter_routes(df, start_location=None, end_location=None, selected_columns=None):
    if start_location:
        df = df[df['Origin'].str.contains(start_location, case=False, na=False)]
    if end_location:
        df = df[df['Destination'].str.contains(end_location, case=False, na=False)]

    if selected_columns:
        base_columns = ['Origin', 'Destination']
        columns_to_show = base_columns + [col for col in selected_columns if col in df.columns]
        df = df[columns_to_show]

    return df

def display_data(dataframe):
    global tree
    if tree:
        tree.destroy()

    tree = ttk.Treeview(root)
    tree.grid(row=99, column=0, columnspan=4, sticky='nsew', padx=10, pady=10)
    tree.bind("<<TreeviewSelect>>", on_row_selected)

    # Configure column weights for resizing
    root.grid_rowconfigure(99, weight=1)
    root.grid_columnconfigure(0, weight=1)

    tree["columns"] = list(dataframe.columns)
    tree["show"] = "headings"

    for col in dataframe.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center")

    for _, row in dataframe.iterrows():
        tree.insert("", "end", values=list(row))

def on_row_selected(event):
    global tree, df

    selected_item = tree.selection()
    if not selected_item:
        return

    row_values = tree.item(selected_item[0], 'values')
    displayed_columns = tree["columns"]

    try:
        origin_idx = displayed_columns.index("Origin")
        dest_idx = displayed_columns.index("Destination")
    except ValueError:
        messagebox.showerror("Error", "Origin or Destination column missing in the preview.")
        return

    origin_val = row_values[origin_idx]
    dest_val = row_values[dest_idx]

    full_row = df[(df['Origin'] == origin_val) & (df['Destination'] == dest_val)]

    if full_row.empty:
        messagebox.showerror("Error", "Could not find full data for the selected row.")
        return

    full_row_series = full_row.iloc[0]

    if hasattr(root, 'detail_popup') and root.detail_popup.winfo_exists():
        root.detail_popup.destroy()

    popup = tk.Toplevel(root)
    popup.title("Row Details")
    popup.geometry("400x600")
    root.detail_popup = popup

    tk.Label(popup, text="Row Details", font=("Helvetica", 16, "bold")).pack(pady=12)

    # Frame with canvas and scrollbar
    container = tk.Frame(popup)
    container.pack(fill="both", expand=True, padx=12, pady=6)

    canvas = tk.Canvas(container, borderwidth=0, highlightthickness=0)
    scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Use grid for two columns: labels and values
    for i, (col, val) in enumerate(full_row_series.items()):
        bg_color = "#f9f9f9" if i % 2 == 0 else "#ffffff"  # alternate row colors

        col_label = tk.Label(scrollable_frame, text=f"{col}:", font=("Helvetica", 11, "bold"),
                             anchor="w", background=bg_color)
        col_label.grid(row=i, column=0, sticky="w", padx=(5, 10), pady=4)

        val_label = tk.Label(scrollable_frame, text=str(val), font=("Helvetica", 11),
                             anchor="w", background=bg_color, wraplength=250, justify="left")
        val_label.grid(row=i, column=1, sticky="w", padx=(0, 5), pady=4)

    # Make columns resize nicely
    scrollable_frame.grid_columnconfigure(0, weight=1, uniform="col")
    scrollable_frame.grid_columnconfigure(1, weight=2, uniform="col")


def reload_data():
    start_location = start_entry.get()
    end_location = end_entry.get()
    selected_columns = [var.get() for var in checkboxes if var.get()]

    if not selected_columns:
        messagebox.showwarning("No Columns Selected", "Please select at least one cost column.")
        return

    filtered_df = filter_routes(df, start_location, end_location, selected_columns)

    if filtered_df.empty:
        messagebox.showinfo("No Results", "No routes found for the given criteria.", parent=root)
    else:
        display_data(filtered_df)

def export_data():
    if tree is None or not tree.get_children():
        messagebox.showinfo("Nothing to Export", "No data to export.")
        return

    save_path = filedialog.asksaveasfilename(parent=root, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        data = [tree.item(child)["values"] for child in tree.get_children()]
        columns = tree["columns"]
        pd.DataFrame(data, columns=columns).to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"Data exported to {save_path}")

def select_file():
    global df, checkboxes, cost_columns, toggle_button_state

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    df = load_data(file_path)

    for widget in root.grid_slaves():
        if int(widget.grid_info()["row"]) >= 4 and widget != tree:
            widget.destroy()

    tk.Label(root, text="Select Cost Columns:").grid(row=3, column=0, columnspan=4, padx=10, pady=5)

    cost_columns = [col for col in df.columns if col not in ['Origin', 'Destination']]
    checkboxes.clear()

    num_columns = 4
    for idx, col in enumerate(cost_columns):
        var = tk.StringVar(value="")
        chk = tk.Checkbutton(root, text=col, variable=var, onvalue=col, offvalue="")
        row = 4 + (idx // num_columns)
        col_position = idx % num_columns
        chk.grid(row=row, column=col_position, padx=10, pady=2, sticky='w')
        checkboxes.append(var)

    def toggle_checkboxes():
        global toggle_button_state
        new_value = cost_columns if not toggle_button_state else ["" for _ in cost_columns]
        for var, val in zip(checkboxes, new_value):
            var.set(val)
        toggle_button_state = not toggle_button_state
        toggle_btn.config(text="Deselect All" if toggle_button_state else "Select All")

    toggle_button_state = False
    toggle_btn = tk.Button(root, text="Select All", command=toggle_checkboxes)
    toggle_btn.grid(row=4 + (len(cost_columns) // num_columns) + 1, column=0, columnspan=2, padx=10, pady=5)

    reload_btn = tk.Button(root, text="Submit", command=reload_data)
    reload_btn.grid(row=4 + (len(cost_columns) // num_columns) + 2, column=0, columnspan=2, padx=10, pady=5)

    export_btn = tk.Button(root, text="Export to Excel", command=export_data)
    export_btn.grid(row=4 + (len(cost_columns) // num_columns) + 2, column=2, columnspan=2, padx=10, pady=5)

    filename = file_path.split("/")[-1]
    file_label.config(text=f"Loaded: {filename}", fg="green")

# GUI setup
root = tk.Tk()
root.title("Rail Route Filter")

def center_window(win, width=800, height=600):
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')

center_window(root)

def bring_root_to_front():
    root.lift()
    root.attributes("-topmost", True)
    root.after(500, lambda: root.attributes("-topmost", False))
    root.focus_force()

root.after(0, bring_root_to_front)

file_label = tk.Label(root, text="No file selected", fg="gray")
file_label.grid(row=0, column=2, columnspan=2, padx=10, pady=10, sticky="w")

tk.Button(root, text="Select Excel File", command=select_file).grid(row=0, column=0, columnspan=2, padx=10, pady=10)

tk.Label(root, text="Start Location:").grid(row=1, column=0, padx=10, pady=5)
start_entry = tk.Entry(root)
start_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="End Location:").grid(row=2, column=0, padx=10, pady=5)
end_entry = tk.Entry(root)
end_entry.grid(row=2, column=1, padx=10, pady=5)


root.mainloop()
