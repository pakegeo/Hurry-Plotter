import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re
import os

# ============================================================
# Hurry Plotter
# A GUI application for plotting metrics from spreadsheet
# and tabular data files.
# ============================================================

def clean_title(text):
    """Remove units/parentheses from plot titles."""
    return re.sub(r"\s*\(.*?\)", "", str(text)).strip()

def clean_column_name(text):
    """Clean column names by removing non-breaking spaces and outer spaces."""
    return str(text).replace("\xa0", " ").strip()

def normalize_name(text):
    """Normalize names for robust column matching."""
    return str(text).replace("\xa0", " ").replace(" ", "").strip().lower()

def get_algorithm_label(sheet_name):
    """
    Create a short legend label from a sheet/table name.
    Example:
        RENO_set1 -> RENO
        CUBIC_set1 -> CUBIC
        BBR_1ST_TRY_set1 -> BBR
    """
    return str(sheet_name).strip().split("_")[0].split(" ")[0]

def draw_logo(canvas):
    """Draw a simple wizard-hat plotting logo."""
    canvas.delete("all")

    # Outer badge
    canvas.create_oval(8, 8, 122, 122, fill="#ffffff", outline="#d7dde5", width=2)

    # Axes
    canvas.create_line(35, 92, 102, 92, width=3, fill="#1f2937", arrow=tk.LAST)
    canvas.create_line(35, 92, 35, 32, width=3, fill="#1f2937", arrow=tk.LAST)
    canvas.create_text(110, 94, text="x", font=("Segoe UI", 10, "bold"), fill="#1f2937")
    canvas.create_text(35, 22, text="y", font=("Segoe UI", 10, "bold"), fill="#1f2937")

    # Smooth curve
    canvas.create_line(
        [40, 88, 52, 78, 65, 84, 78, 58, 96, 48],
        fill="#2563eb",
        width=3,
        smooth=True
    )

    # Generic wizard hat
    canvas.create_polygon(
        64, 24,
        49, 78,
        93, 78,
        fill="#2f2a75",
        outline="#111827",
        width=2
    )

    canvas.create_oval(
        41, 71,
        101, 88,
        fill="#2f2a75",
        outline="#111827",
        width=2
    )

    canvas.create_line(55, 64, 88, 64, fill="#facc15", width=4)
    canvas.create_text(70, 48, text="✦", fill="#facc15", font=("Segoe UI", 13, "bold"))

# ============================================================
# Global state
# ============================================================

file_path = None
xls = None
sheet_map = {}
sheet_names = []
all_columns_global = []
last_fig = None
data_mode = None
table_df = None

color_map = {
    "RENO": "blue",
    "CUBIC": "green",
    "BBR": "red"
}

# ============================================================
# Main window
# ============================================================

root = tk.Tk()
root.title("Hurry Plotter")
root.geometry("1100x720")
root.minsize(900, 620)

# Works on Windows. If unsupported, the exception is ignored.
try:
    root.state("zoomed")
except tk.TclError:
    pass

root.configure(bg="#f3f6fa")

style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Title.TLabel",
    font=("Georgia", 30, "bold italic"),
    background="#f3f6fa",
    foreground="#3b0764"
)

style.configure(
    "Subtitle.TLabel",
    font=("Segoe UI", 11),
    background="#f3f6fa",
    foreground="#6b7280"
)

style.configure(
    "Header.TLabel",
    font=("Segoe UI", 14, "bold"),
    background="white",
    foreground="#111827"
)

style.configure(
    "Info.TLabel",
    font=("Segoe UI", 11),
    background="#f3f6fa",
    foreground="#374151"
)

style.configure("TCombobox", font=("Segoe UI", 12))

style.configure(
    "Import.TButton",
    font=("Segoe UI", 11, "bold"),
    padding=(16, 10),
    background="#2563eb",
    foreground="white"
)
style.map("Import.TButton", background=[("active", "#1d4ed8")])

style.configure(
    "Accent.TButton",
    font=("Segoe UI", 11, "bold"),
    padding=(16, 10),
    background="#22c55e",
    foreground="white"
)
style.map("Accent.TButton", background=[("active", "#15803d")])

style.configure(
    "Secondary.TButton",
    font=("Segoe UI", 11, "bold"),
    padding=(16, 10),
    background="#e5e7eb",
    foreground="#111827"
)
style.map("Secondary.TButton", background=[("active", "#9ca3af")])

# ============================================================
# Header
# ============================================================

header_frame = tk.Frame(root, bg="#f3f6fa")
header_frame.pack(fill="x", padx=35, pady=(18, 6))

title_block = tk.Frame(header_frame, bg="#f3f6fa")
title_block.pack(side="left", anchor="w")

ttk.Label(title_block, text="Hurry Plotter", style="Title.TLabel").pack(anchor="w")
ttk.Label(title_block, text="Metrics Plotting tool", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

logo_canvas = tk.Canvas(
    header_frame,
    width=130,
    height=130,
    bg="#f3f6fa",
    highlightthickness=0
)
logo_canvas.pack(side="right", padx=10)
draw_logo(logo_canvas)

# ============================================================
# File info
# ============================================================

file_frame = tk.Frame(root, bg="#f3f6fa")
file_frame.pack(fill="x", padx=35, pady=(2, 6))

file_label = ttk.Label(
    file_frame,
    text="No file selected",
    style="Info.TLabel"
)
file_label.pack(side="left")

# ============================================================
# Main panels
# ============================================================

main_frame = tk.Frame(root, bg="#f3f6fa")
main_frame.pack(fill="both", expand=True, padx=25, pady=6)

left_frame = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
left_frame.pack(side="left", fill="both", expand=True, padx=(0, 12), pady=8)

right_frame = tk.Frame(main_frame, bg="white", bd=1, relief="solid")
right_frame.pack(side="right", fill="both", expand=True, padx=(12, 0), pady=8)

# ============================================================
# Sheets / Tables panel
# ============================================================

ttk.Label(left_frame, text="Sheets / Tables", style="Header.TLabel").pack(
    anchor="w",
    padx=18,
    pady=(18, 4)
)

tk.Label(
    left_frame,
    text="All sheets/tables are detected automatically.",
    font=("Segoe UI", 10),
    bg="white",
    fg="#6b7280"
).pack(anchor="w", padx=18, pady=(0, 10))

sheet_listbox = tk.Listbox(
    left_frame,
    selectmode=tk.MULTIPLE,
    font=("Segoe UI", 12),
    height=18,
    width=42,
    exportselection=False,
    relief="flat",
    bd=0,
    highlightthickness=1,
    highlightbackground="#d1d5db",
    selectbackground="#2563eb",
    selectforeground="white"
)
sheet_listbox.pack(fill="both", expand=True, padx=18, pady=(0, 18))

# ============================================================
# Axis / Chart options panel
# ============================================================

ttk.Label(right_frame, text="X-axis", style="Header.TLabel").pack(
    anchor="w",
    padx=18,
    pady=(18, 4)
)

selected_x_column = tk.StringVar()

x_combo = ttk.Combobox(
    right_frame,
    textvariable=selected_x_column,
    values=[],
    state="readonly",
    font=("Segoe UI", 12),
    width=50
)
x_combo.pack(anchor="w", padx=18, pady=(5, 18))

ttk.Label(right_frame, text="Y-axis", style="Header.TLabel").pack(
    anchor="w",
    padx=18,
    pady=(10, 4)
)

selected_y_column = tk.StringVar()

y_combo = ttk.Combobox(
    right_frame,
    textvariable=selected_y_column,
    values=[],
    state="readonly",
    font=("Segoe UI", 12),
    width=50
)
y_combo.pack(anchor="w", padx=18, pady=(5, 18))

ttk.Label(right_frame, text="Chart type", style="Header.TLabel").pack(
    anchor="w",
    padx=18,
    pady=(10, 4)
)

selected_chart_type = tk.StringVar()

chart_type_combo = ttk.Combobox(
    right_frame,
    textvariable=selected_chart_type,
    values=["Smooth Line", "Raw Line", "Line + Markers", "Scatter", "Bar", "Step"],
    state="readonly",
    font=("Segoe UI", 12),
    width=50
)
chart_type_combo.pack(anchor="w", padx=18, pady=(5, 18))
chart_type_combo.current(0)

# ============================================================
# Preview area
# ============================================================

preview_frame = tk.Frame(root, bg="#f3f6fa")
preview_frame.pack(fill="x", padx=35, pady=(0, 6))

selected_sheets_label = ttk.Label(
    preview_frame,
    text="Selected sheets/tables: none",
    style="Info.TLabel"
)
selected_sheets_label.pack(anchor="w", pady=3)

selected_x_label = ttk.Label(
    preview_frame,
    text="X-axis: -",
    style="Info.TLabel"
)
selected_x_label.pack(anchor="w", pady=3)

selected_y_label = ttk.Label(
    preview_frame,
    text="Y-axis: -",
    style="Info.TLabel"
)
selected_y_label.pack(anchor="w", pady=3)

selected_chart_label = ttk.Label(
    preview_frame,
    text="Chart type: Smooth Line",
    style="Info.TLabel"
)
selected_chart_label.pack(anchor="w", pady=3)

# ============================================================
# Utility functions
# ============================================================

def update_preview(event=None):
    selected_indices = sheet_listbox.curselection()
    selected_items = [sheet_names[i] for i in selected_indices]

    selected_sheets_label.config(
        text="Selected sheets/tables: " + ", ".join(selected_items)
        if selected_items else "Selected sheets/tables: none"
    )

    selected_x_label.config(
        text=f"X-axis: {selected_x_column.get()}"
        if selected_x_column.get() else "X-axis: -"
    )

    selected_y_label.config(
        text=f"Y-axis: {selected_y_column.get()}"
        if selected_y_column.get() else "Y-axis: -"
    )

    selected_chart_label.config(
        text=f"Chart type: {selected_chart_type.get()}"
        if selected_chart_type.get() else "Chart type: -"
    )

def read_table_file(path):
    ext = os.path.splitext(path)[1].lower()

    if ext in [".xlsx", ".xls", ".ods"]:
        xls_obj = pd.ExcelFile(path)
        local_sheet_map = {name.strip(): name for name in xls_obj.sheet_names}
        local_sheet_names = list(local_sheet_map.keys())
        return "spreadsheet", xls_obj, local_sheet_map, local_sheet_names

    if ext == ".csv":
        df = pd.read_csv(path)
        return "table", df, {"Data": "Data"}, ["Data"]

    if ext == ".tsv":
        df = pd.read_csv(path, sep="\t")
        return "table", df, {"Data": "Data"}, ["Data"]

    if ext == ".txt":
        try:
            df = pd.read_csv(path, sep="\t")
        except Exception:
            df = pd.read_csv(path)
        return "table", df, {"Data": "Data"}, ["Data"]

    raise ValueError("Unsupported file type")

def load_data_file():
    global file_path, xls, sheet_map, sheet_names, all_columns_global, data_mode, table_df

    selected_file = filedialog.askopenfilename(
        title="Select Data File",
        filetypes=[
            ("Supported files", "*.xlsx *.xls *.ods *.csv *.tsv *.txt"),
            ("Excel files", "*.xlsx *.xls"),
            ("LibreOffice files", "*.ods"),
            ("CSV files", "*.csv"),
            ("Text files", "*.txt *.tsv")
        ]
    )

    if not selected_file:
        return

    try:
        file_path = selected_file
        data_mode, loaded_obj, sheet_map, sheet_names = read_table_file(file_path)

        all_columns = []

        if data_mode == "spreadsheet":
            xls = loaded_obj
            table_df = None

            for clean_sheet in sheet_names:
                real_sheet = sheet_map[clean_sheet]
                try:
                    # Headers are expected on the third row for spreadsheets.
                    df_tmp = pd.read_excel(file_path, sheet_name=real_sheet, header=2)
                    df_tmp.columns = [clean_column_name(c) for c in df_tmp.columns]
                    all_columns.extend(df_tmp.columns.tolist())
                except Exception:
                    continue
        else:
            xls = None
            table_df = loaded_obj
            table_df.columns = [clean_column_name(c) for c in table_df.columns]
            all_columns = table_df.columns.tolist()

        unique_columns = []
        seen = set()

        for col in all_columns:
            key = normalize_name(col)
            if key not in seen:
                unique_columns.append(col)
                seen.add(key)

        all_columns_global = unique_columns

        sheet_listbox.delete(0, tk.END)
        for s in sheet_names:
            sheet_listbox.insert(tk.END, s)

        x_combo["values"] = all_columns_global
        y_combo["values"] = all_columns_global

        selected_x_column.set("")
        selected_y_column.set("")

        if all_columns_global:
            time_like = None
            for col in all_columns_global:
                if normalize_name(col) in [
                    normalize_name("Time (sec)"),
                    normalize_name("Time"),
                    normalize_name("Seconds"),
                    normalize_name("sec")
                ]:
                    time_like = col
                    break

            if time_like:
                selected_x_column.set(time_like)
            else:
                x_combo.current(0)

            if len(all_columns_global) > 1:
                for col in all_columns_global:
                    if col != selected_x_column.get():
                        selected_y_column.set(col)
                        break
            else:
                y_combo.current(0)

        file_label.config(text=f"File: {file_path}")
        update_preview()

        messagebox.showinfo(
            "File Loaded",
            f"Loaded {len(sheet_names)} sheet/table item(s) and {len(all_columns_global)} column(s)."
        )

    except Exception as e:
        messagebox.showerror("Error", f"Could not load file:\n{e}")

def find_column(df, wanted_column):
    wanted_key = normalize_name(wanted_column)

    for col in df.columns:
        if normalize_name(col) == wanted_key:
            return col

    return None

def plot_series(ax, x, y, label, color, chart_type, index, total):
    if chart_type == "Smooth Line":
        y_plot = y.rolling(window=5, center=True, min_periods=1).mean()
        ax.plot(
            x,
            y_plot,
            label=label,
            color=color,
            linewidth=2.4,
            solid_capstyle="round",
            solid_joinstyle="round"
        )

    elif chart_type == "Raw Line":
        ax.plot(x, y, label=label, color=color, linewidth=2.1)

    elif chart_type == "Line + Markers":
        ax.plot(
            x,
            y,
            label=label,
            color=color,
            linewidth=2.0,
            marker="o",
            markersize=3,
            markevery=max(1, len(x) // 25)
        )

    elif chart_type == "Scatter":
        ax.scatter(x, y, label=label, color=color, s=18, alpha=0.85)

    elif chart_type == "Bar":
        width = 0.8 / max(total, 1)
        offset = (index - (total - 1) / 2) * width
        ax.bar(x + offset, y, width=width, label=label, color=color, alpha=0.75)

    elif chart_type == "Step":
        ax.step(x, y, label=label, color=color, linewidth=2.1, where="mid")

def make_plot():
    global last_fig

    if not file_path:
        messagebox.showerror("Error", "Please import a data file first.")
        return

    selected_indices = sheet_listbox.curselection()
    chosen_items = [sheet_names[i] for i in selected_indices]
    chosen_x = selected_x_column.get()
    chosen_y = selected_y_column.get()
    chart_type = selected_chart_type.get()

    if not chosen_items:
        messagebox.showerror("Error", "Select at least one sheet/table.")
        return

    if not chosen_x:
        messagebox.showerror("Error", "Select an X-axis column.")
        return

    if not chosen_y:
        messagebox.showerror("Error", "Select a Y-axis column.")
        return

    if not chart_type:
        messagebox.showerror("Error", "Select a chart type.")
        return

    plt.rcParams.update({
        "font.family": "serif",
        "font.size": 12,
        "axes.linewidth": 1.2
    })

    fig, ax = plt.subplots(figsize=(10, 6), dpi=150)
    skipped = []

    for idx, clean_item_name in enumerate(chosen_items):
        try:
            if data_mode == "spreadsheet":
                real_sheet_name = sheet_map[clean_item_name]
                # Headers are expected on the third row for spreadsheets.
                df = pd.read_excel(file_path, sheet_name=real_sheet_name, header=2)
            else:
                df = table_df.copy()

            df.columns = [clean_column_name(c) for c in df.columns]
            df = df.dropna(how="all")

            real_x_col = find_column(df, chosen_x)
            real_y_col = find_column(df, chosen_y)

            if real_x_col is None:
                skipped.append(f"{clean_item_name}: X-axis column not found: {chosen_x}")
                continue

            if real_y_col is None:
                skipped.append(f"{clean_item_name}: Y-axis column not found: {chosen_y}")
                continue

            x = pd.to_numeric(df[real_x_col], errors="coerce")
            y = pd.to_numeric(df[real_y_col], errors="coerce")

            valid = x.notna() & y.notna()
            x = x[valid]
            y = y[valid]

            if len(x) == 0:
                skipped.append(f"{clean_item_name}: no numeric data")
                continue

            label = get_algorithm_label(clean_item_name)
            color = color_map.get(label, "black")

            plot_series(ax, x, y, label, color, chart_type, idx, len(chosen_items))

        except Exception as e:
            skipped.append(f"{clean_item_name}: {e}")

    handles, labels = ax.get_legend_handles_labels()

    if not handles:
        messagebox.showerror(
            "Plot error",
            "Could not plot any selected item.\n\n" + "\n".join(skipped)
        )
        plt.close(fig)
        return

    unique = dict(zip(labels, handles))
    ax.legend(unique.values(), unique.keys(), frameon=True, edgecolor="black")

    ax.set_xlabel(chosen_x)
    ax.set_ylabel(chosen_y)
    ax.set_title(f"{clean_title(chosen_y)} vs {clean_title(chosen_x)}")
    ax.grid(True, linestyle="-", linewidth=0.4, alpha=0.2)

    for spine in ax.spines.values():
        spine.set_linewidth(1.2)

    ax.tick_params(direction="in", length=5, width=1)

    fig.tight_layout()
    last_fig = fig

    if skipped:
        messagebox.showwarning(
            "Some items skipped",
            "Some items were skipped:\n\n" + "\n".join(skipped[:10])
        )

    plt.show()

def save_plot(ext):
    if last_fig is None:
        messagebox.showwarning("Warning", "No plot available.")
        return

    filetypes = {
        "png": [("PNG files", "*.png")],
        "pdf": [("PDF files", "*.pdf")],
        "svg": [("SVG files", "*.svg")]
    }

    path = filedialog.asksaveasfilename(
        defaultextension=f".{ext}",
        filetypes=filetypes[ext]
    )

    if path:
        if ext == "png":
            last_fig.savefig(path, dpi=200, bbox_inches="tight")
        else:
            last_fig.savefig(path, bbox_inches="tight")

        messagebox.showinfo("Saved", f"Plot saved:\n{path}")

def save_png():
    save_plot("png")

def save_pdf():
    save_plot("pdf")

def save_svg():
    save_plot("svg")

def clear_selection():
    sheet_listbox.selection_clear(0, tk.END)
    update_preview()

# ============================================================
# Bindings
# ============================================================

sheet_listbox.bind("<<ListboxSelect>>", update_preview)
x_combo.bind("<<ComboboxSelected>>", update_preview)
y_combo.bind("<<ComboboxSelected>>", update_preview)
chart_type_combo.bind("<<ComboboxSelected>>", update_preview)

# ============================================================
# Buttons
# ============================================================

btn_frame = tk.Frame(root, bg="#f3f6fa")
btn_frame.pack(pady=(4, 14))

ttk.Button(btn_frame, text="Import Data File", command=load_data_file, style="Import.TButton").pack(side="left", padx=7)
ttk.Button(btn_frame, text="Plot", command=make_plot, style="Accent.TButton").pack(side="left", padx=7)
ttk.Button(btn_frame, text="Save PNG", command=save_png, style="Secondary.TButton").pack(side="left", padx=7)
ttk.Button(btn_frame, text="Save PDF", command=save_pdf, style="Secondary.TButton").pack(side="left", padx=7)
ttk.Button(btn_frame, text="Save SVG", command=save_svg, style="Secondary.TButton").pack(side="left", padx=7)
ttk.Button(btn_frame, text="Clear", command=clear_selection, style="Secondary.TButton").pack(side="left", padx=7)

update_preview()
root.mainloop()
