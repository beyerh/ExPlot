import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.ticker import AutoMinorLocator, MultipleLocator, LogLocator, NullLocator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import os
import json
import numpy as np
from scipy import stats

try:
    import pingouin as pg
except ImportError:
    pg = None
try:
    import scikit_posthocs as sp
except ImportError:
    sp = None

DEFAULT_COLORS = {
    "Black": "#000000",
    "Red": "#E74C3C",
    "Blue": "#3498DB",
    "Blueish": "#2B2FFF",
    "Green": "#27AE60",
    "Orange": "#E67E22",
    "Purple": "#8E44AD",
    "Gray": "#7F8C8D"
}
DEFAULT_PALETTES = {
    "Viridis": sns.color_palette("viridis", as_cmap=False),
    "Grayscale": sns.color_palette("gray", as_cmap=False),
    "Set2": sns.color_palette("Set2").as_hex(),
    "Spring Pastels": sns.color_palette("pastel", as_cmap=False),
    "Blue-Black": sns.color_palette(["#2b2fff","#000000"], as_cmap=False),
    "Black-Blue": sns.color_palette(["#000000","#2b2fff"], as_cmap=False)
}

class ExcelPlotterApp:
    def __init__(self, root):
        self.root = root
        self.version = "0.4.0"
        self.root.title('Excel Plotter')
        self.df = None
        self.excel_file = None
        self.preview_label = None
        self.temp_pdf = "temp_plot.pdf"
        self.custom_colors_file = "custom_colors.json"
        self.custom_palettes_file = "custom_palettes.json"
        self.xaxis_renames = {}
        self.xaxis_order = []
        self.linewidth = tk.DoubleVar(value=1.0)
        self.strip_black_var = tk.BooleanVar(value=True)
        self.show_stripplot_var = tk.BooleanVar(value=True)
        self.plot_kind_var = tk.StringVar(value="bar")  # "bar", "box", or "xy"
        # --- XY plot option variables (ensure these are initialized before setup_ui) ---
        self.xy_marker_symbol_var = tk.StringVar(value="o")
        self.xy_marker_size_var = tk.DoubleVar(value=5)
        self.xy_filled_var = tk.BooleanVar(value=True)
        self.xy_line_style_var = tk.StringVar(value="solid")
        self.xy_line_black_var = tk.BooleanVar(value=False)
        self.xy_connect_var = tk.BooleanVar(value=False)
        self.xy_show_mean_var = tk.BooleanVar(value=True)
        self.xy_show_mean_errorbars_var = tk.BooleanVar(value=True)
        self.xy_draw_band_var = tk.BooleanVar(value=False)
        self.load_custom_colors_palettes()

        self.setup_menu()
        self.setup_ui()

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)

    def show_about(self):
        messagebox.showinfo("About Excel Plotter", f"Excel Plotter\nVersion: {self.version}\n\nA tool for plotting Excel data.")

    def load_custom_colors_palettes(self):
        if os.path.exists(self.custom_colors_file):
            with open(self.custom_colors_file, 'r') as f:
                self.custom_colors = json.load(f)
        else:
            self.custom_colors = DEFAULT_COLORS.copy()
        if os.path.exists(self.custom_palettes_file):
            with open(self.custom_palettes_file, 'r') as f:
                loaded = json.load(f)
                # Convert all palettes to hex if not already
                self.custom_palettes = {k: [self._to_hex(c) for c in v] for k, v in loaded.items()}
        else:
            self.custom_palettes = {k: [self._to_hex(c) for c in v] for k, v in DEFAULT_PALETTES.items()}

    def save_custom_colors_palettes(self):
        with open(self.custom_colors_file, 'w') as f:
            json.dump(self.custom_colors, f, indent=2)
        # Always save palettes as hex
        palettes_hex = {k: [self._to_hex(c) for c in v] for k, v in self.custom_palettes.items()}
        with open(self.custom_palettes_file, 'w') as f:
            json.dump(palettes_hex, f, indent=2)

    def _to_hex(self, color):
        # Accepts hex, tuple, or list
        if isinstance(color, str) and color.startswith('#'):
            return color
        if isinstance(color, (tuple, list)) and len(color) == 3:
            # If already 0-1 floats, convert to 0-255
            if all(isinstance(x, float) and 0 <= x <= 1 for x in color):
                color = tuple(int(x * 255) for x in color)
            return '#{:02x}{:02x}{:02x}'.format(*color)
        return str(color)

    def setup_ui(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill='both', expand=True)

        left_frame = tk.Frame(main_frame)
        left_frame.pack(side='left', fill='both', expand=False)

        self.tab_control = ttk.Notebook(left_frame)
        self.tab_control.pack(fill='both', expand=True)

        self.basic_tab = tk.Frame(self.tab_control)
        self.appearance_tab = tk.Frame(self.tab_control)
        self.axis_tab = tk.Frame(self.tab_control)
        self.colors_tab = tk.Frame(self.tab_control)

        self.tab_control.add(self.basic_tab, text="Basic Settings")
        self.tab_control.add(self.appearance_tab, text="Appearance")
        self.tab_control.add(self.axis_tab, text="Axis Settings")
        self.tab_control.add(self.colors_tab, text="Colors")

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side='right', fill='both', expand=True)

        self.canvas_frame = tk.Frame(right_frame)
        self.canvas_frame.pack(fill='both', expand=True)

        tk.Button(right_frame, text='Generate Plot', command=self.plot_graph).pack(pady=5)
        tk.Button(right_frame, text='Save as PDF', command=self.save_pdf).pack(pady=5)

        self.setup_basic_tab()
        self.setup_appearance_tab()
        self.setup_axis_tab()
        self.setup_colors_tab()

    def setup_basic_tab(self):
        frame = self.basic_tab
        # File/Sheet group
        file_grp = tk.LabelFrame(frame, text="File/Sheet", padx=6, pady=6)
        file_grp.pack(fill='x', padx=6, pady=(8,4))
        tk.Button(file_grp, text='Load Excel File', command=self.load_file).pack(fill='x', pady=2)
        tk.Label(file_grp, text="Sheet:").pack(anchor="w")
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(file_grp, textvariable=self.sheet_var, width=18)
        self.sheet_dropdown.pack(fill='x', pady=2)
        self.sheet_dropdown.bind('<<ComboboxSelected>>', self.load_sheet)
        # Columns group
        col_grp = tk.LabelFrame(frame, text="Columns", padx=6, pady=6)
        col_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(col_grp, text="X-axis column:").grid(row=0, column=0, sticky="w", pady=2)
        self.xaxis_var = tk.StringVar()
        self.xaxis_dropdown = ttk.Combobox(col_grp, textvariable=self.xaxis_var, width=18)
        self.xaxis_dropdown.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(col_grp, text="Group column:").grid(row=1, column=0, sticky="w", pady=2)
        self.group_var = tk.StringVar()
        self.group_dropdown = ttk.Combobox(col_grp, textvariable=self.group_var, width=18)
        self.group_dropdown.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(col_grp, text="Y-axis columns:").grid(row=2, column=0, sticky="nw", pady=(4,0))
        # Scrollable frame for value checkbuttons
        value_vars_scroll_frame = tk.Frame(col_grp)
        value_vars_scroll_frame.grid(row=2, column=1, sticky="ew", pady=(2,0))
        value_vars_canvas = tk.Canvas(value_vars_scroll_frame, height=120)
        value_vars_scrollbar = tk.Scrollbar(value_vars_scroll_frame, orient="vertical", command=value_vars_canvas.yview)
        self.value_vars_inner_frame = tk.Frame(value_vars_canvas)
        self.value_vars_inner_frame.bind(
            "<Configure>", lambda e: value_vars_canvas.configure(scrollregion=value_vars_canvas.bbox("all")))
        value_vars_canvas.create_window((0, 0), window=self.value_vars_inner_frame, anchor="nw")
        value_vars_canvas.configure(yscrollcommand=value_vars_scrollbar.set)
        value_vars_canvas.pack(side="left", fill="both", expand=True)
        value_vars_scrollbar.pack(side="right", fill="y")
        self.value_vars = []
        self.value_checkbuttons = []
        col_grp.columnconfigure(1, weight=1)
        # Options group
        opt_grp = tk.LabelFrame(frame, text="Options", padx=6, pady=6)
        opt_grp.pack(fill='x', padx=6, pady=4)
        self.split_yaxis = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_grp, text="Split Y-axis if multiple", variable=self.split_yaxis).pack(anchor="w", pady=1)
        self.use_stats_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_grp, text="Use statistics", variable=self.use_stats_var).pack(anchor="w", pady=1)
        # --- Error bar type (SD/SEM) ---
        self.errorbar_type_var = tk.StringVar(value="SD")
        errorbar_frame = tk.Frame(opt_grp)
        errorbar_frame.pack(anchor="w", pady=2)
        tk.Label(errorbar_frame, text="Error bars:").pack(side="left")
        tk.Radiobutton(errorbar_frame, text="SD", variable=self.errorbar_type_var, value="SD").pack(side="left")
        tk.Radiobutton(errorbar_frame, text="SEM", variable=self.errorbar_type_var, value="SEM").pack(side="left")
        # Black errorbars option
        self.errorbar_black_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opt_grp, text="Black errorbars", variable=self.errorbar_black_var).pack(anchor="w", pady=1)
        btn_fr = tk.Frame(opt_grp)
        btn_fr.pack(fill='x', pady=(2,0))
        tk.Button(btn_fr, text="Rename X labels", command=self.rename_xaxis_labels, width=14).pack(side="left", padx=1)
        tk.Button(btn_fr, text="Reorder X categories", command=self.reorder_xaxis_categories, width=16).pack(side="left", padx=1)
        # --- Plot type group (with horizontal layout for XY options, no tabs) ---
        type_grp = tk.LabelFrame(frame, text="Plot Type", padx=6, pady=6)
        type_grp.pack(fill='x', padx=6, pady=4)
        # Frame for radio buttons (left)
        plot_type_radio_frame = tk.Frame(type_grp)
        plot_type_radio_frame.grid(row=0, column=0, sticky="nw")
        tk.Label(plot_type_radio_frame, text="Plot Type:").pack(anchor="w")
        bar_radio = tk.Radiobutton(plot_type_radio_frame, text="Bar Graph", variable=self.plot_kind_var, value="bar")
        box_radio = tk.Radiobutton(plot_type_radio_frame, text="Box Plot", variable=self.plot_kind_var, value="box")
        xy_radio = tk.Radiobutton(plot_type_radio_frame, text="XY Plot", variable=self.plot_kind_var, value="xy")
        bar_radio.pack(anchor="w")
        box_radio.pack(anchor="w")
        xy_radio.pack(anchor="w")
        # XY options frame (right)
        self.xy_options_frame = tk.Frame(type_grp)
        # XY options widgets in xy_options_frame
        self.xy_marker_symbol_label = tk.Label(self.xy_options_frame, text="XY Marker Symbol:")
        self.xy_marker_symbol_dropdown = ttk.Combobox(self.xy_options_frame, textvariable=self.xy_marker_symbol_var, values=["o", "s", "^", "D", "v", "P", "X", "+", "x", "*", "."], width=10)
        self.xy_marker_size_label = tk.Label(self.xy_options_frame, text="XY Marker Size:")
        self.xy_marker_size_entry = tk.Entry(self.xy_options_frame, textvariable=self.xy_marker_size_var, width=6)
        self.xy_filled_check = tk.Checkbutton(self.xy_options_frame, text="Filled symbols", variable=self.xy_filled_var)
        self.xy_line_style_label = tk.Label(self.xy_options_frame, text="Line style:")
        self.xy_line_style_dropdown = ttk.Combobox(self.xy_options_frame, textvariable=self.xy_line_style_var, values=["solid", "dashed", "dotted", "dashdot"], width=10)
        self.xy_line_black_check = tk.Checkbutton(self.xy_options_frame, text="Lines in black", variable=self.xy_line_black_var)
        self.xy_connect_check = tk.Checkbutton(self.xy_options_frame, text="Connect mean with lines", variable=self.xy_connect_var)
        self.xy_show_mean_check = tk.Checkbutton(self.xy_options_frame, text="Show mean values", variable=self.xy_show_mean_var, command=self.update_xy_mean_errorbar_state)
        self.xy_show_mean_errorbars_check = tk.Checkbutton(self.xy_options_frame, text="With errorbars", variable=self.xy_show_mean_errorbars_var)
        self.xy_draw_band_check = tk.Checkbutton(self.xy_options_frame, text="Draw bands (min-max or error)", variable=self.xy_draw_band_var)
        # Pack XY widgets in order
        self.xy_marker_symbol_label.grid(row=0, column=0, sticky="w", padx=8, pady=1)
        self.xy_marker_symbol_dropdown.grid(row=0, column=1, sticky="w", padx=2, pady=1)
        self.xy_marker_size_label.grid(row=1, column=0, sticky="w", padx=8, pady=1)
        self.xy_marker_size_entry.grid(row=1, column=1, sticky="w", padx=2, pady=1)
        self.xy_filled_check.grid(row=2, column=0, columnspan=2, sticky="w", padx=8, pady=1)
        self.xy_line_style_label.grid(row=3, column=0, sticky="w", padx=8, pady=1)
        self.xy_line_style_dropdown.grid(row=3, column=1, sticky="w", padx=2, pady=1)
        self.xy_line_black_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=1)
        self.xy_connect_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=8, pady=1)
        self.xy_show_mean_check.grid(row=6, column=0, columnspan=2, sticky="w", padx=8, pady=1)
        self.xy_show_mean_errorbars_check.grid(row=7, column=0, columnspan=2, sticky="w", padx=24, pady=1)
        self.xy_draw_band_check.grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=1)
        # Show/hide XY options frame based on plot type
        def update_xy_options(*args):
            if self.plot_kind_var.get() == "xy":
                self.xy_options_frame.grid(row=0, column=1, sticky="nw", padx=(16,0))
                self.show_stripplot_var.set(False)
                self.update_xy_mean_errorbar_state()
            elif self.plot_kind_var.get() == "bar":
                self.show_stripplot_var.set(True)
                self.xy_options_frame.grid_forget()
            else:
                self.xy_options_frame.grid_forget()
            
        self.plot_kind_var.trace_add('write', update_xy_options)
        update_xy_options()

    def update_xy_mean_errorbar_state(self):
        if self.xy_show_mean_var.get():
            self.xy_show_mean_errorbars_check.config(state='normal')
        else:
            self.xy_show_mean_errorbars_var.set(False)
            self.xy_show_mean_errorbars_check.config(state='disabled')

    def setup_appearance_tab(self):
        frame = self.appearance_tab
        # --- Size group ---
        size_grp = tk.LabelFrame(frame, text="Figure Size", padx=6, pady=6)
        size_grp.pack(fill='x', padx=6, pady=(8,4))
        tk.Label(size_grp, text="Plot Width (inches):").grid(row=0, column=0, sticky="w", pady=2)
        self.width_entry = tk.Entry(size_grp)
        self.width_entry.insert(0, "1.5")
        self.width_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(size_grp, text="Plot Height per plot (inches):").grid(row=1, column=0, sticky="w", pady=2)
        self.height_entry = tk.Entry(size_grp)
        self.height_entry.insert(0, "1.5")
        self.height_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        size_grp.columnconfigure(1, weight=1)
        # --- Font/Line group ---
        font_grp = tk.LabelFrame(frame, text="Font & Line", padx=6, pady=6)
        font_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(font_grp, text="Font Size:").grid(row=0, column=0, sticky="w", pady=2)
        self.fontsize_entry = tk.Entry(font_grp)
        self.fontsize_entry.insert(0, "10")
        self.fontsize_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(font_grp, text="Line Width:").grid(row=1, column=0, sticky="w", pady=2)
        self.linewidth_entry = tk.Entry(font_grp, textvariable=self.linewidth)
        self.linewidth_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        font_grp.columnconfigure(1, weight=1)
        # --- Stripplot group ---
        strip_grp = tk.LabelFrame(frame, text="Stripplot", padx=6, pady=6)
        strip_grp.pack(fill='x', padx=6, pady=4)
        self.strip_black_var = tk.BooleanVar(value=True)
        tk.Checkbutton(strip_grp, text="Show stripplot with black dots", variable=self.strip_black_var).pack(anchor="w", pady=1)
        self.show_stripplot_var = tk.BooleanVar(value=True)
        tk.Checkbutton(strip_grp, text="Show stripplot", variable=self.show_stripplot_var).pack(anchor="w", pady=1)
        # --- Display options group ---
        disp_grp = tk.LabelFrame(frame, text="Display Options", padx=6, pady=6)
        disp_grp.pack(fill='x', padx=6, pady=4)
        self.show_frame_var = tk.BooleanVar(value=False)
        tk.Checkbutton(disp_grp, text="Show graph frame", variable=self.show_frame_var).pack(anchor="w", pady=1)
        self.show_hgrid_var = tk.BooleanVar(value=False)
        tk.Checkbutton(disp_grp, text="Show horizontal grid", variable=self.show_hgrid_var).pack(anchor="w", pady=1)
        self.show_vgrid_var = tk.BooleanVar(value=False)
        tk.Checkbutton(disp_grp, text="Show vertical grid", variable=self.show_vgrid_var).pack(anchor="w", pady=1)
        self.swap_axes_var = tk.BooleanVar(value=False)
        tk.Checkbutton(disp_grp, text="Swap X and Y axes", variable=self.swap_axes_var).pack(anchor="w", pady=1)

    def setup_axis_tab(self):
        frame = self.axis_tab
        # --- X Label group ---
        xlabel_grp = tk.LabelFrame(frame, text="X-axis Label", padx=6, pady=6)
        xlabel_grp.pack(fill='x', padx=6, pady=(8,4))
        tk.Label(xlabel_grp, text="X-axis Label:").pack(anchor="w")
        self.xlabel_entry = tk.Entry(xlabel_grp)
        self.xlabel_entry.pack(fill='x', pady=2)
        # --- Y Label group ---
        ylabel_grp = tk.LabelFrame(frame, text="Y-axis Label", padx=6, pady=6)
        ylabel_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(ylabel_grp, text="Y-axis Label:").pack(anchor="w")
        self.ylabel_entry = tk.Entry(ylabel_grp)
        self.ylabel_entry.pack(fill='x', pady=2)
        # --- Label orientation group ---
        orient_grp = tk.LabelFrame(frame, text="X-axis Label Orientation", padx=6, pady=6)
        orient_grp.pack(fill='x', padx=6, pady=4)
        self.label_orientation = tk.StringVar(value="vertical")
        tk.Radiobutton(orient_grp, text="Vertical", variable=self.label_orientation, value="vertical").pack(anchor="w")
        tk.Radiobutton(orient_grp, text="Horizontal", variable=self.label_orientation, value="horizontal").pack(anchor="w")
        # --- X Axis settings group ---
        xaxis_grp = tk.LabelFrame(frame, text="X-Axis Settings", padx=6, pady=6)
        xaxis_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(xaxis_grp, text="Minimum:").grid(row=0, column=0, sticky="w", pady=2)
        self.xmin_entry = tk.Entry(xaxis_grp)
        self.xmin_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(xaxis_grp, text="Maximum:").grid(row=1, column=0, sticky="w", pady=2)
        self.xmax_entry = tk.Entry(xaxis_grp)
        self.xmax_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(xaxis_grp, text="Major Tick Interval:").grid(row=2, column=0, sticky="w", pady=2)
        self.xinterval_entry = tk.Entry(xaxis_grp)
        self.xinterval_entry.grid(row=2, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(xaxis_grp, text="Minor ticks per major:").grid(row=3, column=0, sticky="w", pady=2)
        self.xminor_ticks_entry = tk.Entry(xaxis_grp)
        self.xminor_ticks_entry.grid(row=3, column=1, sticky="ew", padx=2, pady=2)
        xaxis_grp.columnconfigure(1, weight=1)
        # --- Y Axis settings group ---
        yaxis_grp = tk.LabelFrame(frame, text="Y-Axis Settings", padx=6, pady=6)
        yaxis_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(yaxis_grp, text="Minimum:").grid(row=0, column=0, sticky="w", pady=2)
        self.ymin_entry = tk.Entry(yaxis_grp)
        self.ymin_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(yaxis_grp, text="Maximum:").grid(row=1, column=0, sticky="w", pady=2)
        self.ymax_entry = tk.Entry(yaxis_grp)
        self.ymax_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(yaxis_grp, text="Major Tick Interval:").grid(row=2, column=0, sticky="w", pady=2)
        self.yinterval_entry = tk.Entry(yaxis_grp)
        self.yinterval_entry.grid(row=2, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(yaxis_grp, text="Minor ticks per major:").grid(row=3, column=0, sticky="w", pady=2)
        self.minor_ticks_entry = tk.Entry(yaxis_grp)
        self.minor_ticks_entry.grid(row=3, column=1, sticky="ew", padx=2, pady=2)
        self.logscale_var = tk.BooleanVar()
        tk.Checkbutton(yaxis_grp, text="Logarithmic Y-axis", variable=self.logscale_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4,1))
        yaxis_grp.columnconfigure(1, weight=1)

    def setup_colors_tab(self):
        frame = self.colors_tab
        # --- Color management group ---
        color_mgmt_grp = tk.LabelFrame(frame, text="Color Management", padx=6, pady=6)
        color_mgmt_grp.pack(fill='x', padx=6, pady=(8,4))
        tk.Button(color_mgmt_grp, text="Manage Colors & Palettes", command=self.manage_colors_palettes).pack(fill='x', pady=2)
        # --- Single color group ---
        single_grp = tk.LabelFrame(frame, text="Single Data Color", padx=6, pady=6)
        single_grp.pack(fill='x', padx=6, pady=4)
        self.single_color_var = tk.StringVar(value=list(self.custom_colors.keys())[0])
        tk.Label(single_grp, text="Single Data Color:").pack(anchor="w")
        self.single_color_dropdown = ttk.Combobox(single_grp, textvariable=self.single_color_var, values=list(self.custom_colors.keys()))
        self.single_color_dropdown.pack(fill='x', pady=2)
        # Add preview canvas for single color
        self.single_color_preview = tk.Canvas(single_grp, width=60, height=20, bg='white', highlightthickness=0)
        self.single_color_preview.pack(pady=(0, 8))
        def update_single_color_preview(event=None):
            self.single_color_preview.delete('all')
            name = self.single_color_var.get()
            hexcode = self.custom_colors.get(name)
            if hexcode:
                self.single_color_preview.create_rectangle(10, 2, 50, 18, fill=hexcode, outline='black')
        self.single_color_dropdown.bind('<<ComboboxSelected>>', update_single_color_preview)
        update_single_color_preview()
        # --- Palette group ---
        palette_grp = tk.LabelFrame(frame, text="Group Palette", padx=6, pady=6)
        palette_grp.pack(fill='x', padx=6, pady=4)
        self.palette_var = tk.StringVar(value=list(self.custom_palettes.keys())[0])
        tk.Label(palette_grp, text="Group Palette:").pack(anchor="w")
        self.palette_dropdown = ttk.Combobox(palette_grp, textvariable=self.palette_var, values=list(self.custom_palettes.keys()))
        self.palette_dropdown.pack(fill='x', pady=2)
        # Add preview canvas for palette
        self.palette_preview = tk.Canvas(palette_grp, width=120, height=20, bg='white', highlightthickness=0)
        self.palette_preview.pack(pady=(0, 8))
        def update_palette_preview(event=None):
            self.palette_preview.delete('all')
            name = self.palette_var.get()
            colors = self.custom_palettes.get(name, [])
            for i, hexcode in enumerate(colors[:8]):
                x0 = 10 + i*14
                x1 = x0 + 12
                self.palette_preview.create_rectangle(x0, 2, x1, 18, fill=hexcode, outline='black')
        self.palette_dropdown.bind('<<ComboboxSelected>>', update_palette_preview)
        update_palette_preview()

    def update_color_palette_dropdowns(self):
        self.single_color_dropdown['values'] = list(self.custom_colors.keys())
        if self.single_color_var.get() not in self.custom_colors and self.custom_colors:
            self.single_color_var.set(list(self.custom_colors.keys())[0])
        self.palette_dropdown['values'] = list(self.custom_palettes.keys())
        if self.palette_var.get() not in self.custom_palettes and self.custom_palettes:
            self.palette_var.set(list(self.custom_palettes.keys())[0])

    def create_dropdown(self, parent, label_text, attr_name):
        tk.Label(parent, text=label_text).pack()
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var)
        dropdown.pack()
        setattr(self, f"{attr_name}_var", var)
        setattr(self, f"{attr_name}_dropdown", dropdown)

    def add_labeled_entry(self, parent, label):
        tk.Label(parent, text=label).pack()
        entry = tk.Entry(parent)
        entry.pack()
        return entry

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_file = file_path
            xls = pd.ExcelFile(self.excel_file)
            self.sheet_dropdown['values'] = xls.sheet_names

            if "export" in xls.sheet_names:
                self.sheet_var.set("export")
            else:
                self.sheet_var.set(xls.sheet_names[0])

            self.load_sheet()

    def load_sheet(self, event=None):
        try:
            self.df = pd.read_excel(self.excel_file, sheet_name=self.sheet_var.get(), dtype=object)
            self.update_columns()
            self.xaxis_order = []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {e}")

    def update_columns(self):
        columns = list(self.df.columns)
        self.xaxis_dropdown['values'] = columns
        self.group_dropdown['values'] = ['None'] + columns

        for cb in self.value_checkbuttons:
            cb.destroy()
        self.value_vars.clear()

        for col in columns:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(self.value_vars_inner_frame, text=col, variable=var)
            cb.pack(anchor='w')
            self.value_vars.append((var, col))
            self.value_checkbuttons.append(cb)

    def rename_xaxis_labels(self):
        if self.df is None or not self.xaxis_var.get():
            messagebox.showerror("Error", "Load a file and select an X-axis first.")
            return

        # Use current order if set, otherwise unique values from DataFrame
        if self.xaxis_order:
            current_values = self.xaxis_order
        else:
            current_values = list(pd.unique(self.df[self.xaxis_var.get()].dropna()))

        rename_window = tk.Toplevel(self.root)
        rename_window.title("Rename X-axis Labels")

        entries = {}
        for val in current_values:
            frame = tk.Frame(rename_window)
            frame.pack()
            tk.Label(frame, text=str(val)).pack(side='left')
            entry = tk.Entry(frame)
            # Pre-fill with previous rename if exists, else the value itself
            entry.insert(0, self.xaxis_renames.get(val, str(val)))
            entry.pack(side='left')
            entries[val] = entry

        def save_renames():
            self.xaxis_renames = {k: e.get() for k, e in entries.items()}
            # Update the order to the new renamed values
            self.xaxis_order = [self.xaxis_renames.get(k, k) for k in current_values]
            rename_window.destroy()

        tk.Button(rename_window, text="Save", command=save_renames).pack()

    def reorder_xaxis_categories(self):
        if self.df is None or not self.xaxis_var.get():
            messagebox.showerror("Error", "Load a file and select an X-axis first.")
            return

        # Use renamed values if available
        orig_values = list(pd.unique(self.df[self.xaxis_var.get()].dropna()))
        display_values = [self.xaxis_renames.get(val, val) for val in orig_values]

        order_window = tk.Toplevel(self.root)
        order_window.title("Reorder X-axis Categories")

        listbox = tk.Listbox(order_window, width=40)
        for val in display_values:
            listbox.insert(tk.END, val)
        listbox.pack()

        def move_up():
            i = listbox.curselection()
            if i and i[0] > 0:
                val = listbox.get(i)
                listbox.delete(i)
                listbox.insert(i[0] - 1, val)
                listbox.select_set(i[0] - 1)

        def move_down():
            i = listbox.curselection()
            if i and i[0] < listbox.size() - 1:
                val = listbox.get(i)
                listbox.delete(i)
                listbox.insert(i[0] + 1, val)
                listbox.select_set(i[0] + 1)

        tk.Button(order_window, text="Up", command=move_up).pack()
        tk.Button(order_window, text="Down", command=move_down).pack()

        def save_order():
            # Save the order as the displayed (renamed) values
            self.xaxis_order = list(listbox.get(0, tk.END))
            order_window.destroy()

        tk.Button(order_window, text="Save Order", command=save_order).pack()

    def plot_graph(self):
        if self.df is None:
            return

        try:
            linewidth = float(self.linewidth.get())
        except Exception:
            linewidth = 1.0

        x_col = self.xaxis_var.get()
        group_col = self.group_var.get()
        if not group_col or group_col.strip() == '' or group_col == 'None':
            group_col = None

        value_cols = [col for var, col in self.value_vars if var.get() and col != x_col]
        if not x_col or not value_cols:
            messagebox.showerror("Error", "Select X-axis and at least one value column.")
            return

        plot_mode = 'single' if len(value_cols) == 1 else 'overlay'
        plot_width = float(self.width_entry.get())
        plot_height = float(self.height_entry.get())
        fontsize = int(self.fontsize_entry.get())
        n_rows = len(value_cols) if (self.split_yaxis.get() and plot_mode == 'single') else 1

        left_margin = 1.0
        right_margin = 0.5
        top_margin = 1.5  # for legend
        bottom_margin = 1.0 + fontsize * 0.1

        fig_width = plot_width + left_margin + right_margin
        fig_height = plot_height * n_rows + top_margin + bottom_margin

        # Set all possible linewidths via rcParams for errorbars, axes, ticks, grid, legend
        matplotlib.rcParams.update({
            'pdf.fonttype': 42,
            'font.size': fontsize,
            'axes.linewidth': linewidth,
            'xtick.major.width': linewidth,
            'ytick.major.width': linewidth,
            'xtick.minor.width': linewidth,
            'ytick.minor.width': linewidth,
            'grid.linewidth': linewidth,
            'legend.edgecolor': 'inherit'
        })

        sns.set_theme(style='white', rc={"axes.grid": False})
        fig, axes = plt.subplots(n_rows, 1, figsize=(fig_width, fig_height), squeeze=False)
        axes = axes.flatten()

        show_frame = self.show_frame_var.get()
        show_hgrid = self.show_hgrid_var.get()
        show_vgrid = self.show_vgrid_var.get()
        swap_axes = self.swap_axes_var.get()
        strip_black = self.strip_black_var.get()
        show_stripplot = self.show_stripplot_var.get()
        plot_kind = self.plot_kind_var.get()  # "bar", "box", or "xy"
        errorbar_black = self.errorbar_black_var.get()

        for idx, value_col in enumerate(value_cols):
            ax = axes[idx] if n_rows > 1 else axes[0]
            df_plot = self.df.copy()

            if self.xaxis_renames:
                df_plot[x_col] = df_plot[x_col].map(self.xaxis_renames).fillna(df_plot[x_col])

            if self.xaxis_order:
                df_plot[x_col] = pd.Categorical(df_plot[x_col], categories=self.xaxis_order, ordered=True)

            if plot_mode == 'overlay' and len(value_cols) > 1:
                df_plot = pd.melt(df_plot, id_vars=[x_col] + ([group_col] if group_col else []),
                                  value_vars=value_cols, var_name='Measurement', value_name='Value')
                hue_col = 'Measurement'
            else:
                df_plot['Value'] = df_plot[value_col]
                hue_col = group_col if group_col else None

            # --- Palette assignment: always use full palette for groups ---
            if hue_col:
                palette_name = self.palette_var.get()
                palette = self.custom_palettes.get(palette_name, "Set2")
                # Ensure palette is a list of colors and long enough
                n_hue = len(df_plot[hue_col].dropna().unique())
                if isinstance(palette, list):
                    if len(palette) < n_hue:
                        palette = (palette * ((n_hue // len(palette)) + 1))[:n_hue]
                else:
                    palette = [palette] * n_hue
            else:
                single_color_name = self.single_color_var.get()
                palette = [self.custom_colors.get(single_color_name, 'black')]

            # --- XY Plot: use one color per group from palette ---
            if plot_kind == "xy" and hue_col:
                n_hue = len(df_plot[hue_col].unique())
                if isinstance(palette, list):
                    if len(palette) < n_hue:
                        palette = "black"
                        messagebox.showwarning("Palette Error", "Palette does not have enough colors for all groups. Using black instead.")
                    else:
                        # Use only as many colors as needed
                        palette = palette[:n_hue]

            # --- Error bar type logic ---
            errorbar_type = self.errorbar_type_var.get()
            sem_mode = False
            if plot_kind == "bar":
                if errorbar_type == "SD":
                    ci_val = 'sd'
                    estimator = np.mean
                elif errorbar_type == "SEM":
                    ci_val = None
                    estimator = np.mean
                    sem_mode = True
                else:
                    ci_val = 'sd'
                    estimator = np.mean

            # --- Swap axes logic ---
            if swap_axes:
                plot_args = dict(
                    data=df_plot, y=x_col, x='Value', hue=hue_col, ax=ax,
                )
                if plot_kind == "bar":
                    plot_args.update(dict(ci=ci_val, capsize=0.2, palette=palette, errcolor='black', errwidth=linewidth, estimator=estimator))
                elif plot_kind == "box":
                    plot_args.update(dict(palette=palette, linewidth=linewidth, showcaps=True, boxprops=dict(linewidth=linewidth), medianprops=dict(linewidth=linewidth), dodge=True, width=0.7))
                elif plot_kind == "xy":
                    marker_size = self.xy_marker_size_var.get()
                    marker_symbol = self.xy_marker_symbol_var.get()
                    connect = self.xy_connect_var.get()
                    draw_band = self.xy_draw_band_var.get()
                    show_mean = self.xy_show_mean_var.get()
                    show_mean_errorbars = self.xy_show_mean_errorbars_var.get()
                    filled = self.xy_filled_var.get()
                    line_style = self.xy_line_style_var.get()
                    line_black = self.xy_line_black_var.get()
                    # Choose color logic for XY plot
                    if len(value_cols) == 1:
                        color = self.custom_colors.get(self.single_color_var.get(), 'black')
                        palette = [color]
                    else:
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        palette = palette_full[:len(value_cols)]
                    # --- Always use full palette for grouped XY means ---
                    if show_mean:
                        groupers = [x_col]
                        if hue_col:
                            groupers.append(hue_col)
                        grouped = df_plot.groupby(groupers)['Value']
                        means = grouped.mean().reset_index()
                        if self.errorbar_type_var.get() == "SEM":
                            errors = grouped.apply(lambda x: np.std(x.dropna().astype(float), ddof=1) / np.sqrt(len(x.dropna())) if len(x.dropna()) > 1 else 0).reset_index(name='err')
                        else:
                            errors = grouped.std(ddof=1).reset_index(name='err')
                        means = means.merge(errors, on=groupers)
                        if hue_col:
                            group_names = list(df_plot[hue_col].dropna().unique())
                            palette_name = self.palette_var.get()
                            palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                            if len(palette_full) < len(group_names):
                                palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                            color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                            for name in group_names:
                                group = means[means[hue_col] == name]
                                if group.empty:
                                    continue
                                c = color_map[name]
                                x = group[x_col]
                                y = group['Value']
                                yerr = group['err']
                                ecolor = 'black' if errorbar_black else c
                                mfc = c if filled else 'none'
                                mec = c
                                # Draw errorbars only if requested
                                if show_mean_errorbars:
                                    ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor, label=str(name))
                                # Always plot the mean points
                                ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None', label=None if show_mean_errorbars else str(name))
                                # Always draw error bands if requested (independent of errorbars)
                                if draw_band:
                                    df_band = pd.DataFrame({
                                        'x': pd.to_numeric(x, errors='coerce'),
                                        'y': pd.to_numeric(y, errors='coerce'),
                                        'yerr': pd.to_numeric(yerr, errors='coerce')
                                    }).dropna().sort_values('x')
                                    if not df_band.empty:
                                        ax.fill_between(
                                            df_band['x'],
                                            df_band['y'] - df_band['yerr'],
                                            df_band['y'] + df_band['yerr'],
                                            color=c, alpha=0.18, zorder=1
                                        )
                                if connect:
                                    ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        else:
                            c = palette[0]
                            x = means[x_col]
                            y = means['Value']
                            yerr = means['err']
                            ecolor = 'black' if errorbar_black else c
                            mfc = c if filled else 'none'
                            mec = c
                            if show_mean_errorbars:
                                ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor)
                            ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None')
                            if draw_band:
                                df_band = pd.DataFrame({
                                    'x': pd.to_numeric(x, errors='coerce'),
                                    'y': pd.to_numeric(y, errors='coerce'),
                                    'yerr': pd.to_numeric(yerr, errors='coerce')
                                }).dropna().sort_values('x')
                                if not df_band.empty:
                                    ax.fill_between(
                                        df_band['x'],
                                        df_band['y'] - df_band['yerr'],
                                        df_band['y'] + df_band['yerr'],
                                        color=c, alpha=0.18, zorder=1
                                    )
                            if connect:
                                ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        if hue_col:
                            ax.legend()
                    else:
                        # Plot raw data points (scatter) when show_mean is False
                        marker_kwargs = dict(marker=marker_symbol, s=marker_size**2)
                        if hue_col:
                            group_names = list(df_plot[hue_col].dropna().unique())
                            palette_name = self.palette_var.get()
                            palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                            if len(palette_full) < len(group_names):
                                palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                            color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                            for name in group_names:
                                group = df_plot[df_plot[hue_col] == name]
                                if group.empty:
                                    continue
                                c = color_map[name]
                                scatter = ax.scatter(group[x_col], group['Value'], marker=marker_symbol, s=marker_size**2, color=c, label=str(name), edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                                if draw_band:
                                    x_sorted = np.sort(group[x_col].unique())
                                    min_vals = [group[group[x_col] == x]['Value'].min() for x in x_sorted]
                                    max_vals = [group[group[x_col] == x]['Value'].max() for x in x_sorted]
                                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                    min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                    max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                    ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                                if connect:
                                    # Connect means of raw data at each x value
                                    if hue_col:
                                        for name in group_names:
                                            group = df_plot[df_plot[hue_col] == name]
                                            if group.empty:
                                                continue
                                            c = color_map[name]
                                            # Calculate means at each x value
                                            x_sorted = np.sort(group[x_col].unique())
                                            means = [group[group[x_col] == x]['Value'].mean() for x in x_sorted]
                                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                            means_numeric = pd.to_numeric(means, errors='coerce')
                                            ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                                    else:
                                        c = palette[0]
                                        x_sorted = np.sort(df_plot[x_col].unique())
                                        means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                        means_numeric = pd.to_numeric(means, errors='coerce')
                                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                            ax.legend()
                        else:
                            c = palette[0]
                            ax.scatter(df_plot[x_col], df_plot['Value'], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                            if draw_band:
                                x_sorted = np.sort(df_plot[x_col].unique())
                                min_vals = [df_plot[df_plot[x_col] == x]['Value'].min() for x in x_sorted]
                                max_vals = [df_plot[df_plot[x_col] == x]['Value'].max() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                            if connect:
                                # Connect means of raw data at each x value
                                c = palette[0]
                                x_sorted = np.sort(df_plot[x_col].unique())
                                means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                means_numeric = pd.to_numeric(means, errors='coerce')
                                ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                stripplot_args = dict(
                    data=df_plot, y=x_col, x='Value', hue=hue_col, dodge=True,
                    jitter=True, marker='o', alpha=0.55,
                    ax=ax
                )
            else:
                plot_args = dict(
                    data=df_plot, x=x_col, y='Value', hue=hue_col, ax=ax,
                )
                # --- Bar plot: always horizontal bars if swap_axes, for both categorical and numerical x ---
                if plot_kind == "bar":
                    if swap_axes:
                        plot_args = dict(
                            data=df_plot, y=x_col, x='Value', hue=hue_col, ax=ax,
                            ci=ci_val, capsize=0.2, palette=palette, errcolor='black', errwidth=linewidth, estimator=estimator
                        )
                    else:
                        plot_args = dict(
                            data=df_plot, x=x_col, y='Value', hue=hue_col, ax=ax,
                            ci=ci_val, capsize=0.2, palette=palette, errcolor='black', errwidth=linewidth, estimator=estimator
                        )
                elif plot_kind == "box":
                    plot_args.update(dict(palette=palette, linewidth=linewidth, showcaps=True, boxprops=dict(linewidth=linewidth), medianprops=dict(linewidth=linewidth), dodge=True, width=0.7))
                elif plot_kind == "xy":
                    marker_size = self.xy_marker_size_var.get()
                    marker_symbol = self.xy_marker_symbol_var.get()
                    connect = self.xy_connect_var.get()
                    draw_band = self.xy_draw_band_var.get()
                    show_mean = self.xy_show_mean_var.get()
                    show_mean_errorbars = self.xy_show_mean_errorbars_var.get()
                    filled = self.xy_filled_var.get()
                    line_style = self.xy_line_style_var.get()
                    line_black = self.xy_line_black_var.get()
                    # Choose color logic for XY plot
                    if len(value_cols) == 1:
                        color = self.custom_colors.get(self.single_color_var.get(), 'black')
                        palette = [color]
                    else:
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        palette = palette_full[:len(value_cols)]
                    # --- Always use full palette for grouped XY means ---
                    if show_mean:
                        groupers = [x_col]
                        if hue_col:
                            groupers.append(hue_col)
                        grouped = df_plot.groupby(groupers)['Value']
                        means = grouped.mean().reset_index()
                        if self.errorbar_type_var.get() == "SEM":
                            errors = grouped.apply(lambda x: np.std(x.dropna().astype(float), ddof=1) / np.sqrt(len(x.dropna())) if len(x.dropna()) > 1 else 0).reset_index(name='err')
                        else:
                            errors = grouped.std(ddof=1).reset_index(name='err')
                        means = means.merge(errors, on=groupers)
                        if hue_col:
                            group_names = list(df_plot[hue_col].dropna().unique())
                            palette_name = self.palette_var.get()
                            palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                            if len(palette_full) < len(group_names):
                                palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                            color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                            for name in group_names:
                                group = means[means[hue_col] == name]
                                if group.empty:
                                    continue
                                c = color_map[name]
                                x = group[x_col]
                                y = group['Value']
                                yerr = group['err']
                                ecolor = 'black' if errorbar_black else c
                                mfc = c if filled else 'none'
                                mec = c
                                if show_mean_errorbars:
                                    ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor, label=str(name))
                                # Always plot the mean points
                                ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None', label=None if show_mean_errorbars else str(name))
                                # Always draw error bands if requested (independent of errorbars)
                                if draw_band:
                                    x_numeric = pd.to_numeric(x, errors='coerce')
                                    y_numeric = pd.to_numeric(y, errors='coerce')
                                    yerr_numeric = pd.to_numeric(yerr, errors='coerce')
                                    ax.fill_between(x_numeric, y_numeric - yerr_numeric, y_numeric + yerr_numeric, color=c, alpha=0.18, zorder=1)
                                if connect:
                                    ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        else:
                            c = palette[0]
                            x = means[x_col]
                            y = means['Value']
                            yerr = means['err']
                            ecolor = 'black' if errorbar_black else c
                            mfc = c if filled else 'none'
                            mec = c
                            if show_mean_errorbars:
                                ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor, label=str(name))
                            # Always plot the mean points
                            ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None', label=None if show_mean_errorbars else str(name))
                            # Always draw error bands if requested (independent of errorbars)
                            if draw_band:
                                x_numeric = pd.to_numeric(x, errors='coerce')
                                y_numeric = pd.to_numeric(y, errors='coerce')
                                yerr_numeric = pd.to_numeric(yerr, errors='coerce')
                                ax.fill_between(x_numeric, y_numeric - yerr_numeric, y_numeric + yerr_numeric, color=c, alpha=0.18, zorder=1)
                            if connect:
                                ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        if hue_col:
                            ax.legend()
                    else:
                        # Plot raw data points (scatter) when show_mean is False
                        marker_kwargs = dict(marker=marker_symbol, s=marker_size**2)
                        if hue_col:
                            group_names = list(df_plot[hue_col].dropna().unique())
                            palette_name = self.palette_var.get()
                            palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                            if len(palette_full) < len(group_names):
                                palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                            color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                            for name in group_names:
                                group = df_plot[df_plot[hue_col] == name]
                                if group.empty:
                                    continue
                                c = color_map[name]
                                scatter = ax.scatter(group[x_col], group['Value'], marker=marker_symbol, s=marker_size**2, color=c, label=str(name), edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                                if draw_band:
                                    x_sorted = np.sort(group[x_col].unique())
                                    min_vals = [group[group[x_col] == x]['Value'].min() for x in x_sorted]
                                    max_vals = [group[group[x_col] == x]['Value'].max() for x in x_sorted]
                                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                    min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                    max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                    ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                                if connect:
                                    # Connect means of raw data at each x value
                                    if hue_col:
                                        for name in group_names:
                                            group = df_plot[df_plot[hue_col] == name]
                                            if group.empty:
                                                continue
                                            c = color_map[name]
                                            # Calculate means at each x value
                                            x_sorted = np.sort(group[x_col].unique())
                                            means = [group[group[x_col] == x]['Value'].mean() for x in x_sorted]
                                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                            means_numeric = pd.to_numeric(means, errors='coerce')
                                            ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                                    else:
                                        c = palette[0]
                                        x_sorted = np.sort(df_plot[x_col].unique())
                                        means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                        means_numeric = pd.to_numeric(means, errors='coerce')
                                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                            ax.legend()
                        else:
                            c = palette[0]
                            ax.scatter(df_plot[x_col], df_plot['Value'], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                            if draw_band:
                                x_sorted = np.sort(df_plot[x_col].unique())
                                min_vals = [df_plot[df_plot[x_col] == x]['Value'].min() for x in x_sorted]
                                max_vals = [df_plot[df_plot[x_col] == x]['Value'].max() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                            if connect:
                                # Connect means of raw data at each x value
                                c = palette[0]
                                x_sorted = np.sort(df_plot[x_col].unique())
                                means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                means_numeric = pd.to_numeric(means, errors='coerce')
                                ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                stripplot_args = dict(
                    data=df_plot, x=x_col, y='Value', hue=hue_col, dodge=True,
                    jitter=True, marker='o', alpha=0.55,
                    ax=ax
                )

            # --- Plotting ---
            if plot_kind == "bar":
                sns.barplot(**plot_args)
                # --- Add SEM error bars manually if needed ---
                if sem_mode:
                    # Determine grouping columns
                    groupers = [x_col]
                    if hue_col:
                        groupers.append(hue_col)
                    grouped = df_plot.groupby(groupers)['Value']
                    means = grouped.mean()
                    sems = grouped.apply(lambda x: np.std(x.dropna().astype(float), ddof=1) / np.sqrt(len(x.dropna())) if len(x.dropna()) > 1 else 0)
                    # Get bar positions
                    if swap_axes:
                        # y is category, x is value
                        for i, (bar, mean) in enumerate(zip(ax.patches, means)):
                            # bar: Rectangle
                            y = bar.get_y() + bar.get_height() / 2
                            x = bar.get_x() + bar.get_width() / 2
                            sem = sems.iloc[i]
                            ax.errorbar(x, y, xerr=sem, fmt='none', ecolor='black', elinewidth=linewidth, capsize=3, zorder=10)
                    else:
                        for i, (bar, mean) in enumerate(zip(ax.patches, means)):
                            x = bar.get_x() + bar.get_width() / 2
                            y = bar.get_height() + bar.get_y()
                            sem = sems.iloc[i]
                            ax.errorbar(x, y, yerr=sem, fmt='none', ecolor='black', elinewidth=linewidth, capsize=3, zorder=10)
            elif plot_kind == "box":
                sns.boxplot(**plot_args)
                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)
            elif plot_kind == "xy":
                marker_size = self.xy_marker_size_var.get()
                marker_symbol = self.xy_marker_symbol_var.get()
                connect = self.xy_connect_var.get()
                draw_band = self.xy_draw_band_var.get()
                show_mean = self.xy_show_mean_var.get()
                show_mean_errorbars = self.xy_show_mean_errorbars_var.get()
                filled = self.xy_filled_var.get()
                line_style = self.xy_line_style_var.get()
                line_black = self.xy_line_black_var.get()
                # Choose color logic for XY plot
                if len(value_cols) == 1:
                    color = self.custom_colors.get(self.single_color_var.get(), 'black')
                    palette = [color]
                else:
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    palette = palette_full[:len(value_cols)]
                # --- Always use full palette for grouped XY means ---
                if show_mean:
                    groupers = [x_col]
                    if hue_col:
                        groupers.append(hue_col)
                    grouped = df_plot.groupby(groupers)['Value']
                    means = grouped.mean().reset_index()
                    if self.errorbar_type_var.get() == "SEM":
                        errors = grouped.apply(lambda x: np.std(x.dropna().astype(float), ddof=1) / np.sqrt(len(x.dropna())) if len(x.dropna()) > 1 else 0).reset_index(name='err')
                    else:
                        errors = grouped.std(ddof=1).reset_index(name='err')
                    means = means.merge(errors, on=groupers)
                    if hue_col:
                        group_names = list(df_plot[hue_col].dropna().unique())
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        if len(palette_full) < len(group_names):
                            palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                        color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                        for name in group_names:
                            group = means[means[hue_col] == name]
                            if group.empty:
                                continue
                            c = color_map[name]
                            x = group[x_col]
                            y = group['Value']
                            yerr = group['err']
                            ecolor = 'black' if errorbar_black else c
                            mfc = c if filled else 'none'
                            mec = c
                            if show_mean_errorbars:
                                ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor, label=str(name))
                                if draw_band:
                                    x_numeric = pd.to_numeric(x, errors='coerce')
                                    y_numeric = pd.to_numeric(y, errors='coerce')
                                    yerr_numeric = pd.to_numeric(yerr, errors='coerce')
                                    ax.fill_between(x_numeric, y_numeric - yerr_numeric, y_numeric + yerr_numeric, color=c, alpha=0.18, zorder=1)
                            else:
                                ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None', label=str(name))
                            if connect:
                                ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                    else:
                        c = palette[0]
                        x = means[x_col]
                        y = means['Value']
                        yerr = means['err']
                        ecolor = 'black' if errorbar_black else c
                        mfc = c if filled else 'none'
                        mec = c
                        if show_mean_errorbars:
                            ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor)
                            if draw_band:
                                x_numeric = pd.to_numeric(x, errors='coerce')
                                y_numeric = pd.to_numeric(y, errors='coerce')
                                yerr_numeric = pd.to_numeric(yerr, errors='coerce')
                                ax.fill_between(x_numeric, y_numeric - yerr_numeric, y_numeric + yerr_numeric, color=c, alpha=0.18, zorder=1)
                        else:
                            ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None')
                        if connect:
                            ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                    if hue_col:
                        ax.legend()
                else:
                    # Plot raw data points (scatter) when show_mean is False
                    marker_kwargs = dict(marker=marker_symbol, s=marker_size**2)
                    if hue_col:
                        group_names = list(df_plot[hue_col].dropna().unique())
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        if len(palette_full) < len(group_names):
                            palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                        color_map = {name: palette_full[i] for i, name in enumerate(group_names)}
                        for name in group_names:
                            group = df_plot[df_plot[hue_col] == name]
                            if group.empty:
                                continue
                            c = color_map[name]
                            scatter = ax.scatter(group[x_col], group['Value'], marker=marker_symbol, s=marker_size**2, color=c, label=str(name), edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                            if draw_band:
                                x_sorted = np.sort(group[x_col].unique())
                                min_vals = [group[group[x_col] == x]['Value'].min() for x in x_sorted]
                                max_vals = [group[group[x_col] == x]['Value'].max() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                            if connect:
                                # Connect means of raw data at each x value
                                if hue_col:
                                    for name in group_names:
                                        group = df_plot[df_plot[hue_col] == name]
                                        if group.empty:
                                            continue
                                        c = color_map[name]
                                        # Calculate means at each x value
                                        x_sorted = np.sort(group[x_col].unique())
                                        means = [group[group[x_col] == x]['Value'].mean() for x in x_sorted]
                                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                        means_numeric = pd.to_numeric(means, errors='coerce')
                                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                                else:
                                    c = palette[0]
                                    x_sorted = np.sort(df_plot[x_col].unique())
                                    means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                    means_numeric = pd.to_numeric(means, errors='coerce')
                                    ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        ax.legend()
                    else:
                        c = palette[0]
                        ax.scatter(df_plot[x_col], df_plot['Value'], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=c if filled else 'none', facecolors=c if filled else 'none', linewidth=linewidth)
                        if draw_band:
                            x_sorted = np.sort(df_plot[x_col].unique())
                            min_vals = [df_plot[df_plot[x_col] == x]['Value'].min() for x in x_sorted]
                            max_vals = [df_plot[df_plot[x_col] == x]['Value'].max() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                            max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                            ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                        if connect:
                            # Connect means of raw data at each x value
                            c = palette[0]
                            x_sorted = np.sort(df_plot[x_col].unique())
                            means = [df_plot[df_plot[x_col] == x]['Value'].mean() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            means_numeric = pd.to_numeric(means, errors='coerce')
                            ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)

            # --- Stripplot (if enabled) ---
            if show_stripplot:
                if strip_black:
                    stripplot_args["palette"] = ["black"]
                    stripplot_args["color"] = "black"
                else:
                    if hue_col:
                        stripplot_args["palette"] = palette
                    else:
                        stripplot_args["palette"] = palette
                # Suppress legend for stripplot
                stripplot_args["legend"] = False
                sns.stripplot(**stripplot_args)

            # --- Always rebuild legend after all plotting ---
            if hue_col and plot_kind == "box":
                import matplotlib.patches as mpatches
                hue_levels = list(df_plot[hue_col].dropna().unique())
                palette_name = self.palette_var.get()
                palette = self.custom_palettes.get(palette_name, "Set2")
                if len(palette) < len(hue_levels):
                    palette = (palette * ((len(hue_levels) // len(palette)) + 1))[:len(hue_levels)]
                handles = [mpatches.Patch(facecolor=palette[i], edgecolor='black', label=str(hue_levels[i])) for i in range(len(hue_levels))]
                ax.legend(
                    handles,
                    [str(l) for l in hue_levels],
                    loc="upper center", bbox_to_anchor=(0.5, 1.18), borderaxespad=0,
                    frameon=False, fontsize=fontsize, ncol=max(1, len(handles))
                )
            elif hue_col and plot_kind == "bar":
                # Only use bar handles (Rectangle, not PathCollection)
                handles, labels = ax.get_legend_handles_labels()
                from matplotlib.patches import Rectangle
                bar_handles = [h for h in handles if isinstance(h, Rectangle) and h.get_height() != 0]
                bar_labels = [l for h, l in zip(handles, labels) if isinstance(h, Rectangle) and h.get_height() != 0]
                if not bar_handles:  # fallback: use all handles
                    bar_handles, bar_labels = handles, labels
                ax.legend(
                    bar_handles,
                    bar_labels,
                    loc="upper center", bbox_to_anchor=(0.5, 1.18), borderaxespad=0,
                    frameon=False, fontsize=fontsize, ncol=max(1, len(bar_handles))
                )
            elif hue_col and plot_kind == "xy":
                handles, labels = ax.get_legend_handles_labels()
                by_label = dict(zip(labels, handles))
                ax.legend(
                    by_label.values(), by_label.keys(),
                    loc="upper center", bbox_to_anchor=(0.5, 1.18), borderaxespad=0,
                    frameon=False, fontsize=fontsize, ncol=max(1, len(by_label))
                )

            # --- Grids ---
            if show_hgrid:
                ax.grid(which='major', axis='y' if not swap_axes else 'x', linestyle='-', color='gray', alpha=0.5, visible=True, linewidth=linewidth)
            if show_vgrid:
                ax.grid(which='major', axis='x' if not swap_axes else 'y', linestyle='-', color='gray', alpha=0.5, visible=True, linewidth=linewidth)

            # --- Frame (spines) ---
            for spine in ax.spines.values():
                spine.set_visible(show_frame)
                spine.set_linewidth(linewidth)
            for side in ['bottom', 'left']:
                if not show_frame:
                    ax.spines[side].set_visible(True)
                    ax.spines[side].set_linewidth(linewidth)

            # --- Axis labels ---
            if swap_axes:
                ax.set_ylabel(self.xlabel_entry.get() or x_col, fontsize=fontsize)
                ax.set_xlabel(value_col if n_rows > 1 else (self.ylabel_entry.get() or value_col), fontsize=fontsize)
            else:
                ax.set_xlabel(self.xlabel_entry.get() or x_col, fontsize=fontsize)
                ax.set_ylabel(value_col if n_rows > 1 else (self.ylabel_entry.get() or value_col), fontsize=fontsize)

            rotation = 90 if self.label_orientation.get() == 'vertical' and not swap_axes else 0
            ax.tick_params(axis='x', labelsize=fontsize, rotation=rotation, direction='in', length=4, width=linewidth)
            ax.tick_params(axis='y', labelsize=fontsize, direction='in', length=4, width=linewidth, color='black', left=True)

            # --- The rest of your axis/tick/statistics/annotation code ---
            try:
                ymin = float(self.ymin_entry.get()) if self.ymin_entry.get() else None
                ymax = float(self.ymax_entry.get()) if self.ymax_entry.get() else None
                yinterval = float(self.yinterval_entry.get()) if self.yinterval_entry.get() else None
                minor_ticks_str = self.minor_ticks_entry.get()
                minor_ticks = int(minor_ticks_str) if minor_ticks_str else None

                xmin = float(self.xmin_entry.get()) if self.xmin_entry.get() else None
                xmax = float(self.xmax_entry.get()) if self.xmax_entry.get() else None
                xinterval = float(self.xinterval_entry.get()) if self.xinterval_entry.get() else None
                xminor_ticks_str = self.xminor_ticks_entry.get()
                xminor_ticks = int(xminor_ticks_str) if xminor_ticks_str else None

                if not swap_axes:
                    # Y-axis settings
                    if ymin is not None or ymax is not None:
                        ax.set_ylim(bottom=ymin, top=ymax)
                    if yinterval:
                        ax.yaxis.set_major_locator(MultipleLocator(yinterval))
                    if minor_ticks:
                        ax.yaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                        ax.tick_params(axis='y', which='minor', direction='in', length=2, width=linewidth, color='black', left=True)
                    else:
                        ax.yaxis.set_minor_locator(NullLocator())
                        ax.tick_params(axis='y', which='minor', length=0)
                    # X-axis settings (for XY plots)
                    if xmin is not None or xmax is not None:
                        ax.set_xlim(left=xmin, right=xmax)
                    if xinterval:
                        ax.xaxis.set_major_locator(MultipleLocator(xinterval))
                    if xminor_ticks:
                        ax.xaxis.set_minor_locator(AutoMinorLocator(xminor_ticks + 1))
                        ax.tick_params(axis='x', which='minor', direction='in', length=2, width=linewidth, color='black', bottom=True)
                    else:
                        ax.xaxis.set_minor_locator(NullLocator())
                        ax.tick_params(axis='x', which='minor', length=0)
                else:
                    # X-axis settings (swapped)
                    if ymin is not None or ymax is not None:
                        ax.set_xlim(left=ymin, right=ymax)
                    if yinterval:
                        ax.xaxis.set_major_locator(MultipleLocator(yinterval))
                    if minor_ticks:
                        ax.xaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                        ax.tick_params(axis='x', which='minor', direction='in', length=2, width=linewidth, color='black', bottom=True)
                    else:
                        ax.xaxis.set_minor_locator(NullLocator())
                        ax.tick_params(axis='x', which='minor', length=0)
                    # Y-axis settings (swapped, for XY plots)
                    if xmin is not None or xmax is not None:
                        ax.set_ylim(bottom=xmin, top=xmax)
                    if xinterval:
                        ax.yaxis.set_major_locator(MultipleLocator(xinterval))
                    if xminor_ticks:
                        ax.yaxis.set_minor_locator(AutoMinorLocator(xminor_ticks + 1))
                        ax.tick_params(axis='y', which='minor', direction='in', length=2, width=linewidth, color='black', left=True)
                    else:
                        ax.yaxis.set_minor_locator(NullLocator())
                        ax.tick_params(axis='y', which='minor', length=0)
            except Exception as e:
                print(f"Axis setting error: {e}")

            use_log = self.logscale_var.get()
            if use_log:
                if not swap_axes:
                    ax.set_yscale('log')
                else:
                    ax.set_xscale('log')

            minor_ticks_str = self.minor_ticks_entry.get()
            if minor_ticks_str:
                try:
                    minor_ticks = int(minor_ticks_str)
                    if use_log:
                        subs = np.logspace(0, 1, minor_ticks + 2)[1:-1]
                        if not swap_axes:
                            ax.yaxis.set_minor_locator(LogLocator(base=10.0, subs=subs, numticks=100))
                        else:
                            ax.xaxis.set_minor_locator(LogLocator(base=10.0, subs=subs, numticks=100))
                    else:
                        if not swap_axes:
                            ax.yaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                        else:
                            ax.xaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                    ax.tick_params(
                        axis='y' if not swap_axes else 'x', which='minor', direction='in',
                        length=2, width=linewidth, color='black', left=True
                    )
                except Exception as e:
                    print(f"Minor ticks setting error: {e}")
            else:
                if not swap_axes:
                    ax.yaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='y', which='minor', length=0)
                else:
                    ax.xaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='x', which='minor', length=0)

            ax_position = [left_margin / fig_width,
                           bottom_margin / fig_height,
                           plot_width / fig_width,
                           plot_height / fig_height]
            ax.set_position(ax_position)

            # --- Statistics/Annotations ---
            if self.use_stats_var.get():
                try:
                    import itertools
                    y_max = df_plot['Value'].astype(float).max()
                    y_min = df_plot['Value'].astype(float).min()
                    pval_height = y_max + 0.07 * (y_max - y_min if y_max != y_min else 1)
                    step = 0.06 * (y_max - y_min if y_max != y_min else 1)
                    test_used = ""
                    if not hue_col:
                        groups = [g for g in df_plot[x_col].dropna().unique()]
                        positions = {g: i for i, g in enumerate(groups)}
                        pairs = list(itertools.combinations(groups, 2))
                        if pg is not None and sp is not None and len(groups) > 2:
                            test_used = "Welch ANOVA + Tamhane's T2"
                            posthoc = sp.posthoc_tamhane(df_plot, val_col='Value', group_col=x_col)
                            for idx, (g1, g2) in enumerate(pairs):
                                try:
                                    pval = posthoc.loc[g1, g2] if g2 in posthoc.columns and g1 in posthoc.index else posthoc.loc[g2, g1]
                                except Exception:
                                    pval = np.nan
                                x1, x2 = positions[g1], positions[g2]
                                y = pval_height + idx * step
                                # annotation lines
                                if swap_axes:
                                    ax.plot([y-0.01, y, y, y-0.01], [x1, x1, x2, x2], lw=linewidth, c='k', zorder=10)
                                    ax.text(y+0.01, (x1+x2)/2, f"p={pval:.2e}", ha='left', va='center', fontsize=max(int(fontsize*0.8), 6), zorder=11)
                                else:
                                    ax.plot([x1, x1, x2, x2], [y-0.01, y, y, y-0.01], lw=linewidth, c='k', zorder=10)
                                    ax.text((x1+x2)/2, y+0.01, f"p={pval:.2e}", ha='center', va='bottom', fontsize=max(int(fontsize*0.8), 6), zorder=11)
                        elif len(groups) == 2:
                            test_used = "Welch's t-test (two-sided, unequal variances)"
                            vals0 = df_plot[df_plot[x_col] == groups[0]]['Value'].astype(float)
                            vals1 = df_plot[df_plot[x_col] == groups[1]]['Value'].astype(float)
                            ttest = stats.ttest_ind(vals0, vals1, equal_var=False, alternative='two-sided')
                            x1, x2 = positions[groups[0]], positions[groups[1]]
                            y = pval_height
                            if swap_axes:
                                ax.plot([y-0.01, y, y, y-0.01], [x1, x1, x2, x2], lw=linewidth, c='k', zorder=10)
                                ax.text(y+0.01, (x1+x2)/2, f"p={ttest.pvalue:.2e}", ha='left', va='center', fontsize=max(int(fontsize*0.8), 6), zorder=11)
                            else:
                                ax.plot([x1, x1, x2, x2], [y-0.01, y, y, y-0.01], lw=linewidth, c='k', zorder=10)
                                ax.text((x1+x2)/2, y+0.01, f"p={ttest.pvalue:.2e}", ha='center', va='bottom', fontsize=max(int(fontsize*0.8), 6), zorder=11)
                        else:
                            test_used = "No valid test"
                    else:
                        base_groups = [g for g in df_plot[x_col].dropna().unique()]
                        hue_groups = [g for g in df_plot[hue_col].dropna().unique()]
                        n_hue = len(hue_groups)
                        if len(hue_groups) > 2 and pg is not None and sp is not None:
                            test_used = "Welch ANOVA + Tamhane's T2"
                        elif len(hue_groups) == 2:
                            test_used = "Welch's t-test (two-sided, unequal variances)"
                        else:
                            test_used = "No valid test"
                        for i, g in enumerate(base_groups):
                            df_sub = df_plot[df_plot[x_col] == g]
                            pairs = list(itertools.combinations(hue_groups, 2))
                            x_base = i
                            for idx, (h1, h2) in enumerate(pairs):
                                vals1 = df_sub[df_sub[hue_col] == h1]['Value'].astype(float)
                                vals2 = df_sub[df_sub[hue_col] == h2]['Value'].astype(float)
                                if len(hue_groups) == 2:
                                    ttest = stats.ttest_ind(vals1, vals2, equal_var=False, alternative='two-sided')
                                    pval = ttest.pvalue
                                elif pg is not None and sp is not None and len(vals1) > 0 and len(vals2) > 0:
                                    try:
                                        posthoc = sp.posthoc_tamhane(df_sub, val_col='Value', group_col=hue_col)
                                        pval = posthoc.loc[h1, h2] if h2 in posthoc.columns and h1 in posthoc.index else posthoc.loc[h2, h1]
                                    except Exception:
                                        pval = np.nan
                                else:
                                    continue
                                bar_centers = []
                                for j, hue_val in enumerate(hue_groups):
                                    bar_pos = x_base - 0.4 + (j + 0.5) * (0.8 / n_hue)
                                    bar_centers.append(bar_pos)
                                x1 = bar_centers[hue_groups.index(h1)]
                                x2 = bar_centers[hue_groups.index(h2)]
                                y = pval_height + idx * step + i * step * len(pairs)
                                if swap_axes:
                                    ax.plot([y-0.01, y, y, y-0.01], [x1, x1, x2, x2], lw=linewidth, c='k', zorder=10)
                                    ax.text(y+0.01, (x1+x2)/2, f"p={pval:.2e}", ha='left', va='center', fontsize=max(int(fontsize*0.7), 6), zorder=11)
                                else:
                                    ax.plot([x1, x1, x2, x2], [y-0.01, y, y, y-0.01], lw=linewidth, c='k', zorder=10)
                                    ax.text((x1+x2)/2, y+0.01, f"p={pval:.2e}", ha='center', va='bottom', fontsize=max(int(fontsize*0.7), 6), zorder=11)
                    fig = ax.figure
                    fig.text(0.01, 0.99, f"Statistical test: {test_used}", ha='left', va='top', fontsize=max(int(fontsize*0.9), 8), color='darkslategray')
                except Exception as e:
                    ax.text(0.01, 0.98, f"Stats error: {e}", transform=ax.transAxes, fontsize=fontsize*0.8, va='top', ha='left',
                            bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.8))
            # Reset all linewidths (defensive: after annotations/statistics)
            for spine in ax.spines.values():
                spine.set_linewidth(linewidth)
            ax.tick_params(axis='both', which='both', width=linewidth)
            # (grid lines already set above)

        fig.savefig(self.temp_pdf, format='pdf', bbox_inches='tight')
        self.display_preview()

    def display_preview(self):
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()
        pages = convert_from_path(self.temp_pdf, dpi=150)
        image = pages[0]
        photo = ImageTk.PhotoImage(image)
        self.preview_label = tk.Label(self.canvas_frame, image=photo)
        self.preview_label.image = photo
        self.preview_label.pack()

    def save_pdf(self):
        if os.path.exists(self.temp_pdf):
            file_path = filedialog.asksaveasfilename(defaultextension='.pdf', filetypes=[("PDF files", "*.pdf")], initialfile='plot_output.pdf')
            if file_path:
                os.replace(self.temp_pdf, file_path)

    def manage_colors_palettes(self):
        window = tk.Toplevel(self.root)
        window.title("Manage Colors & Palettes")

        tk.Label(window, text="Single Data Colors:").pack(pady=(10,2))
        colors_frame = tk.Frame(window)
        colors_frame.pack()

        color_listbox = tk.Listbox(colors_frame, height=7)
        color_listbox.pack(side='left')
        color_scrollbar = tk.Scrollbar(colors_frame)
        color_scrollbar.pack(side='right', fill='y')
        color_listbox.config(yscrollcommand=color_scrollbar.set)
        color_scrollbar.config(command=color_listbox.yview)

        # Color preview canvas
        color_preview = tk.Canvas(window, width=60, height=20, bg='white', highlightthickness=0)
        color_preview.pack(pady=(0, 8))

        def show_color_preview(event=None):
            color_preview.delete('all')
            sel = color_listbox.curselection()
            if sel:
                name = color_listbox.get(sel[0]).split(":")[0].strip()
                hexcode = self.custom_colors.get(name)
                if hexcode:
                    color_preview.create_rectangle(10, 2, 50, 18, fill=hexcode, outline='black')
        color_listbox.bind('<<ListboxSelect>>', show_color_preview)

        def refresh_color_list():
            color_listbox.delete(0, tk.END)
            for name, val in self.custom_colors.items():
                color_listbox.insert(tk.END, f"{name}: {val}")
            show_color_preview()
        refresh_color_list()

        def add_color():
            def save_color():
                name = name_entry.get().strip()
                if not name:
                    messagebox.showerror("Error", "Color name required.")
                    return
                color_code = colorchooser.askcolor()[1]
                if not color_code:
                    messagebox.showerror("Error", "No color selected.")
                    return
                self.custom_colors[name] = color_code
                refresh_color_list()
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                popup.destroy()
            popup = tk.Toplevel(window)
            tk.Label(popup, text="Color Name:").pack()
            name_entry = tk.Entry(popup)
            name_entry.pack()
            tk.Button(popup, text="Pick Color and Save", command=save_color).pack()
        tk.Button(window, text="Add Color", command=add_color).pack(pady=2)

        def remove_color():
            selected = color_listbox.curselection()
            if not selected:
                return
            name = color_listbox.get(selected[0]).split(":")[0].strip()
            if name in self.custom_colors:
                del self.custom_colors[name]
                refresh_color_list()
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
        tk.Button(window, text="Remove Selected Color", command=remove_color).pack(pady=2)

        tk.Label(window, text="Palettes:").pack(pady=(15,2))
        palettes_frame = tk.Frame(window)
        palettes_frame.pack()

        palette_listbox = tk.Listbox(palettes_frame, height=7)
        palette_listbox.pack(side='left')
        palette_scrollbar = tk.Scrollbar(palettes_frame)
        palette_scrollbar.pack(side='right', fill='y')
        palette_listbox.config(yscrollcommand=palette_scrollbar.set)
        palette_scrollbar.config(command=palette_listbox.yview)

        # Palette preview canvas
        palette_preview = tk.Canvas(window, width=120, height=20, bg='white', highlightthickness=0)
        palette_preview.pack(pady=(0, 8))

        def show_palette_preview(event=None):
            palette_preview.delete('all')
            sel = palette_listbox.curselection()
            if sel:
                name = palette_listbox.get(sel[0]).split(":")[0].strip()
                colors = self.custom_palettes.get(name, [])
                for i, hexcode in enumerate(colors[:8]):
                    x0 = 10 + i*14
                    x1 = x0 + 12
                    palette_preview.create_rectangle(x0, 2, x1, 18, fill=hexcode, outline='black')
        palette_listbox.bind('<<ListboxSelect>>', show_palette_preview)

        def refresh_palette_list():
            palette_listbox.delete(0, tk.END)
            for name, val in self.custom_palettes.items():
                preview = ', '.join([self._to_hex(c) for c in val[:5]])
                palette_listbox.insert(tk.END, f"{name}: {preview}...")
            show_palette_preview()
        refresh_palette_list()

        def add_palette():
            def save_palette():
                name = name_entry.get().strip()
                colors = colors_entry.get().strip().split(",")
                if not name or not colors:
                    messagebox.showerror("Error", "Palette name and color list required.")
                    return
                colors = [self._to_hex(c.strip()) for c in colors if c.strip()]
                self.custom_palettes[name] = colors
                refresh_palette_list()
                self.save_custom_palettes_palettes()
                self.update_color_palette_dropdowns()
                popup.destroy()
            popup = tk.Toplevel(window)
            tk.Label(popup, text="Palette Name:").pack()
            name_entry = tk.Entry(popup)
            name_entry.pack()
            tk.Label(popup, text="Colors (comma separated hex codes):").pack()
            colors_entry = tk.Entry(popup)
            colors_entry.pack()
            tk.Button(popup, text="Save Palette", command=save_palette).pack()
        tk.Button(window, text="Add Palette", command=add_palette).pack(pady=2)

        def remove_palette():
            selected = palette_listbox.curselection()
            if not selected:
                return
            name = palette_listbox.get(selected[0]).split(":")[0].strip()
            if name in self.custom_palettes:
                del self.custom_palettes[name]
                refresh_palette_list()
                self.save_custom_palettes_palettes()
                self.update_color_palette_dropdowns()
        tk.Button(window, text="Remove Selected Palette", command=remove_palette).pack(pady=2)

        tk.Button(window, text="Close", command=window.destroy).pack(pady=10)

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelPlotterApp(root)
    root.mainloop()