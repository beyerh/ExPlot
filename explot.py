# ExPlot - Data visualization tool for Excel files

VERSION = "0.7.2"
# =====================================================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import pandas as pd
import seaborn as sns
from pathlib import Path
import os
os.environ['MPLCONFIGDIR'] = str(Path.home())+"/.matplotlib/"
import matplotlib
import matplotlib.pyplot as plt
import matplotlib
from matplotlib.ticker import AutoMinorLocator, MultipleLocator, LogLocator, NullLocator, FixedLocator, FuncFormatter
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.path import Path as MplPath
from matplotlib.lines import Line2D
from matplotlib.transforms import blended_transform_factory
from PIL import Image, ImageTk
import json
import time
import numpy as np
import copy
from scipy import stats
import sys
import tempfile
import pingouin as pg
import scikit_posthocs as sp
import math
from scipy.optimize import curve_fit
import warnings
import traceback  # For better error reporting
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from tkinter import font as tkfont
from statannotations.Annotator import Annotator

# Disable dask imports, nuitka fix for macOS
'''import sys
sys.modules['dask'] = None
sys.modules['dask.array'] = None
sys.modules['dask.dataframe'] = None
sys.modules['dask.array.core'] = None
sys.modules['dask.array.compat'] = None
'''

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
    "Viridis": sns.color_palette("viridis", as_cmap=False).as_hex(),
    "Grayscale": sns.color_palette("gray", as_cmap=False).as_hex(),
    "Set2": sns.color_palette("Set2").as_hex(),
    "Spring Pastels": sns.color_palette("pastel", as_cmap=False).as_hex(),
    "Blue-Black": sns.color_palette(["#2b2fff","#000000"], as_cmap=False).as_hex(),
    "Black-Blue": sns.color_palette(["#000000","#2b2fff"], as_cmap=False).as_hex(),
    
    # New palettes added
    "Ocean": sns.color_palette("deep", as_cmap=False).as_hex(),
    "Colorblind": sns.color_palette("colorblind", as_cmap=False).as_hex(),
    "Muted": sns.color_palette("muted", as_cmap=False).as_hex(),
    "Bright": sns.color_palette("bright", as_cmap=False).as_hex(),
    "Dark": sns.color_palette("dark", as_cmap=False).as_hex(),
    "Plasma": sns.color_palette("plasma", as_cmap=False).as_hex(),
    "Inferno": sns.color_palette("inferno", as_cmap=False).as_hex(),
    "Paired": sns.color_palette("Paired", as_cmap=False).as_hex(),
    
    # Custom scientific publication-friendly palettes
    "Publication": ["#4878D0", "#EE854A", "#6ACC64", "#D65F5F", "#956CB4", "#8C613C", "#DC7EC0", "#797979"],
    "Nature": ["#1F77B4", "#FF7F0E", "#2CA02C", "#D62728", "#9467BD", "#8C564B", "#E377C2", "#7F7F7F"],
    "Scientific": ["#3182bd", "#e6550d", "#31a354", "#756bb1", "#636363", "#6baed6", "#fd8d3c", "#74c476"],
    
    # Earth/Natural tones for pleasing visualization
    "Earth Tones": ["#5A8F29", "#8F5F2A", "#0F728F", "#8F2F0F", "#474A51", "#8F8F8F", "#291F19", "#598F8F"],
    
    # High contrast palette for maximum distinction
    "High Contrast": ["#E69F00", "#56B4E9", "#009E73", "#F0E442", "#0072B2", "#D55E00", "#CC79A7", "#000000"],
    
    # Blue-focused gradient for sequential data
    "Blues": sns.color_palette("Blues_r", 8).as_hex(),
    
    # Red-focused gradient for sequential data
    "Reds": sns.color_palette("Reds_r", 8).as_hex()
}

# --- in show_statistical_details, replace all key = ... assignments for latest_pvals with stat_key
# --- in plot_graph, replace all key = ... and latest_pvals lookups with stat_key
# --- add debug print if a key is missing in plot_graph when drawing annotation

class ExPlotApp:
    def optimize_legend_layout(self, ax, handles, labels, fontsize=10, max_fraction=0.8, min_ncol=1, max_ncol=None):
        """
        Determine the optimal number of legend columns so the legend doesn't exceed max_fraction of figure width.
        Returns the best ncol value.
        """
        from matplotlib.text import Text

        fig = ax.figure
        fig_width = fig.get_figwidth() * fig.dpi  # in pixels
        if max_ncol is None:
            max_ncol = len(labels)
        if len(labels) == 0:
            return 1
        # Estimate width of each label (roughly; matplotlib doesn't provide exact width until drawn)
        label_widths = []
        for label in labels:
            # Use a simple estimate: number of characters * fontsize * 0.6
            width = len(str(label)) * fontsize * 0.6
            label_widths.append(width)
        # Try different ncol values
        best_ncol = 1
        for ncol in range(1, max_ncol + 1):
            rows = int(np.ceil(len(labels) / ncol))
            # Estimate row width as sum of the widest label in each col
            col_widths = [0] * ncol
            for idx, width in enumerate(label_widths):
                col = idx % ncol
                col_widths[col] = max(col_widths[col], width)
            total_width = sum(col_widths) + 40  # Padding
            if total_width < fig_width * max_fraction:
                best_ncol = ncol
            else:
                break
        return best_ncol

    def stat_key(self, *args):
        """Create a standardized key for latest_pvals dict"""
        # Special case for 3-tuple (used for grouped data)
        if len(args) == 3:
            # For (category, group1, group2) keys, keep the format as is
            return tuple(args)
        
        # For standard 2-value keys, sort them for consistent retrieval
        else:
            # Convert all args to strings for consistent handling
            return tuple(sorted([str(a) for a in args]))
            
    def debug(self, message):
        """Print debug message with prefix"""
        if getattr(self, '_debug', False):
            print(f"[DEBUG] {message}")

    def _parse_custom_marks(self, text, axis):
        if text is None:
            return []
        items = []
        for raw in str(text).split(','):
            spec = raw.strip()
            if not spec:
                continue

            parts = [p.strip() for p in spec.split('|')]
            while len(parts) < 3:
                parts.append('')

            pos_str, label, flags = parts[0], parts[1], parts[2]
            # Backward compatibility:
            # - old format may omit flags entirely -> treat as minor tick
            # - tick flags: t=minor tick, T=major tick, n=no tick
            # - line flag: l=line
            # - major label flag: c=use custom text as major tick label (no numeric)
            if flags is None:
                flags = ''
            flags_raw = str(flags)
            if not flags_raw.strip():
                flags_raw = 't'
            flags_norm = ''
            for c in flags_raw:
                if c in ('T',):
                    flags_norm += 'T'
                elif c in ('t', 'l', 'n', 'c'):
                    flags_norm += c
            if not flags_norm:
                flags_norm = 't'

            tick_type = 'none' if 'n' in flags_norm else ('major' if 'T' in flags_norm else ('minor' if 't' in flags_norm else 'none'))
            line = 'l' in flags_norm
            label_mode = 'custom' if ('c' in flags_norm and tick_type == 'major') else 'number'

            pos = None
            try:
                pos = float(pos_str)
            except Exception:
                if axis == 'x' and hasattr(self, 'x_categorical_map'):
                    try:
                        pos = float(self.x_categorical_map[pos_str])
                    except Exception:
                        pos = None

            if pos is None:
                continue

            items.append({'pos': pos, 'label': label, 'tick_type': tick_type, 'line': line, 'label_mode': label_mode})
        return items

    def _apply_custom_marks_to_axis(self, ax, axis, marks, linewidth, fontsize, draw_labels=True):
        if not marks:
            return

        if axis == 'x':
            lo, hi = ax.get_xlim()
            in_view = [m for m in marks if min(lo, hi) <= m['pos'] <= max(lo, hi)]
        else:
            lo, hi = ax.get_ylim()
            in_view = [m for m in marks if min(lo, hi) <= m['pos'] <= max(lo, hi)]

        if not in_view:
            return

        major_positions = [m['pos'] for m in in_view if m.get('tick_type') == 'major']
        minor_positions = [m['pos'] for m in in_view if m.get('tick_type') == 'minor']

        major_custom_label_map = {
            m['pos']: (m.get('label') or '')
            for m in in_view
            if (m.get('tick_type') == 'major' and m.get('label_mode') == 'custom' and (m.get('label') or '').strip())
        }

        # Major ticks: merge into existing major ticks so they look identical
        if major_positions:
            try:
                if axis == 'x':
                    current = list(ax.get_xticks())
                    merged = sorted(set([*current, *major_positions]))
                    ax.xaxis.set_major_locator(FixedLocator(merged))
                else:
                    current = list(ax.get_yticks())
                    merged = sorted(set([*current, *major_positions]))
                    ax.yaxis.set_major_locator(FixedLocator(merged))
            except Exception:
                pass

        # If any major ticks have custom labels, override the major formatter to replace numeric text
        if major_custom_label_map:
            try:
                if axis == 'x':
                    base_formatter = ax.xaxis.get_major_formatter()
                else:
                    base_formatter = ax.yaxis.get_major_formatter()

                keys = list(major_custom_label_map.keys())
                def _fmt(x, pos=None):
                    try:
                        for k in keys:
                            if abs(float(x) - float(k)) <= 1e-9 * max(1.0, abs(float(k))):
                                return major_custom_label_map[k]
                    except Exception:
                        pass
                    try:
                        return base_formatter(x, pos)
                    except Exception:
                        return str(x)

                if axis == 'x':
                    ax.xaxis.set_major_formatter(FuncFormatter(_fmt))
                else:
                    ax.yaxis.set_major_formatter(FuncFormatter(_fmt))
            except Exception:
                pass

        # Minor ticks: merge into existing minor ticks (and ensure minor tick params are visible)
        if minor_positions:
            try:
                if axis == 'x':
                    current = list(ax.xaxis.get_minorticklocs())
                    merged = sorted(set([*current, *minor_positions]))
                    ax.xaxis.set_minor_locator(FixedLocator(merged))
                    ax.tick_params(axis='x', which='minor', direction='in', length=2, width=linewidth, color='black', bottom=True)
                else:
                    current = list(ax.yaxis.get_minorticklocs())
                    merged = sorted(set([*current, *minor_positions]))
                    ax.yaxis.set_minor_locator(FixedLocator(merged))
                    ax.tick_params(axis='y', which='minor', direction='in', length=2, width=linewidth, color='black', left=True)
            except Exception:
                pass

        for m in in_view:
            pos = m['pos']
            label = m.get('label') or ''
            tick_type = m.get('tick_type')
            line = bool(m.get('line'))
            label_mode = m.get('label_mode') or 'number'

            if line:
                try:
                    if axis == 'x':
                        ax.axvline(pos, color='black', linewidth=max(linewidth, 0.5), alpha=1.0, zorder=1, label='_nolegend_')
                    else:
                        ax.axhline(pos, color='black', linewidth=max(linewidth, 0.5), alpha=1.0, zorder=1, label='_nolegend_')
                except Exception:
                    pass

            # Avoid overlapping with major tick labels.
            # - major+custom is handled by formatter
            # - major+number should not draw a second label
            if draw_labels and label and tick_type != 'major':
                try:
                    if axis == 'x':
                        ax.text(pos, -0.04, label, transform=ax.get_xaxis_transform(), ha='center', va='top', fontsize=fontsize, clip_on=False)
                    else:
                        ax.text(-0.04, pos, label, transform=ax.get_yaxis_transform(), ha='right', va='center', fontsize=fontsize, clip_on=False)
                except Exception:
                    pass

    def _serialize_custom_marks_from_tree(self, tree):
        try:
            items = []
            for iid in tree.get_children():
                pos, label, tick_type, label_mode, line = tree.item(iid, 'values')
                flags = ''
                tick_type_s = str(tick_type).strip().lower()
                if tick_type_s == 'major':
                    flags += 'T'
                elif tick_type_s == 'minor':
                    flags += 't'
                else:
                    flags += 'n'
                if str(label_mode).strip().lower() == 'custom' and tick_type_s == 'major':
                    flags += 'c'
                if str(line).strip().upper() == 'Y':
                    flags += 'l'
                items.append(f"{pos}|{label}|{flags}")
            return ', '.join(items)
        except Exception:
            return ''

    def _load_custom_marks_into_tree(self, tree, text, axis):
        try:
            for iid in tree.get_children():
                tree.delete(iid)
            marks = self._parse_custom_marks(text, axis=axis)
            for m in marks:
                pos = m.get('pos')
                label = m.get('label') or ''
                tick_type = m.get('tick_type') or 'minor'
                label_mode = m.get('label_mode') or 'number'
                line = 'Y' if m.get('line') else 'N'
                tree.insert('', 'end', values=(pos, label, tick_type, label_mode, line))
        except Exception:
            pass

    def _sync_custom_marks_vars_from_ui(self):
        try:
            if hasattr(self, 'custom_x_marks_tree') and hasattr(self, 'custom_x_marks_var'):
                self.custom_x_marks_var.set(self._serialize_custom_marks_from_tree(self.custom_x_marks_tree))
            if hasattr(self, 'custom_y_marks_tree') and hasattr(self, 'custom_y_marks_var'):
                self.custom_y_marks_var.set(self._serialize_custom_marks_from_tree(self.custom_y_marks_tree))
        except Exception:
            pass

    def open_custom_marks_editor(self):
        if not hasattr(self, 'custom_x_marks_var'):
            self.custom_x_marks_var = tk.StringVar(value="")
        if not hasattr(self, 'custom_y_marks_var'):
            self.custom_y_marks_var = tk.StringVar(value="")

        win = tk.Toplevel(self.root)
        win.title("Custom Marks")
        win.transient(self.root)

        container = ttk.Frame(win, padding=8)
        container.pack(fill='both', expand=True)

        xgrp = ttk.LabelFrame(container, text="X marks", padding=6)
        xgrp.pack(fill='x', padx=2, pady=(0, 8))
        x_tree = ttk.Treeview(xgrp, columns=('pos', 'label', 'tick', 'label_mode', 'line'), show='headings', height=5)
        x_tree.heading('pos', text='Position')
        x_tree.heading('label', text='Label')
        x_tree.heading('tick', text='Tick')
        x_tree.heading('label_mode', text='Label')
        x_tree.heading('line', text='Line')
        x_tree.column('pos', width=110, anchor='w')
        x_tree.column('label', width=220, anchor='w')
        x_tree.column('tick', width=70, anchor='center')
        x_tree.column('label_mode', width=80, anchor='center')
        x_tree.column('line', width=55, anchor='center')
        x_tree.pack(fill='x', padx=2, pady=2)
        self._load_custom_marks_into_tree(x_tree, self.custom_x_marks_var.get(), axis='x')

        x_form = ttk.Frame(xgrp)
        x_form.pack(fill='x', padx=2, pady=2)
        ttk.Label(x_form, text="Pos:").grid(row=0, column=0, sticky='w')
        x_pos = ttk.Entry(x_form, width=12)
        x_pos.grid(row=0, column=1, sticky='w', padx=(4, 10))
        ttk.Label(x_form, text="Label:").grid(row=0, column=2, sticky='w')
        x_label = ttk.Entry(x_form, width=22)
        x_label.grid(row=0, column=3, sticky='w', padx=(4, 10))
        x_tick_type_var = tk.StringVar(value='minor')
        x_label_mode_var = tk.StringVar(value='number')
        x_line_var = tk.BooleanVar(value=False)
        ttk.Combobox(x_form, textvariable=x_tick_type_var, values=['none', 'minor', 'major'], width=7, state='readonly').grid(row=0, column=4, padx=(0, 6))
        ttk.Combobox(x_form, textvariable=x_label_mode_var, values=['number', 'custom'], width=8, state='readonly').grid(row=0, column=5, padx=(0, 10))
        ttk.Checkbutton(x_form, text="Line", variable=x_line_var).grid(row=0, column=6, padx=(0, 10))

        def _add_x():
            pos = x_pos.get().strip()
            if not pos:
                return
            tick_type = x_tick_type_var.get()
            label_mode = x_label_mode_var.get()
            line = 'Y' if x_line_var.get() else 'N'
            x_tree.insert('', 'end', values=(pos, x_label.get().strip(), tick_type, label_mode, line))
            x_pos.delete(0, tk.END)
            x_label.delete(0, tk.END)

        def _load_x_selection(*_args):
            sel = x_tree.selection()
            if not sel:
                return
            pos_v, label_v, tick_v, label_mode_v, line_v = x_tree.item(sel[0], 'values')
            x_pos.delete(0, tk.END)
            x_pos.insert(0, str(pos_v))
            x_label.delete(0, tk.END)
            x_label.insert(0, str(label_v))
            x_tick_type_var.set(str(tick_v).lower() if tick_v else 'minor')
            x_label_mode_var.set(str(label_mode_v).lower() if label_mode_v else 'number')
            x_line_var.set(str(line_v).strip().upper() == 'Y')

        def _update_x():
            sel = x_tree.selection()
            if not sel:
                return
            pos = x_pos.get().strip()
            if not pos:
                return
            tick_type = x_tick_type_var.get()
            label_mode = x_label_mode_var.get()
            line = 'Y' if x_line_var.get() else 'N'
            x_tree.item(sel[0], values=(pos, x_label.get().strip(), tick_type, label_mode, line))

        def _remove_x():
            for iid in x_tree.selection():
                x_tree.delete(iid)

        ttk.Button(x_form, text="Add", command=_add_x, width=10).grid(row=0, column=7, padx=(0, 4))
        ttk.Button(x_form, text="Update", command=_update_x, width=10).grid(row=0, column=8, padx=(0, 4))
        ttk.Button(x_form, text="Remove", command=_remove_x, width=10).grid(row=0, column=9)
        x_tree.bind('<<TreeviewSelect>>', _load_x_selection)

        ygrp = ttk.LabelFrame(container, text="Y marks", padding=6)
        ygrp.pack(fill='x', padx=2, pady=(0, 8))
        y_tree = ttk.Treeview(ygrp, columns=('pos', 'label', 'tick', 'label_mode', 'line'), show='headings', height=5)
        y_tree.heading('pos', text='Position')
        y_tree.heading('label', text='Label')
        y_tree.heading('tick', text='Tick')
        y_tree.heading('label_mode', text='Label')
        y_tree.heading('line', text='Line')
        y_tree.column('pos', width=110, anchor='w')
        y_tree.column('label', width=220, anchor='w')
        y_tree.column('tick', width=70, anchor='center')
        y_tree.column('label_mode', width=80, anchor='center')
        y_tree.column('line', width=55, anchor='center')
        y_tree.pack(fill='x', padx=2, pady=2)
        self._load_custom_marks_into_tree(y_tree, self.custom_y_marks_var.get(), axis='y')

        y_form = ttk.Frame(ygrp)
        y_form.pack(fill='x', padx=2, pady=2)
        ttk.Label(y_form, text="Pos:").grid(row=0, column=0, sticky='w')
        y_pos = ttk.Entry(y_form, width=12)
        y_pos.grid(row=0, column=1, sticky='w', padx=(4, 10))
        ttk.Label(y_form, text="Label:").grid(row=0, column=2, sticky='w')
        y_label = ttk.Entry(y_form, width=22)
        y_label.grid(row=0, column=3, sticky='w', padx=(4, 10))
        y_tick_type_var = tk.StringVar(value='minor')
        y_label_mode_var = tk.StringVar(value='number')
        y_line_var = tk.BooleanVar(value=False)
        ttk.Combobox(y_form, textvariable=y_tick_type_var, values=['none', 'minor', 'major'], width=7, state='readonly').grid(row=0, column=4, padx=(0, 6))
        ttk.Combobox(y_form, textvariable=y_label_mode_var, values=['number', 'custom'], width=8, state='readonly').grid(row=0, column=5, padx=(0, 10))
        ttk.Checkbutton(y_form, text="Line", variable=y_line_var).grid(row=0, column=6, padx=(0, 10))

        def _add_y():
            pos = y_pos.get().strip()
            if not pos:
                return
            tick_type = y_tick_type_var.get()
            label_mode = y_label_mode_var.get()
            line = 'Y' if y_line_var.get() else 'N'
            y_tree.insert('', 'end', values=(pos, y_label.get().strip(), tick_type, label_mode, line))
            y_pos.delete(0, tk.END)
            y_label.delete(0, tk.END)

        def _load_y_selection(*_args):
            sel = y_tree.selection()
            if not sel:
                return
            pos_v, label_v, tick_v, label_mode_v, line_v = y_tree.item(sel[0], 'values')
            y_pos.delete(0, tk.END)
            y_pos.insert(0, str(pos_v))
            y_label.delete(0, tk.END)
            y_label.insert(0, str(label_v))
            y_tick_type_var.set(str(tick_v).lower() if tick_v else 'minor')
            y_label_mode_var.set(str(label_mode_v).lower() if label_mode_v else 'number')
            y_line_var.set(str(line_v).strip().upper() == 'Y')

        def _update_y():
            sel = y_tree.selection()
            if not sel:
                return
            pos = y_pos.get().strip()
            if not pos:
                return
            tick_type = y_tick_type_var.get()
            label_mode = y_label_mode_var.get()
            line = 'Y' if y_line_var.get() else 'N'
            y_tree.item(sel[0], values=(pos, y_label.get().strip(), tick_type, label_mode, line))

        def _remove_y():
            for iid in y_tree.selection():
                y_tree.delete(iid)

        ttk.Button(y_form, text="Add", command=_add_y, width=10).grid(row=0, column=7, padx=(0, 4))
        ttk.Button(y_form, text="Update", command=_update_y, width=10).grid(row=0, column=8, padx=(0, 4))
        ttk.Button(y_form, text="Remove", command=_remove_y, width=10).grid(row=0, column=9)
        y_tree.bind('<<TreeviewSelect>>', _load_y_selection)

        button_row = ttk.Frame(container)
        button_row.pack(fill='x', pady=(4, 0))

        def _save_and_close():
            self.custom_x_marks_var.set(self._serialize_custom_marks_from_tree(x_tree))
            self.custom_y_marks_var.set(self._serialize_custom_marks_from_tree(y_tree))
            win.destroy()

        ttk.Button(button_row, text="Cancel", command=win.destroy, width=12).pack(side='right', padx=(4, 0))
        ttk.Button(button_row, text="Save", command=_save_and_close, width=12).pack(side='right')

    def _plot_xy_base(self, ax, df_plot, x_col, value_col, hue_col, value_cols, errorbar_black, linewidth, allow_legend=True):
        marker_size = self.xy_marker_size_var.get()
        marker_symbol = self.xy_marker_symbol_var.get()
        connect = self.xy_connect_var.get()
        draw_band = self.xy_draw_band_var.get()
        show_mean = self.xy_show_mean_var.get()
        show_mean_errorbars = self.xy_show_mean_errorbars_var.get()
        filled = self.xy_filled_var.get()
        line_style = self.xy_line_style_var.get()
        line_black = self.xy_line_black_var.get()

        if len(value_cols) == 1:
            color = self.custom_colors.get(self.single_color_var.get(), 'black')
            palette = [color]
        else:
            palette_name = self.palette_var.get()
            palette_full = self.custom_palettes.get(palette_name, ["#333333"])
            if hue_col and hue_col in df_plot.columns:
                hue_groups = df_plot[hue_col].dropna().unique()
                if len(palette_full) < len(hue_groups):
                    palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                palette = palette_full[:len(hue_groups)]
            else:
                palette = palette_full[:len(value_cols)]

        if show_mean:
            groupers = [x_col]
            if hue_col:
                groupers.append(hue_col)
            grouped = df_plot.groupby(groupers)[value_col]
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
                    y = group[value_col]
                    yerr = group['err']
                    ecolor = 'black' if errorbar_black else c
                    mfc = c if filled else 'none'
                    mec = c

                    if show_mean_errorbars:
                        ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec,
                                    markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor,
                                    label=str(name) if allow_legend else '_nolegend_')
                    else:
                        ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec,
                                markersize=marker_size, linewidth=linewidth, linestyle='None',
                                label=str(name) if allow_legend else '_nolegend_')

                    if draw_band:
                        df_band = pd.DataFrame({
                            'x': pd.to_numeric(x, errors='coerce'),
                            'y': pd.to_numeric(y, errors='coerce'),
                            'yerr': pd.to_numeric(yerr, errors='coerce')
                        }).dropna().sort_values('x')
                        if not df_band.empty:
                            ax.fill_between(df_band['x'], df_band['y'] - df_band['yerr'], df_band['y'] + df_band['yerr'],
                                            color=c, alpha=0.18, zorder=1)

                    if connect:
                        ax.plot(x, y, color='black' if line_black else c, linewidth=linewidth, alpha=0.7,
                                linestyle=line_style, label='_nolegend_')

                if allow_legend:
                    handles, labels = ax.get_legend_handles_labels()
                    if handles and len(handles) > 0:
                        self.place_legend(ax, handles, labels)
            else:
                c = palette[0]
                x_sorted = np.sort(df_plot[x_col].unique())
                y_means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                y_errors = [df_plot[df_plot[x_col] == x][value_col].std(ddof=1) if self.errorbar_type_var.get() == 'SD' else
                            df_plot[df_plot[x_col] == x][value_col].std(ddof=1) / np.sqrt(len(df_plot[df_plot[x_col] == x]))
                            for x in x_sorted]

                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                y_means_numeric = pd.to_numeric(y_means, errors='coerce')
                y_errors_numeric = pd.to_numeric(y_errors, errors='coerce')

                mfc = c if filled else 'none'
                mec = c
                ecolor = 'black' if errorbar_black else c

                if show_mean_errorbars:
                    ax.errorbar(x_sorted_numeric, y_means_numeric, yerr=y_errors_numeric, fmt=marker_symbol,
                                color=c, markerfacecolor=mfc, markeredgecolor=mec,
                                markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor)
                else:
                    ax.plot(x_sorted_numeric, y_means_numeric, marker=marker_symbol,
                            color=c, markerfacecolor=mfc, markeredgecolor=mec,
                            markersize=marker_size, linewidth=linewidth, linestyle='None')

                if draw_band:
                    df_band = pd.DataFrame({
                        'x': pd.to_numeric(x_sorted_numeric, errors='coerce'),
                        'y': pd.to_numeric(y_means_numeric, errors='coerce'),
                        'yerr': pd.to_numeric(y_errors_numeric, errors='coerce')
                    }).dropna().sort_values('x')
                    if not df_band.empty:
                        ax.fill_between(df_band['x'], df_band['y'] - df_band['yerr'], df_band['y'] + df_band['yerr'],
                                        color=c, alpha=0.18, zorder=1)
                if connect:
                    ax.plot(x_sorted_numeric, y_means_numeric,
                            color='black' if line_black else c,
                            linewidth=linewidth, alpha=0.7, linestyle=line_style)
        else:
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
                    edge = c
                    face = c if filled else 'none'
                    ax.scatter(group[x_col], group[value_col], marker=marker_symbol, s=marker_size**2, color=c,
                               label=str(name) if allow_legend else '_nolegend_', edgecolors=edge, facecolors=face,
                               linewidth=linewidth)
                    if draw_band:
                        x_sorted = np.sort(group[x_col].unique())
                        min_vals = [group[group[x_col] == x][value_col].min() for x in x_sorted]
                        max_vals = [group[group[x_col] == x][value_col].max() for x in x_sorted]
                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                        min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                        max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                        ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                    if connect:
                        x_sorted = np.sort(group[x_col].unique())
                        means = [group[group[x_col] == x][value_col].mean() for x in x_sorted]
                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                        means_numeric = pd.to_numeric(means, errors='coerce')
                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c,
                                linewidth=linewidth, alpha=0.7, linestyle=line_style)
                if allow_legend:
                    handles, labels = ax.get_legend_handles_labels()
                    if handles and len(handles) > 0:
                        self.place_legend(ax, handles, labels)
            else:
                c = palette[0]
                edge = c
                face = c if filled else 'none'
                ax.scatter(df_plot[x_col], df_plot[value_col], marker=marker_symbol, s=marker_size**2, color=c,
                           edgecolors=edge, facecolors=face, linewidth=linewidth)
                if draw_band:
                    x_sorted = np.sort(df_plot[x_col].unique())
                    min_vals = [df_plot[df_plot[x_col] == x][value_col].min() for x in x_sorted]
                    max_vals = [df_plot[df_plot[x_col] == x][value_col].max() for x in x_sorted]
                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                    min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                    max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                    ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                if connect:
                    x_sorted = np.sort(df_plot[x_col].unique())
                    means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                    means_numeric = pd.to_numeric(means, errors='coerce')
                    ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c,
                            linewidth=linewidth, alpha=0.7, linestyle=line_style)

    def _plot_xy_fitting(self, ax, df_plot, x_col, value_col, hue_col, palette, linewidth, update_results=True, allow_legend=True):
        if not (hasattr(self, 'use_fitting_var') and self.use_fitting_var.get()):
            return

        try:
            model_name = self.fitting_model_var.get()
            model_func = self.generate_model_function(model_name)
            model_info = self.fitting_models.get(model_name, {})
            parameters = model_info.get("parameters", [])
            param_names = [p[0] for p in parameters]
            p0 = [var.get() for _, var in self.param_entries]

            ci_option = self.fitting_ci_var.get()
            if ci_option == "68% (1σ)":
                sigma = 1.0
            elif ci_option == "95% (2σ)":
                sigma = 2.0
            else:
                sigma = 1.0

            if update_results:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"=== {model_name} Fitting Results ===\n\n")

            if hue_col and hue_col in df_plot.columns and len(df_plot[hue_col].unique()) > 1:
                group_names = df_plot[hue_col].unique()
                palette_name = self.palette_var.get()
                palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                if len(palette_full) < len(group_names):
                    palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                color_map = {name: palette_full[i] for i, name in enumerate(group_names)}

                any_fit = False
                for group_name in group_names:
                    group_df = df_plot[df_plot[hue_col] == group_name]
                    x_fit = pd.to_numeric(group_df[x_col], errors='coerce')
                    y_fit = pd.to_numeric(group_df[value_col], errors='coerce')
                    mask = ~(np.isnan(x_fit) | np.isnan(y_fit))
                    x_fit = x_fit[mask].values
                    y_fit = y_fit[mask].values
                    if len(x_fit) < len(p0) + 1:
                        continue
                    if len(x_fit) == 0 or model_func is None or len(p0) == 0:
                        continue

                    try:
                        x_smooth = np.linspace(min(x_fit), max(x_fit), 1000)
                        c = color_map.get(group_name, palette_full[0])
                        with warnings.catch_warnings():
                            warnings.simplefilter("ignore")
                            popt, pcov = curve_fit(model_func, x_fit, y_fit, p0=p0)
                            perr = np.sqrt(np.diag(pcov))

                        with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                            y_fit_curve = model_func(x_smooth, *popt)
                        finite_mask = np.isfinite(x_smooth) & np.isfinite(y_fit_curve)
                        x_smooth_plot = x_smooth[finite_mask]
                        y_fit_curve_plot = y_fit_curve[finite_mask]
                        if self.fitting_use_black_lines_var.get():
                            fit_color = 'black'
                        elif self.fitting_use_group_colors_var.get():
                            fit_color = c
                        else:
                            fit_color = 'red'

                        ax.plot(x_smooth_plot, y_fit_curve_plot, color=fit_color, linewidth=linewidth * 1.5,
                                linestyle='solid', label=f'Fit: {group_name}' if allow_legend else '_nolegend_')

                        if ci_option != "None":
                            y_lower = []
                            y_upper = []
                            x_ci = []
                            for x_val in x_smooth_plot:
                                with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                                    y_val = model_func(x_val, *popt)
                                if not np.isfinite(y_val):
                                    continue
                                y_err = 0
                                for i, param in enumerate(popt):
                                    delta = param * 0.001 if param != 0 else 0.001
                                    params_plus = popt.copy()
                                    params_plus[i] += delta
                                    with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                                        y_plus = model_func(x_val, *params_plus)
                                    if not np.isfinite(y_plus):
                                        continue
                                    partial_deriv = (y_plus - y_val) / delta
                                    y_err += (partial_deriv * perr[i]) ** 2
                                y_err = np.sqrt(y_err) * sigma
                                if not np.isfinite(y_err):
                                    continue
                                x_ci.append(x_val)
                                y_lower.append(y_val - y_err)
                                y_upper.append(y_val + y_err)

                            if self.fitting_use_black_bands_var.get():
                                band_color = 'black'
                            else:
                                band_color = fit_color
                            if len(x_ci) > 1:
                                ax.fill_between(x_ci, y_lower, y_upper, alpha=0.2, color=band_color,
                                            label=f'{group_name} {ci_option} CI' if allow_legend else '_nolegend_')

                        y_pred = model_func(x_fit, *popt)
                        ss_res = np.sum((y_fit - y_pred) ** 2)
                        ss_tot = np.sum((y_fit - np.mean(y_fit)) ** 2)
                        r_squared = 1 - (ss_res / ss_tot)

                        if update_results:
                            self.result_text.insert(tk.END, f"Group: {group_name}\n")
                            for i, param_name in enumerate(param_names):
                                if i < len(popt):
                                    self.result_text.insert(tk.END, f"  {param_name} = {popt[i]:.6f} ± {perr[i]:.6f}\n")
                            self.result_text.insert(tk.END, f"  R² = {r_squared:.6f}\n\n")
                            equation = model_info.get("formula", "")
                            for line in equation.split('\n'):
                                if line.strip().startswith('y ='):
                                    eq = line.strip()
                                    for i, param_name in enumerate(param_names):
                                        if i < len(popt):
                                            eq = eq.replace(param_name, f"{popt[i]:.4f}")
                                    self.result_text.insert(tk.END, f"  {eq}\n")
                            self.result_text.insert(tk.END, "\n")

                        any_fit = True
                    except Exception as e:
                        if update_results:
                            self.result_text.insert(tk.END, f"Group: {group_name} - Fitting failed: {str(e)}\n\n")

                if update_results and not any_fit:
                    self.result_text.insert(tk.END, "No groups could be successfully fitted.\n")
                    self.result_text.insert(tk.END, "Check that your data has enough points per group and try different initial parameters.")

                if allow_legend:
                    handles, labels = ax.get_legend_handles_labels()
                    if handles and len(handles) > 0:
                        self.place_legend(ax, handles, labels)

            else:
                x_fit = pd.to_numeric(df_plot[x_col], errors='coerce')
                y_fit = pd.to_numeric(df_plot[value_col], errors='coerce')
                mask = ~(np.isnan(x_fit) | np.isnan(y_fit))
                x_fit = x_fit[mask].values
                y_fit = y_fit[mask].values
                if len(x_fit) == 0 or model_func is None or len(p0) == 0:
                    return

                x_smooth = np.linspace(min(x_fit), max(x_fit), 1000)
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    popt, pcov = curve_fit(model_func, x_fit, y_fit, p0=p0)
                    perr = np.sqrt(np.diag(pcov))

                with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                    y_fit_curve = model_func(x_smooth, *popt)
                finite_mask = np.isfinite(x_smooth) & np.isfinite(y_fit_curve)
                x_smooth_plot = x_smooth[finite_mask]
                y_fit_curve_plot = y_fit_curve[finite_mask]
                if self.fitting_use_black_lines_var.get():
                    fit_color = 'black'
                elif self.fitting_use_group_colors_var.get():
                    fit_color = palette[0] if palette else 'black'
                else:
                    fit_color = 'red'

                ax.plot(x_smooth_plot, y_fit_curve_plot, color=fit_color, linewidth=linewidth * 1.5,
                        linestyle='solid', label=f'Fit: {model_name}' if allow_legend else '_nolegend_')

                if ci_option != "None":
                    y_lower = []
                    y_upper = []
                    x_ci = []
                    for x_val in x_smooth_plot:
                        with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                            y_val = model_func(x_val, *popt)
                        if not np.isfinite(y_val):
                            continue
                        y_err = 0
                        for i, param in enumerate(popt):
                            delta = param * 0.001 if param != 0 else 0.001
                            params_plus = popt.copy()
                            params_plus[i] += delta
                            with np.errstate(divide='ignore', invalid='ignore', over='ignore', under='ignore'):
                                y_plus = model_func(x_val, *params_plus)
                            if not np.isfinite(y_plus):
                                continue
                            partial_deriv = (y_plus - y_val) / delta
                            y_err += (partial_deriv * perr[i]) ** 2
                        y_err = np.sqrt(y_err) * sigma
                        if not np.isfinite(y_err):
                            continue
                        x_ci.append(x_val)
                        y_lower.append(y_val - y_err)
                        y_upper.append(y_val + y_err)

                    if self.fitting_use_black_bands_var.get():
                        band_color = 'black'
                    else:
                        band_color = fit_color
                    if len(x_ci) > 1:
                        ax.fill_between(x_ci, y_lower, y_upper, alpha=0.2, color=band_color,
                                    label=f'{ci_option} Confidence' if allow_legend else '_nolegend_')

                if update_results:
                    for i, param_name in enumerate(param_names):
                        if i < len(popt):
                            self.result_text.insert(tk.END, f"{param_name} = {popt[i]:.6f} ± {perr[i]:.6f}\n")
                    y_pred = model_func(x_fit, *popt)
                    ss_res = np.sum((y_fit - y_pred) ** 2)
                    ss_tot = np.sum((y_fit - np.mean(y_fit)) ** 2)
                    r_squared = 1 - (ss_res / ss_tot)
                    self.result_text.insert(tk.END, f"\nR² = {r_squared:.6f}\n")
                    self.result_text.insert(tk.END, f"\nFitted equation:\n")
                    equation = model_info.get("formula", "")
                    for line in equation.split('\n'):
                        if line.strip().startswith('y ='):
                            eq = line.strip()
                            for i, param_name in enumerate(param_names):
                                if i < len(popt):
                                    eq = eq.replace(param_name, f"{popt[i]:.4f}")
                            self.result_text.insert(tk.END, f"{eq}\n")

                if allow_legend:
                    handles, labels = ax.get_legend_handles_labels()
                    if handles and len(handles) > 0:
                        self.place_legend(ax, handles, labels)

        except Exception as e:
            if update_results:
                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"Error in model fitting: {str(e)}\n")
                self.result_text.insert(tk.END, "Make sure you've selected numeric data columns for XY plotting.")
        
    def _add_significance_legend(self, details_text):
        """Add the significance level legend to the details panel"""
        details_text.insert(tk.END, "Significance levels:\n")
        details_text.insert(tk.END, "     ns p > 0.05\n")
        details_text.insert(tk.END, "    * p ≤ 0.05\n")
        details_text.insert(tk.END, "   ** p ≤ 0.01\n")
        details_text.insert(tk.END, "  *** p ≤ 0.001\n")
        details_text.insert(tk.END, " **** p ≤ 1e-05\n\n")
        
        # Get the current alpha value from the variable
        try:
            alpha_value = float(self.alpha_level_var.get()) if hasattr(self, 'alpha_level_var') else 0.05
        except (ValueError, AttributeError):
            alpha_value = 0.05  # Default if there's an error
            
        details_text.insert(tk.END, f"Current alpha level: {alpha_value:.2f}\n\n")
        
    def _show_old_metric_comparisons(self, details_text, metric_pairs, metric_results, x_categories):
        """Display metric comparisons using the old format as a fallback"""
        details_text.insert(tk.END, "P-values for metric comparisons:\n\n")
        for pair in metric_pairs:
            metric1, metric2 = pair
            details_text.insert(tk.END, f"Comparing {metric1} vs {metric2}:\n")
            
            for x_cat in x_categories:
                result_key = (x_cat, metric1, metric2)
                if result_key in metric_results:
                    p_val = metric_results[result_key]
                    sig = self.pval_to_annotation(p_val)
                    details_text.insert(tk.END, f"  {x_cat}: p = {p_val:.4g} {sig}\n")
        details_text.insert(tk.END, "\n")

    def calculate_statistics(self, df_plot, x_col, value_col, hue_col=None):
        """
        Centralized statistics calculation method for ExPlot.
        
        This function is a wrapper that uses the explot_stats module to perform all statistical
        calculations, ensuring consistency between plot annotations and the statistical details panel.
        
        Args:
            df_plot (pd.DataFrame): The DataFrame containing the data to analyze
            x_col (str): The column name to use for x-axis categories
            value_col (str): The column name containing the values to compare
            hue_col (str, optional): The column name for grouping data, if applicable
        
        Returns:
            dict: Statistical results from the analysis
        """
        try:
            # Check if we have grouped data with only one group
            if hue_col and hue_col in df_plot.columns:
                unique_groups = df_plot[hue_col].dropna().unique()
                if len(unique_groups) == 1:
                    import tkinter.messagebox as messagebox
                    messagebox.showinfo(
                        "Statistical Analysis",
                        "Only one group detected in the group column. "
                        "To calculate statistics between X categories, please set the group to 'None'."
                    )
                    # Return empty results
                    return {
                        'pvals': {},
                        'x_col': x_col,
                        'value_col': value_col,
                        'hue_col': hue_col,
                        'summary': "Only one group detected. Set group to 'None' to calculate statistics.",
                        'error': True
                    }
            
            # Import the statistical module
            from explot_stats import calculate_statistics as stats_calc
            
            # Get relevant settings from the app
            app_settings = {
                'alpha_level': float(self.alpha_level_var.get()) if hasattr(self, 'alpha_level_var') else 0.05,
                'test_type': self.ttest_type_var.get() if hasattr(self, 'ttest_type_var') else "Independent t-test",
                'alternative': self.ttest_alternative_var.get() if hasattr(self, 'ttest_alternative_var') else "two-sided",
                'anova_type': self.anova_type_var.get() if hasattr(self, 'anova_type_var') else "One-way ANOVA",
                'posthoc_type': self.posthoc_type_var.get() if hasattr(self, 'posthoc_type_var') else "Tukey's HSD"
            }
            
            # Delegate to the statistics module
            results = stats_calc(df_plot, x_col, value_col, hue_col, app_settings)
            
            # Store results for backward compatibility
            self.latest_stats = results
            self.latest_pvals = results.get('pvals', {})
            self.latest_test_info = results.get('test_info', {})
            
            # Debug what tests were performed
            if self.latest_test_info:
                print(f"[DEBUG] Stored {len(self.latest_test_info)} test details")
                for key, test_details in list(self.latest_test_info.items())[:3]:  # Show first 3 for debug
                    test_used = test_details.get('test_used', 'Unknown test')
                    print(f"[DEBUG] Test for {key}: {test_used}")
            
            return results
            
        except ImportError as e:
            error_msg = "Error: The explot_stats module is required for statistical analysis. Please install the module or contact the developer."
            print(f"[ERROR] {error_msg}")
            
            # Return error result
            return {
                'pvals': {},
                'x_col': x_col,
                'value_col': value_col,
                'hue_col': hue_col,
                'summary': error_msg,
                'error': True
            }
    
    def show_app_status(self, message):
        """Display status message in the status bar"""
        if hasattr(self, 'status_label'):
            self.status_label.config(text=message)

    def prepare_working_dataframe(self):
        """Prepare a working dataframe that will be used for both plotting and statistics.
        This ensures consistency between plot and statistics by using the same data source.
        Returns the working dataframe and relevant column info.
        """
        if self.df is None:
            return None, None, None, None
            
        # Get column selections
        x_col = self.xaxis_var.get()
        group_col = self.group_var.get()
        if not group_col or group_col.strip() == '' or group_col == 'None':
            group_col = None
            
        value_cols = [col for var, col in self.value_vars if var.get() and col != x_col]
        
        # Handle empty selections gracefully
        if not x_col or not value_cols:
            return None, None, None, None
            
        # Create a working copy of the dataframe
        df_work = self.df.copy()
        
        # Filter out excluded X values if any exist
        if hasattr(self, 'excluded_x_values') and self.excluded_x_values:
            df_work = df_work[~df_work[x_col].isin(self.excluded_x_values)]
            if df_work.empty:
                return None, None, None, None
        
        # Apply renames to original values if any exist
        if self.xaxis_renames:
            renamed_series = df_work[x_col].map(lambda x: self.xaxis_renames.get(x, x))
            df_work['_renamed_x'] = renamed_series
            original_x_col = x_col
            x_col = '_renamed_x'
        else:
            original_x_col = x_col
            
        # Store the working dataframe as an instance attribute for later use
        self.df_work = df_work
        
        return df_work, x_col, value_cols, group_col
        

                
    def export_data(self):
        """Export the current working dataframe to a CSV file."""
        if self.df_work is None and self.df is not None:
            # Prepare the working dataframe if it doesn't exist yet
            self.prepare_working_dataframe()
            
        if self.df_work is None:
            messagebox.showinfo("No Data", "There is no data to export.")
            return
            
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Export Data As"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Export the dataframe to CSV
            self.df_work.to_csv(file_path, index=False)
            messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Failed to export data: {str(e)}")
            
    def _add_statannotations(self, ax, df, x_col, group_col=None):
        """Add statistical annotations using the statannotations package.
        
        Args:
            ax: Matplotlib axis to add annotations to
            df: DataFrame containing the data
            x_col: Column name for x-axis categories
            group_col: Optional column name for grouping within x-categories
        """
        if not hasattr(self, 'latest_pvals') or not self.latest_pvals:
            print("[DEBUG] No p-values available for annotations")
            return
            
        # Get the value column to plot
        if hasattr(self, 'current_plot_info'):
            value_cols = self.current_plot_info.get('value_cols', [])
            if value_cols:
                value_col = value_cols[0]
            else:
                # Fallback to first numeric column
                numeric_cols = df.select_dtypes(include=['number']).columns
                if len(numeric_cols) > 0:
                    value_col = numeric_cols[0]
                else:
                    print("[DEBUG] No numeric columns found for annotations")
                    return
        else:
            print("[DEBUG] No current_plot_info available")
            return
        
        print(f"[DEBUG] Using value column: {value_col}")
        print(f"[DEBUG] X column: {x_col}")
        print(f"[DEBUG] Group column: {group_col}")
        
        # Create pairs for annotation based on p-values
        pairs = []
        pvalues = []
        
        if group_col and group_col in df.columns:
            # Grouped data - pairs are within each x-category
            x_categories = df[x_col].dropna().unique()
            groups = sorted(df[group_col].dropna().unique())  # Sort groups for consistency
            print(f"[DEBUG] X categories: {x_categories}")
            print(f"[DEBUG] Groups: {groups}")
            
            for x_cat in x_categories:
                # Get all group pairs for this x-category
                for i in range(len(groups)):
                    for j in range(i+1, len(groups)):
                        group1, group2 = groups[i], groups[j]
                        
                        # Try different key formats to find the p-value
                        key_formats = [
                            (x_cat, group1, group2),
                            (x_cat, group2, group1),
                            (str(x_cat), group1, group2),
                            (str(x_cat), group2, group1)
                        ]
                        
                        # Look for a p-value match
                        p_val = None
                        for key in key_formats:
                            if key in self.latest_pvals:
                                p_val = self.latest_pvals[key]
                                break
                        
                        # If significant, add to pairs
                        if p_val is not None and p_val <= float(self.alpha_level_var.get()):
                            # Format for statannotations: (x_category, group1) vs (x_category, group2)
                            pairs.append([(x_cat, group1), (x_cat, group2)])
                            pvalues.append(p_val)
                            print(f"[DEBUG] Added pair: {(x_cat, group1)} vs {(x_cat, group2)} with p={p_val}")
        else:
            # Ungrouped data - pairs are between x-categories
            x_categories = df[x_col].dropna().unique()
            
            for i in range(len(x_categories)):
                for j in range(i+1, len(x_categories)):
                    cat1, cat2 = x_categories[i], x_categories[j]
                    
                    # Try different key formats to find the p-value
                    key_formats = [
                        (cat1, cat2),
                        (cat2, cat1),
                        (str(cat1), str(cat2)),
                        (str(cat2), str(cat1))
                    ]
                    
                    # Look for a p-value match
                    p_val = None
                    for key in key_formats:
                        if key in self.latest_pvals:
                            p_val = self.latest_pvals[key]
                            break
                    
                    # If significant, add to pairs
                    if p_val is not None and p_val <= float(self.alpha_level_var.get()):
                        pairs.append([cat1, cat2])
                        pvalues.append(p_val)
                        print(f"[DEBUG] Added pair: {cat1} vs {cat2} with p={p_val}")
        
        if not pairs:
            print("[DEBUG] No significant pairs to annotate")
            return
        
        print(f"[DEBUG] Found {len(pairs)} significant pairs to annotate")
        
        try:
            # Import statannotations here to avoid dependency if not used
            from statannotations.Annotator import Annotator
            
            # Create annotator
            annotator = Annotator(
                ax=ax,
                pairs=pairs,
                data=df,
                x=x_col,
                y=value_col,
                hue=group_col if group_col else None,
                verbose=True
            )
            
            # Get current alpha level for custom p-value thresholds
            try:
                alpha = float(self.alpha_level_var.get())
            except (ValueError, AttributeError):
                alpha = 0.05  # Default if not set or invalid
                    
            # Configure p-value thresholds directly in the format required by statannotations
            pvalue_thresholds = [
                [1e-4, "****"],  # p < 0.01%
                [1e-3, "***"],   # p < 0.1%
                [1e-2, "**"],    # p < 1%
                [alpha, "*"]     # p < alpha (default 0.05 or user-specified)
            ]
                
            print(f"[DEBUG] Setting annotator pvalue format: {pvalue_thresholds}")
                
            # Configure annotations with our custom pvalue thresholds
            annotator.configure(test=None, text_format='star', 
                                pvalue_format=pvalue_format,
                                loc='inside',
                                line_width=self.linewidth.get()
                                )
            # Set custom p-values and annotate using the current alpha level
            # Get current alpha level from dropdown
            try:
                current_alpha = float(self.alpha_level_var.get())
            except (ValueError, AttributeError):
                current_alpha = 0.05  # Default if not set or invalid
                
            # Generate annotations with the current alpha setting
            annotations = []
            for p in pvalues:
                annotation = self.pval_to_annotation(p)  # This already uses alpha_level_var internally
                annotations.append(annotation)
                
            # Debug to check consistency
            print(f"[DEBUG] Alpha level: {current_alpha}, creating annotations: {list(zip([f'{p:.4g}' for p in pvalues], annotations))}")
                
            annotator.set_custom_annotations(annotations)
            annotator.annotate()
            
            print("[DEBUG] Successfully added statannotations")
            
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error with statannotations: {e}")
            print(traceback.format_exc())
            raise  # Re-raise the exception to be handled by the caller
    
    def add_statistics_to_plot(self):
        """Add statistical annotations to the current plot based on calculated p-values.
        Uses the statannotations package for cleaner annotations.
        """
        print("\n[DEBUG] Starting add_statistics_to_plot...")
        
        # Check if we have a figure and p-values
        if not hasattr(self, 'fig') or self.fig is None:
            print("[DEBUG] No figure available for annotations")
            return
        
        if not hasattr(self, 'latest_pvals') or not self.latest_pvals:
            print("[DEBUG] No p-values available for annotations")
            return
        else:
            print(f"[DEBUG] Found {len(self.latest_pvals)} p-values in latest_pvals")
            print("[DEBUG] Latest p-values:", list(self.latest_pvals.items())[:5], "...")
        
        # Get current axes
        ax = self.fig.axes[0] if self.fig.axes else None
        if not ax:
            print("[DEBUG] No axes available for annotations")
            return
        
        # Get current plot info
        if not hasattr(self, 'current_plot_info'):
            print("[DEBUG] No current_plot_info available for annotations")
            return
        
        # Get plot information
        plot_kind = self.plot_kind_var.get()
        x_col = self.current_plot_info.get('x_col')
        group_col = self.current_plot_info.get('group_col')
        
        # Only implement for bar and box plots
        if plot_kind not in ["bar", "box", "violin"]:
            print(f"[DEBUG] Statistical annotations for {plot_kind} plot not implemented")
            return
        
        # Get working dataframe
        if not hasattr(self, 'df_work') or self.df_work is None:
            print("[DEBUG] No working dataframe available for annotations")
            return
        
        df_work = self.df_work.copy()
        print(f"[DEBUG] Adding annotations with data shape: {df_work.shape}")
        print(f"[DEBUG] Columns: {df_work.columns.tolist()}")
        print(f"[DEBUG] X column: {x_col}, Group column: {group_col}")
        
        try:
            # Clear existing annotations if any
            for child in ax.get_children():
                if isinstance(child, matplotlib.text.Annotation):
                    child.remove()
            
            # Extend the y-axis limits to make room for annotations
            ymin, ymax = ax.get_ylim()
            y_range = ymax - ymin
            ax.set_ylim(ymin, ymax + y_range * 0.3)  # Add 30% more space at the top
            
            # Use statannotations for annotations
            self._add_statannotations(ax, df_work, x_col, group_col)
            
            # Update the canvas
            if hasattr(self, 'canvas'):
                self.canvas.draw()
                
        except Exception as e:
            import traceback
            print(f"[DEBUG] Error adding annotations: {e}")
            print(traceback.format_exc())
            messagebox.showerror("Annotation Error", 
                               f"Error adding statistical annotations: {str(e)}")
            
    def _add_ungrouped_annotations(self, ax, df, x_col):
        """Add statistical annotations for ungrouped data (comparing x-axis categories)"""
        print("[DEBUG] Starting _add_ungrouped_annotations...")
        # Get x-axis categories
        x_categories = df[x_col].dropna().unique()
        print(f"[DEBUG] Found {len(x_categories)} x-axis categories: {x_categories}")
        
        if len(x_categories) < 2:
            return  # Need at least 2 categories for comparison
            
        # Create pairs for annotation
        pairs = []
        print(f"[DEBUG] Creating pairs from x_categories: {x_categories}")
        for i in range(len(x_categories)):
            for j in range(i+1, len(x_categories)):
                cat1, cat2 = x_categories[i], x_categories[j]
                print(f"[DEBUG] Checking pair: {cat1} vs {cat2}")
                
                # Check if we have a p-value for this pair
                p_val = None
                
                # Try different key formats
                key_formats = [
                    self.stat_key(cat1, cat2),  # String-based key
                    (cat1, cat2),              # Direct tuple
                    (cat2, cat1),              # Reversed tuple
                    (i, j),                    # Index-based
                    (j, i),                    # Reversed index
                    (float(i), float(j)),      # Float indices
                    (float(j), float(i))       # Reversed float indices
                ]
                
                # Look for a match in latest_pvals
                for key in key_formats:
                    if key in self.latest_pvals:
                        p_val = self.latest_pvals[key]
                        break
                        
                # If we found a significant p-value, add it to the pairs
                if p_val is not None and p_val <= float(self.alpha_level_var.get()):
                    # IMPORTANT: Use simple tuple format for statannotations
                    pairs.append((cat1, cat2))
                    
        # If we have pairs to annotate, create the annotations
        if pairs:
            # Get formatting parameters
            fontsize = int(self.fontsize_entry.get())
            linewidth = float(self.linewidth.get())
            
            # Create annotator
            # Choose the value column carefully
            if value_col := self.current_plot_info.get('value_cols', [None])[0]:
                y_col = value_col
            else:
                # Fallback to the last numeric column as a best guess
                numeric_cols = df.select_dtypes(include=['number']).columns
                y_col = numeric_cols[-1] if len(numeric_cols) > 0 else df.columns[-1]
                
            annotator = Annotator(ax, pairs, data=df, x=x_col, y=y_col)
            
            # Generate custom p-value text based on stored p-values
            pvalues = []
            for pair in pairs:
                cat1, cat2 = pair[0][0], pair[1][0]
                
                # Find p-value for this pair
                p_val = None
                for key in self.latest_pvals:
                    if (isinstance(key, tuple) and len(key) == 2 and 
                        ((str(key[0]) == str(cat1) and str(key[1]) == str(cat2)) or
                         (str(key[0]) == str(cat2) and str(key[1]) == str(cat1)))):
                        p_val = self.latest_pvals[key]
                        break
                        
                # If we couldn't find by string, try index
                if p_val is None:
                    i, j = np.where(x_categories == cat1)[0][0], np.where(x_categories == cat2)[0][0]
                    index_keys = [(i, j), (j, i), (float(i), float(j)), (float(j), float(i))]
                    
                    for key in index_keys:
                        if key in self.latest_pvals:
                            p_val = self.latest_pvals[key]
                            break
                            
                pvalues.append(p_val)
                
            # Add annotations with custom p-values
            try:
                # Configure the annotator with enhanced settings for better visibility
                print(f"[DEBUG] Configuring annotator with {len(pairs)} pairs")
                
                # Get current alpha level for custom p-value thresholds
                try:
                    alpha = float(self.alpha_level_var.get())
                except (ValueError, AttributeError):
                    alpha = 0.05  # Default if not set or invalid
                    
                # Configure p-value thresholds directly in the format required by statannotations
                pvalue_format = {
                    'text_format': 'star',
                    'pvalue_thresholds': [
                        (alpha/5000, '****'),  # 4 stars threshold (e.g., 0.00001 at alpha=0.05)
                        (alpha/50, '***'),    # 3 stars threshold (e.g., 0.001 at alpha=0.05)
                        (alpha/5, '**'),      # 2 stars threshold (e.g., 0.01 at alpha=0.05)
                        (alpha, '*'),         # 1 star threshold (alpha itself)
                        (1, 'ns')             # Not significant
                    ]
                }
                
                print(f"[DEBUG] Setting annotator pvalue format: {pvalue_format_dict['pvalue_thresholds']}")
                
                # Configure the annotator with very aggressive settings for maximum visibility
                annotator.configure(test=None, text_format='star', 
                                    pvalue_format=pvalue_format,
                                    loc='inside',
                                    line_width=self.linewidth.get()
                                    )
                
                # Set custom p-values and annotations
                if pvalues:
                    print(f"[DEBUG] Setting {len(pvalues)} custom annotations")
                    annotations = [self.pval_to_annotation(p) for p in pvalues]
                    print(f"[DEBUG] Annotation symbols: {annotations}")
                    annotator.set_custom_annotations(annotations)
                    print("[DEBUG] About to call annotator.annotate()")
                    annotator.annotate()
                    print(f"[DEBUG] Successfully added {len(pvalues)} annotation(s) for categories")
                else:
                    print("[DEBUG] No significant p-values to annotate")
            except Exception as e:
                import traceback
                print(f"[DEBUG] Error with statannotations: {e}")
                print(traceback.format_exc())
                # Fallback to a simpler method if statannotations fails
                self._fallback_add_annotations(ax, pairs, pvalues)
            
    def _add_grouped_annotations(self, ax, df, x_col, group_col):
        """Add statistical annotations for grouped data (comparing groups within each x-axis category)"""
        # Get x-axis categories and groups
        x_categories = df[x_col].dropna().unique()
        groups = df[group_col].dropna().unique()
        
        if len(groups) < 2:
            return  # Need at least 2 groups for comparison
        
        # For each x-axis category, compare all groups
        for x_cat in x_categories:
            # Get the subset of data for this x-category
            df_cat = df[df[x_col] == x_cat]
            
            # Create pairs for annotation within this category
            pairs = []
            pvalues = []
            
            for i in range(len(groups)):
                for j in range(i+1, len(groups)):
                    group1, group2 = groups[i], groups[j]
                    
                    # Check if we have a p-value for this triplet
                    p_val = None
                    
                    # Try different key formats for triplets (x_cat, group1, group2)
                    key_formats = [
                        (x_cat, group1, group2),   # Direct triplet
                        (x_cat, group2, group1),   # Reversed groups
                        (str(x_cat), group1, group2),  # String version
                        (str(x_cat), group2, group1)   # String reversed
                    ]
                    
                    # Also try with index numbers
                    x_idx = np.where(x_categories == x_cat)[0][0] if len(np.where(x_categories == x_cat)[0]) > 0 else -1
                    if x_idx >= 0:
                        key_formats.extend([
                            (x_idx, group1, group2),
                            (x_idx, group2, group1),
                            (float(x_idx), group1, group2),
                            (float(x_idx), group2, group1)
                        ])
                    
                    # Look for a match in latest_pvals
                    for key in key_formats:
                        if key in self.latest_pvals:
                            p_val = self.latest_pvals[key]
                            break
                    
                    # If we found a significant p-value, add it to the pairs
                    if p_val is not None and p_val <= float(self.alpha_level_var.get()):
                        # Modify x_cat to include the group for statannotations format
                        # This creates labels like 'CategoryA_Group1' that statannotations can work with
                        formatted_group1 = f"{x_cat}_{group1}"
                        formatted_group2 = f"{x_cat}_{group2}"
                        pairs.append((formatted_group1, formatted_group2))
                        pvalues.append(p_val)
            
            # If we have pairs to annotate for this category, create the annotations
            if pairs:
                # For grouped data, we need to create a modified dataframe with combined x_cat and group columns
                # This creates a new column with values like 'CategoryA_Group1' that statannotations can use
                df_modified = df.copy()
                df_modified['x_group'] = df_modified[x_col].astype(str) + '_' + df_modified[group_col].astype(str)
                
                # Choose the value column carefully
                if value_col := self.current_plot_info.get('value_cols', [None])[0]:
                    y_col = value_col
                else:
                    # Fallback to the last numeric column as a best guess
                    numeric_cols = df_modified.select_dtypes(include=['number']).columns
                    y_col = numeric_cols[-1] if len(numeric_cols) > 0 else df_modified.columns[-1]
                
                # Now we use the modified dataframe with the 'x_group' column instead of separate x and hue
                print(f"[DEBUG] Creating grouped annotator with pairs: {pairs}")
                annotator = Annotator(ax, pairs, data=df_modified, x='x_group', y=y_col)
                
                # Add annotations with custom p-values
                try:
                    # Configure the annotator with enhanced settings for better visibility
                    # Get plot limits for positioning
                    ymin, ymax = ax.get_ylim()
                    y_range = ymax - ymin
                    
                    # Calculate a good position above the data
                    # Adjust these values as needed for your specific plots
                    line_pos = ymax + y_range * 0.1  # Start 10% above the top of the plot
                    
                    # Get current alpha level for custom p-value thresholds
                    try:
                        alpha = float(self.alpha_level_var.get())
                    except (ValueError, AttributeError):
                        alpha = 0.05  # Default if not set or invalid
                    
                    # Get the current alpha level for custom p-value thresholds
                    try:
                        alpha = float(self.alpha_level_var.get())
                    except (ValueError, AttributeError):
                        alpha = 0.05  # Default if not set or invalid
                    
                        # Configure p-value thresholds directly in the format required by statannotations
                    pvalue_format = {
                        'text_format': 'star',
                        'pvalue_thresholds': [
                            (alpha/5000, '****'),  # 4 stars threshold (e.g., 0.00001 at alpha=0.05)
                            (alpha/50, '***'),    # 3 stars threshold (e.g., 0.001 at alpha=0.05)
                            (alpha/5, '**'),      # 2 stars threshold (e.g., 0.01 at alpha=0.05)
                            (alpha, '*'),         # 1 star threshold (alpha itself)
                            (1, 'ns')             # Not significant
                        ]
                    }
                    
                    # Configure the annotator with the same settings as _add_ungrouped_annotations
                    annotator.configure(test=None, text_format='star', 
                                        pvalue_format=pvalue_format,
                                        loc='inside',
                                        line_width=self.linewidth.get()
                                        )
                    
                    # Generate annotations with the current alpha setting using the same method as _add_ungrouped_annotations
                    annotations = [self.pval_to_annotation(p) for p in pvalues]
                    
                    # Debug to check consistency
                    print(f"[DEBUG] Alpha level: {alpha}, creating grouped annotations: {list(zip([f'{p:.4g}' for p in pvalues], annotations))}")
                    
                    # Set the custom annotations
                    annotator.set_custom_annotations(annotations)
                    
                    # Increase the y-limit to make room for annotations
                    ax.set_ylim(ymin, ymax + y_range * 0.3)
                    
                    # Set custom p-values and annotations using the current alpha level
                    if pvalues:
                        # Get current alpha level from dropdown
                        try:
                            current_alpha = float(self.alpha_level_var.get())
                        except (ValueError, AttributeError):
                            current_alpha = 0.05  # Default if not set or invalid
                        
                        # Generate annotations with the current alpha setting
                        annotations = []
                        for p in pvalues:
                            annotation = self.pval_to_annotation(p)  # This already uses alpha_level_var internally
                            annotations.append(annotation)
                        
                        # Debug to check consistency
                        print(f"[DEBUG] Alpha level: {current_alpha}, creating annotations for {x_cat}: {list(zip([f'{p:.4g}' for p in pvalues], annotations))}")
                            
                        annotator.set_custom_annotations(annotations)
                        annotator.annotate()
                        print(f"[DEBUG] Added {len(pvalues)} annotation(s) for groups in category {x_cat}")
                    else:
                        print(f"[DEBUG] No significant p-values to annotate for category {x_cat}")
                except Exception as e:
                    print(f"[DEBUG] Error with statannotations for groups: {e}")
                    # Fallback to a simpler method if statannotations fails
                    self._fallback_add_annotations(ax, pairs, pvalues)
                
    # The generate_statistics method has been removed as its functionality is now integrated
    # directly into the plot_graph method for a streamlined workflow.
    
    def calculate_and_store_pvals(self, df_plot, x_col, value_col, hue_col=None):
        """Backward-compatible wrapper for calculate_statistics"""
        stats_result = self.calculate_statistics(df_plot, x_col, value_col, hue_col)
        # Return the result for compatibility
        return stats_result
    def __init__(self, root):
        self.latest_pvals = {}  # {(group, h1, h2): pval or (h1, h2): pval}
        self.latest_test_info = {}  # Stores detailed test information for each p-value key
        
        # Theme settings
        self.theme_name = 'light'  # Default theme
        self.dark_mode = False  # Track dark mode state
        
        # Initialize preview_scale early so it's available when loading preferences
        # Scale factor as percentage (100 = 100% = 1x, 175 = 175% = 1.75x)
        self.preview_dpi = tk.IntVar(value=175)  # Keep variable name for compatibility, but now means scale %

        self.start_maximized_var = tk.BooleanVar(value=True)

        self.root = root
        self.version = VERSION  # Use the global VERSION constant
        self.root.title(f'ExPlot v{VERSION}')
        self.df = None           # Original raw dataframe
        self.df_work = None      # Working dataframe for plotting and statistics
        self.excel_file = None
        self.preview_label = None
        self.config_dir = self.get_config_dir()
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.custom_colors_file = str(self.config_dir / "custom_colors.json")
        self.custom_palettes_file = str(self.config_dir / "custom_palettes.json")
        self.default_settings_file = str(self.config_dir / "default_settings.json")
        self.theme_settings_file = str(self.config_dir / "theme_settings.json")
        self.models_file = str(self.config_dir / "fitting_models.json")
        self.temp_pdf = str(Path(tempfile.gettempdir()) / "explot_temp_plot.pdf")
        self.xaxis_renames = {}
        self.xaxis_order = []
        self.use_stats_var = tk.BooleanVar(value=False)
        self.black_errorbars_var = tk.BooleanVar(value=False)
        self.linewidth = tk.DoubleVar(value=1.0)
        self.errorbar_capsize_var = tk.StringVar(value="Default")  # Capsize style for error bars
        self.upward_errorbar_var = tk.BooleanVar(value=True)  # True = upward-only, False = bidirectional
        self.strip_black_var = tk.BooleanVar(value=True)
        self.show_stripplot_var = tk.BooleanVar(value=True)
        self.bar_outline_var = tk.BooleanVar(value=False)
        self.violin_inner_box_var = tk.BooleanVar(value=True)  # Show box inside violin plots
        self.bar_gap_multiplier_var = tk.DoubleVar(value=0.75)  # Multiplier for bar gap width (higher = less gap)
        self.plot_kind_var = tk.StringVar(value="bar")  # "bar", "box", "violin", or "xy"
        # Outline color setting ("as_set", "black", "gray", "white")
        self.outline_color_var = tk.StringVar(value="as_set")
        # Error bar settings
        self.errorbar_type_var = tk.StringVar(value="SD")
        # Statistics settings
        self.ttest_type_var = tk.StringVar(value="Welch's t-test (unpaired, unequal variances)")
        self.ttest_alternative_var = tk.StringVar(value="two-sided")
        self.anova_type_var = tk.StringVar(value="Welch's ANOVA")
        self.posthoc_type_var = tk.StringVar(value="Tamhane's T2")
        self.alpha_level_var = tk.StringVar(value="0.05")
        # Log scale variables
        self.xlogscale_var = tk.BooleanVar(value=False)
        self.xlog_base_var = tk.StringVar(value="10")
        self.logscale_var = tk.BooleanVar(value=False)
        self.ylog_base_var = tk.StringVar(value="10")
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
        
        # --- Legend settings ---
        self.legend_visible_var = tk.BooleanVar(value=True)
        self.legend_position_var = tk.StringVar(value="outside top")
        self.legend_ncol_var = tk.IntVar(value=0)  # 0 = auto

        self.ybreak_marker_style_var = tk.StringVar(value="Connected")
        self.ybreak_marker_style_user_set_var = tk.BooleanVar(value=False)
        
        # --- XY Fitting variables ---
        self.use_fitting_var = tk.BooleanVar(value=False)
        self.fitting_model_var = tk.StringVar(value="Linear Regression")
        self.fitting_ci_var = tk.StringVar(value="None")
        self.fitting_use_black_lines_var = tk.BooleanVar(value=False)
        self.fitting_use_black_bands_var = tk.BooleanVar(value=False)
        self.fitting_use_group_colors_var = tk.BooleanVar(value=True)
        # Initialize default models dictionary for reference
        self.default_fitting_models = {
            # Basic models
            "Linear Regression": {
                "parameters": [("A0", 1.0), ("A1", 1.0)],
                "formula": "# a simple linear model\ny = A0 + A1 * x",
                "description": "A basic linear relationship between variables. Use for data showing a constant rate of change or a straight-line trend. Common in many fields when two variables have a simple proportional relationship."
            },
            "Quadratic": {
                "parameters": [("A0", 1.0), ("A1", 1.0), ("A2", 0.1)],
                "formula": "# quadratic polynomial model\ny = A0 + A1 * x + A2 * x**2",
                "description": "Models data with one curve or inflection point. Useful for processes with acceleration/deceleration or parabolic relationships. Applications include projectile motion, cost functions, and simple optimization problems."
            },
            "Cubic": {
                "parameters": [("A0", 1.0), ("A1", 1.0), ("A2", 0.1), ("A3", 0.01)],
                "formula": "# cubic polynomial model\ny = A0 + A1 * x + A2 * x**2 + A3 * x**3",
                "description": "Models data with two possible inflection points. Useful for more complex curved relationships that change direction multiple times. Common in economics, physics, and engineering when modeling complex systems."
            },
            "Power Law": {
                "parameters": [("A", 1.0), ("b", 1.0)],
                "formula": "# power law relationship (allometric)\ny = A * x**b",
                "description": "Models scaling relationships where one variable changes as a power of another. Essential for allometric scaling in biology (e.g., body mass vs. metabolic rate). Also used in physics, economics, and network science."
            },
            "Exponential": {
                "parameters": [("A", 1.0), ("k", 0.1)],
                "formula": "# simple exponential growth/decay\ny = A * exp(k * x)",
                "description": "Models exponential growth (k > 0) or decay (k < 0) processes. Used for population growth, radioactive decay, compound interest, and many biological processes with constant growth/decay rates."
            },
            
            # Biological/biochemical models
            "Michaelis-Menten": {
                "parameters": [("Vmax", 100.0), ("Km", 10.0)],
                "formula": "# enzyme kinetics model\ny = Vmax * x / (Km + x)",
                "description": "The standard model for enzyme kinetics where reaction velocity approaches a maximum (Vmax) as substrate concentration increases. Km is the substrate concentration at half-maximum velocity. Fundamental in biochemistry and pharmaceutical research."
            },
            "Substrate Inhibition": {
                "parameters": [("Vmax", 100.0), ("Km", 10.0), ("Ki", 200.0)],
                "formula": "# enzyme kinetics with substrate inhibition\ny = Vmax * x / (Km + x + (x**2/Ki))",
                "description": "Extension of Michaelis-Menten that accounts for substrate inhibition at high concentrations. The reaction rate decreases after reaching a peak. Common in various enzyme systems where high substrate concentrations inhibit enzyme activity."
            },
            "Sigmoidal Growth": {
                "parameters": [("A", 1.0), ("k", 0.1), ("x0", 5.0)],
                "formula": "# logistic/sigmoidal growth curve\ny = A / (1 + exp(-k * (x - x0)))",
                "description": "S-shaped curve for processes with initial slow growth, rapid middle phase, and plateau. Models population growth with carrying capacity, cell growth, disease spread, and learning curves. x0 is the inflection point."
            },
            "Gompertz Growth": {
                "parameters": [("A", 1.0), ("b", 1.0), ("k", 0.1)],
                "formula": "# Gompertz growth model (tumor growth)\ny = A * exp(-b * exp(-k * x))",
                "description": "Modified growth model with asymmetric S-shape (steeper initial phase). Standard model for tumor growth dynamics. Also used for cell population growth, mortality modeling, and market penetration of technologies."
            },
            "Four-Parameter Logistic": {
                "parameters": [("A", 0.0), ("B", 1.0), ("C", 0.5), ("D", 1.0)],
                "formula": "# 4PL model for dose-response, ELISA\ny = D + (A - D) / (1 + (x / C)**B)",
                "description": "Standard model for symmetric sigmoidal dose-response curves. A is the minimum asymptote, D is the maximum, C is the inflection point (EC50/IC50), and B is the slope. Widely used in ELISA assays, pharmacology, and immunology."
            },
            "Five-Parameter Logistic": {
                "parameters": [("A", 0.0), ("B", 1.0), ("C", 0.5), ("D", 1.0), ("E", 1.0)],
                "formula": "# 5PL model for asymmetric dose-response curves\ny = D + (A - D) / (1 + (x / C)**B)**E",
                "description": "Extension of 4PL allowing for asymmetric sigmoidal curves. The additional parameter E controls asymmetry. More accurate for many biological responses where the lower and upper plateaus are approached at different rates."
            },
            "Biphasic Dose Response": {
                "parameters": [("ymin", 0.0), ("ymax1", 0.5), ("ymax2", 1.0), ("logEC50_1", 0.5), ("logEC50_2", 2.0), ("nH1", 1.0), ("nH2", 1.0)],
                "formula": "# biphasic dose response model\ny = ymin + (ymax1 - ymin) / (1 + 10**((logEC50_1 - log10(x)) * nH1)) + (ymax2 - ymin) / (1 + 10**((logEC50_2 - log10(x)) * nH2))",
                "description": "Models complex dose-response relationships with two distinct phases or peaks. Used for receptor systems with multiple binding sites or when a compound has dual effects. Common in pharmacology when drugs affect multiple receptor populations."
            },
            
            # Binding/kinetics models
            "Binding Isotherm": {
                "parameters": [("A0", 1.0), ("A1", 1.0), ("KD", 1.0)],
                "formula": "# a binding isotherm\ny = A0 + A1 * x / (x + KD)",
                "description": "Models single-site binding between a ligand and receptor. A0 is baseline, A1 is maximum binding, and KD is the dissociation constant (concentration at half-maximum binding). Fundamental in pharmacology and biochemistry."
            },
            "Hill Binding": {
                "parameters": [("A0", 1.0), ("A1", 1.0), ("KD", 1.0), ("n", 1.0)],
                "formula": "# a binding isotherm with Hill-type cooperativity\ny = A0 + A1 * x ** n / (x ** n + KD ** n)",
                "description": "Extension of binding isotherm that accounts for cooperative binding. The Hill coefficient (n) represents cooperativity: n > 1 indicates positive cooperativity, n < 1 negative cooperativity. Essential for modeling hemoglobin-oxygen binding and other cooperative systems."
            },
            "Competitive Inhibition": {
                "parameters": [("Vmax", 100.0), ("Km", 10.0), ("Ki", 50.0), ("I", 1.0)],
                "formula": "# competitive enzyme inhibition model\ny = Vmax * x / (Km * (1 + I/Ki) + x)",
                "description": "Models enzyme inhibition where inhibitor (I) competes with substrate for the active site. Ki is the inhibition constant. The apparent Km increases with inhibitor concentration while Vmax remains unchanged. Common in drug development and biochemistry."
            },
            "Non-competitive Inhibition": {
                "parameters": [("Vmax", 100.0), ("Km", 10.0), ("Ki", 50.0), ("I", 1.0)],
                "formula": "# non-competitive enzyme inhibition model\ny = Vmax * x / ((Km + x) * (1 + I/Ki))",
                "description": "Models enzyme inhibition where inhibitor binds to enzyme at a site distinct from substrate. Decreases apparent Vmax while Km remains constant. Used in enzyme studies where inhibitors alter enzyme activity without affecting substrate binding."
            },
            
            # Specialized dose-response models
            "Hormesis (U-shaped)": {
                "parameters": [("ymin", 1.0), ("ymax", 0.2), ("EC50", 100.0), ("slope", 2.0), ("stim", 0.3), ("stimEC50", 1.0)],
                "formula": "# Hormesis model for U-shaped dose-response curve\ny = ymin + (ymax - ymin) / (1 + (x/EC50)**slope) + stim / (1 + (stimEC50/x)**slope)",
                "description": "Models U-shaped or J-shaped dose-response relationships where low doses cause stimulation (hormesis) before inhibition at higher doses. Common in toxicology and environmental science where compounds can have opposite effects at different concentrations."
            },
            "EC50 with Variable Slope": {
                "parameters": [("bottom", 0.0), ("top", 1.0), ("EC50", 10.0), ("hillSlope", 1.0)],
                "formula": "# Variable slope EC50 model\ny = bottom + (top - bottom) / (1 + 10**((log10(EC50) - log10(x)) * hillSlope))",
                "description": "Standard dose-response model used to determine EC50/IC50 with variable slope. HillSlope controls steepness (efficacy) while EC50 represents potency. Preferred model for receptor-ligand interactions in pharmacology and drug discovery where slope is not assumed to be 1."
            },
            "Combination Index": {
                "parameters": [("EC50_A", 10.0), ("EC50_B", 20.0), ("h_A", 1.0), ("h_B", 1.0), ("alpha", 1.0), ("A_conc", 5.0), ("max_effect", 100.0)],
                "formula": "# Combination index model for drug interactions\ny = max_effect * (A_conc / EC50_A)**h_A / (1 + (A_conc / EC50_A)**h_A + (x / EC50_B)**h_B + alpha * (A_conc / EC50_A)**h_A * (x / EC50_B)**h_B)",
                "description": "Models interactions between two drugs, where one drug is at fixed concentration (A_conc) and the other varies (x). The alpha parameter indicates synergy (α<1), additivity (α=1), or antagonism (α>1). Essential for combination therapy design in oncology and infectious disease treatment."
            },
            "Isobologram": {
                "parameters": [("EC50_A", 10.0), ("EC50_B", 20.0), ("alpha", 1.0), ("effect_level", 0.5)],
                "formula": "# Isobologram equation for drug combinations\ny = EC50_B * (1 - x/EC50_A) / (1 - alpha * x/EC50_A)",
                "description": "Represents the concentrations of two drugs needed to achieve a specific effect level (e.g., 50% inhibition). The curve shape indicates synergy (concave), additivity (linear), or antagonism (convex). Used to design optimal drug combination strategies in pharmacology."
            },
            
            # Survival/time-to-event models
            "Weibull Survival": {
                "parameters": [("scale", 10.0), ("shape", 2.0)],
                "formula": "# Weibull survival function\ny = exp(-(x/scale)**shape)",
                "description": "Models survival probability as a function of time. The shape parameter determines whether hazard increases (>1), decreases (<1), or remains constant (=1) over time. Widely used in reliability engineering, medical survival analysis, and lifetime testing due to its flexibility."
            },
            "Exponential Survival": {
                "parameters": [("lambda", 0.05)],
                "formula": "# Exponential survival function\ny = exp(-lambda * x)",
                "description": "Models survival with constant hazard rate (lambda) over time. Simplest survival model, assuming risk of failure/death doesn't change with time. Used as a baseline in clinical trials and for modeling events that occur randomly without aging/wear effects."
            },
            "Log-logistic Survival": {
                "parameters": [("alpha", 10.0), ("beta", 2.0)],
                "formula": "# Log-logistic survival function\ny = 1 / (1 + (x/alpha)**beta)",
                "description": "Models survival with non-monotonic hazard rates (initially increasing, then decreasing). Used for event times that first become more likely then less likely over time. Common in pharmacokinetics and for modeling time-to-progression in diseases with initial crisis followed by stabilization."
            },
            "Gompertz-Makeham": {
                "parameters": [("a", 0.01), ("b", 0.1), ("c", 0.001)],
                "formula": "# Gompertz-Makeham survival model\ny = exp(-(c*x + a/b * (exp(b*x) - 1)))",
                "description": "Extends Gompertz model by adding age-independent mortality component (c). Models human mortality combining background risk (c) with age-dependent risk (a,b). Standard in actuarial science, demography, and gerontology for human lifespan modeling."
            },
            
            # Signal/peak models
            "Gaussian": {
                "parameters": [("A", 1.0), ("mu", 5.0), ("sigma", 1.0), ("y0", 0.0)],
                "formula": "# Gaussian/normal distribution peak\ny = y0 + A * exp(-(x - mu)**2 / (2 * sigma**2))",
                "description": "Bell-shaped curve for modeling normally distributed data or peaks. mu is the center, sigma is the width, A is amplitude, and y0 is baseline. Used for chromatography peaks, spectral lines, and many natural distributions in biology and chemistry."
            },
            "Lorentzian": {
                "parameters": [("A", 1.0), ("x0", 5.0), ("gamma", 1.0), ("y0", 0.0)],
                "formula": "# Lorentzian peak (spectroscopy)\ny = y0 + A * gamma**2 / ((x - x0)**2 + gamma**2)",
                "description": "Models peaks with wider 'tails' than Gaussian. Common in spectroscopy, particularly for modeling spectral line shapes in NMR, IR, and other resonance phenomena. x0 is center, gamma controls width, and A is amplitude."
            },
            "Boltzmann Sigmoid": {
                "parameters": [("A1", 0.0), ("A2", 1.0), ("x0", 5.0), ("dx", 1.0)],
                "formula": "# Boltzmann sigmoid for transitions\ny = A2 + (A1 - A2) / (1 + exp((x - x0) / dx))",
                "description": "Models sharp transitions between two states. Used for voltage-dependent channel activation in electrophysiology, phase transitions, and other threshold phenomena. A1 and A2 are the asymptotes, x0 is the midpoint, and dx controls steepness."
            },
            
            # Periodic/oscillation models
            "Sine Wave": {
                "parameters": [("A", 1.0), ("f", 0.1), ("phi", 0.0), ("y0", 0.0)],
                "formula": "# Sine wave oscillation\ny = y0 + A * sin(2 * pi * f * x + phi)",
                "description": "Basic model for oscillatory processes. A is amplitude, f is frequency, phi is phase shift, and y0 is vertical offset. Used for modeling circadian rhythms, seasonal cycles, sound waves, electrical oscillations, and other periodic phenomena."
            },
            "Damped Oscillation": {
                "parameters": [("A", 1.0), ("f", 0.1), ("phi", 0.0), ("lambda", 0.05), ("y0", 0.0)],
                "formula": "# Damped oscillation\ny = y0 + A * exp(-lambda * x) * sin(2 * pi * f * x + phi)",
                "description": "Models oscillations that decay over time. Lambda controls damping rate. Used for systems returning to equilibrium like pendulums with friction, RLC circuits, population oscillations, and mechanical resonance with damping."
            },
            
            # Exponential models
            "Single Exponential": {
                "parameters": [("A0", 1.0), ("A1", 1.0), ("k1", 1.0)],
                "formula": "# a single exponential function\ny = A0 + A1 * exp(-k1 * x)",
                "description": "Models processes with one characteristic decay/growth rate. A0 is baseline, A1 is amplitude, and k1 is rate constant. Used for simple radioactive decay, material cooling, drug clearance, and first-order chemical reactions."
            },
            "Double Exponential": {
                "parameters": [("A0", 1.0), ("A1", -1.0), ("k1", 1.0), ("A2", 1.0), ("k2", 5.0)],
                "formula": "# a double exponential function\ny = A0 + A1 * exp(-k1 * x) + A2 * exp(-k2 * x)",
                "description": "Models processes with two different decay/growth rates. Used for biexponential drug clearance, complex fluorescence decay, hormone release, and systems with two distinct compartments or kinetic processes."
            },
            "Triple Exponential": {
                "parameters": [("A0", 1.0), ("A1", -1.0), ("k1", 1.0), ("A2", 1.0), ("k2", 5.0), ("A3", -2.0), ("k3", 1.2)],
                "formula": "# a triple exponential function\ny = A0 + A1 * exp(-k1 * x) + A2 * exp(-k2 * x) + A3 * exp(-k3 * x)",
                "description": "Models complex systems with three different time scales or compartments. Used for multicompartment pharmacokinetic models, complex decay processes in physics, and systems with multiple parallel or sequential processes with different rates."
            }
        }
        
        # Load saved models or initialize with defaults
        self.fitting_models = self.load_fitting_models()
        
        # If no saved models exist, initialize with defaults
        if not self.fitting_models:
            self.fitting_models = self.default_fitting_models.copy()
            self.save_fitting_models()
        
        self.load_custom_colors_palettes()
        
        # Initialize default plot dimensions for the plot area
        self.plot_width_var = tk.DoubleVar(value=1.5)
        self.plot_height_var = tk.DoubleVar(value=1.5)
        
        # Initialize color variables *before* loading preferences
        self.single_color_var = tk.StringVar(value=list(self.custom_colors.keys())[0] if self.custom_colors else "")
        self.palette_var = tk.StringVar(value=list(self.custom_palettes.keys())[0] if self.custom_palettes else "")
        
        self.load_user_preferences()
        self.setup_menu()
        self.setup_ui()
        self.setup_statistics_settings_tab()

    def pval_to_annotation(self, pval, alpha=None):
        """Return p-value in scientific notation.
        Returns '?' if pval is None or NaN, otherwise returns p-value in scientific notation."""
        if pval is None or (isinstance(pval, float) and math.isnan(pval)):
            return "?"
            
        # Format p-value in scientific notation with 2 decimal places
        # Use 'e' format for scientific notation with 2 decimal places
        return f"{pval:.2e}"

    def format_pvalue_matrix(self, matrix):
        # Check if the matrix is empty
        if matrix is None or matrix.empty:
            return "No data available"
        
        # Create a copy of the matrix with string dtype to avoid warnings
        formatted = pd.DataFrame(index=matrix.index, columns=matrix.columns, dtype=str)
        
        # Format the values in the matrix
        for idx in matrix.index:
            for col in matrix.columns:
                val = matrix.loc[idx, col]
                if idx == col or (isinstance(val, float) and (val == 1.0 or pd.isna(val))):
                    formatted.loc[idx, col] = "—"  # Diagonal, NaN, or 1.0 values
                elif isinstance(val, (int, float, np.number)):
                    # Handle extremely small values that might be 0 in floating point precision
                    if val < 0.0001:
                        formatted.loc[idx, col] = f"{val:.2e}"  # Scientific notation for very small values
                    else:
                        formatted.loc[idx, col] = f"{val:.4f}"  # 4 decimal places for other values
                else:
                    formatted.loc[idx, col] = str(val)
        
        # Get column names and calculate widths
        col_names = formatted.columns.tolist()
        col_widths = [
            max(len(str(col)), max((len(str(formatted.loc[idx, col])) for idx in formatted.index), default=0))
            for col in col_names
        ]
        idx_width = max((len(str(idx)) for idx in formatted.index), default=0)
        
        # Build header
        header = " " * (idx_width + 2) + "| " + " | ".join(
            f"{col:^{w}}" for col, w in zip(col_names, col_widths)
        ) + " |"
        
        # Build separator
        sep = "-" * (idx_width + 2) + "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
        
        # Build rows
        rows = []
        for idx in formatted.index:
            row = f" {str(idx):<{idx_width}} | " + " | ".join(
                f"{str(formatted.loc[idx, col]):^{w}}" for col, w in zip(col_names, col_widths)
            ) + " |"
            rows.append(row)
            
        return "\n".join([header, sep] + rows)
    def setup_statistics_settings_tab(self):
        frame = self.stats_settings_tab
        
        # Add stats info button at the top
        info_frame = ttk.Frame(frame)
        info_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(4, 12))
        
        ttk.Label(info_frame, text="Statistical Tests", font=(None, 12, 'bold')).pack(side='left', padx=8)
        info_button = ttk.Button(info_frame, text="ℹ️ Info", command=self.show_stats_info)
        info_button.pack(side='right', padx=8)
        
        # t-test type
        ttk.Label(frame, text="t-test type:").grid(row=1, column=0, sticky="w", padx=8, pady=8)
        # Use the existing ttest_type_var variable initialized in __init__
        ttest_options = [
            "Student's t-test (unpaired, equal variances)",
            "Welch's t-test (unpaired, unequal variances)",
            "Paired t-test",
            "Mann-Whitney U test (non-parametric)",
            "Wilcoxon signed-rank test (non-parametric)"
        ]
        ttest_dropdown = ttk.Combobox(frame, textvariable=self.ttest_type_var, values=ttest_options, state='readonly', width=30)
        ttest_dropdown.grid(row=1, column=1, sticky="ew", padx=8, pady=8)
        
        # T-test alternative hypothesis
        ttk.Label(frame, text="T-test Alternative:").grid(row=2, column=0, sticky="w", padx=8, pady=8)
        # Use the existing ttest_alternative_var variable initialized in __init__
        ttest_alternative_options = [
            "two-sided",
            "less",
            "greater"
        ]
        ttest_alternative_dropdown = ttk.Combobox(frame, textvariable=self.ttest_alternative_var, values=ttest_alternative_options, state='readonly', width=30)
        ttest_alternative_dropdown.grid(row=2, column=1, sticky="ew", padx=8, pady=8)
        
        # ANOVA type
        ttk.Label(frame, text="ANOVA type:").grid(row=3, column=0, sticky="w", padx=8, pady=8)
        # Use the existing anova_type_var variable initialized in __init__
        anova_options = [
            "One-way ANOVA",
            "Welch's ANOVA",
            "Repeated measures ANOVA",
            "Kruskal-Wallis H test (non-parametric)",
            "Friedman test (non-parametric)"
        ]
        anova_dropdown = ttk.Combobox(frame, textvariable=self.anova_type_var, values=anova_options, state='readonly', width=30)
        anova_dropdown.grid(row=3, column=1, sticky="ew", padx=8, pady=8)
        
        # Alpha level
        ttk.Label(frame, text="Alpha level:").grid(row=4, column=0, sticky="w", padx=8, pady=8)
        alpha_options = ["0.05", "0.01", "0.001", "0.0001"]
        alpha_dropdown = ttk.Combobox(frame, textvariable=self.alpha_level_var, values=alpha_options, state='readonly', width=30)
        alpha_dropdown.grid(row=4, column=1, sticky="ew", padx=8, pady=8)
        
        # Post-hoc test
        ttk.Label(frame, text="Post-hoc test:").grid(row=5, column=0, sticky="w", padx=8, pady=8)
        # Use the existing posthoc_type_var variable initialized in __init__
        posthoc_options = [
            "Tukey's HSD",
            "Tamhane's T2",
            "Scheffe's test",
            "Dunn's test",
            "Conover's test (non-parametric)",
            "Nemenyi test (non-parametric)"
        ]
        posthoc_dropdown = ttk.Combobox(frame, textvariable=self.posthoc_type_var, values=posthoc_options, state='readonly', width=30)
        posthoc_dropdown.grid(row=5, column=1, sticky="ew", padx=8, pady=8)


    def get_config_dir(self):
        """Return the user config directory for settings (cross-platform)."""
        if sys.platform == "darwin":
            return Path.home() / "Library" / "Application Support" / "ExPlot"
        elif sys.platform.startswith("win"):
            return Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "ExPlot"
        else:
            # Linux and other
            return Path.home() / ".config" / "ExPlot"

    def setup_menu(self):
        """Setup the application menu."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel File", command=self.open_file)
        file_menu.add_command(label="Load Example Data", command=self.load_example_data)
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Save PNG", command=lambda: self.save_graph('png'))
        file_menu.add_command(label="Save PDF", command=lambda: self.save_graph('pdf'))
        file_menu.add_separator()
        file_menu.add_command(label="Save Project", command=self.save_project)
        file_menu.add_command(label="Load Project", command=self.load_project)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Options menu
        option_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Options", menu=option_menu)
        option_menu.add_command(label="Default Settings", command=self.show_settings)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Package Information", command=self.show_package_info)

    def save_project(self):
        """Save current plot configuration to a project file."""
        # Check if we have data to save
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data before saving a project.")
            return
            
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".explt",
            filetypes=[("ExPlot Projects", "*.explt"), ("All files", "*.*")],
            title="Save Project"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Convert dataframe to JSON-serializable format
            # We need to convert the DataFrame into a nested dictionary/list structure 
            # that can be serialized to JSON
            
            # First convert to dict with records orientation
            df_records = self.df.to_dict(orient='records')
            
            # Get the columns list
            df_columns = list(self.df.columns)
            
            # Get index if it's not the default RangeIndex
            df_index = None
            if not isinstance(self.df.index, pd.RangeIndex):
                df_index = self.df.index.tolist()
                    
            # Create project dictionary with all relevant settings
            project = {
                'version': VERSION,
                'timestamp': pd.Timestamp.now().isoformat(),
                'excel_file': self.current_excel_file if hasattr(self, 'current_excel_file') else None,
                'sheet_name': self.selected_sheet.get() if hasattr(self, 'selected_sheet') else None,
                
                # Save the data for offline use
                'data': {
                    'columns': df_columns,
                    'records': df_records,
                    'index': df_index
                },
                
                # Plot configuration
                'plot_kind': self.plot_kind_var.get(),
                'x_column': self.xaxis_var.get() if hasattr(self, 'xaxis_var') else None,
                'y_column': self.selected_y_col.get() if hasattr(self, 'selected_y_col') else None,
                'hue_column': self.group_var.get() if hasattr(self, 'group_var') and self.group_var.get() != 'None' else None,
                
                # Value columns selected in the Basic tab
                'selected_columns': [col for var, col in self.value_vars if var.get()] if hasattr(self, 'value_vars') else [],
                
                # X-axis label renaming and reordering
                'xaxis_renames': self.xaxis_renames if hasattr(self, 'xaxis_renames') else {},
                'xaxis_order': self.xaxis_order if hasattr(self, 'xaxis_order') else [],
                'use_stats': self.use_stats_var.get(),
                'errorbar_type': self.errorbar_type_var.get(),
                'errorbar_black': self.errorbar_black_var.get() if hasattr(self, 'errorbar_black_var') else True,
                'show_stripplot': self.show_stripplot_var.get(),
                'strip_black': self.strip_black_var.get() if hasattr(self, 'strip_black_var') else True,
                
                # Statistics settings
                'ttest_type': self.ttest_type_var.get(),
                'ttest_alternative': self.ttest_alternative_var.get(),
                'anova_type': self.anova_type_var.get(),
                'posthoc_type': self.posthoc_type_var.get(),
                'alpha_level': self.alpha_level_var.get(),
                
                # Plot appearance
                'plot_width': self.plot_width_var.get(),
                'plot_height': self.plot_height_var.get(),
                'bar_gap_multiplier': self.bar_gap_multiplier_var.get() if hasattr(self, 'bar_gap_multiplier_var') else 0.8,
                'xlogscale': self.xlogscale_var.get() if hasattr(self, 'xlogscale_var') else False,
                'xlog_base': self.xlog_base_var.get() if hasattr(self, 'xlog_base_var') else "10",
                'logscale': self.logscale_var.get() if hasattr(self, 'logscale_var') else False,
                'ylog_base': self.ylog_base_var.get() if hasattr(self, 'ylog_base_var') else "10",
                
                # XY plot settings
                'xy_marker_symbol': self.xy_marker_symbol_var.get() if hasattr(self, 'xy_marker_symbol_var') else "o",
                'xy_marker_size': self.xy_marker_size_var.get() if hasattr(self, 'xy_marker_size_var') else 5,
                'xy_filled': self.xy_filled_var.get() if hasattr(self, 'xy_filled_var') else True,
                'xy_line_style': self.xy_line_style_var.get() if hasattr(self, 'xy_line_style_var') else "solid",
                'xy_line_black': self.xy_line_black_var.get() if hasattr(self, 'xy_line_black_var') else False,
                'xy_connect': self.xy_connect_var.get() if hasattr(self, 'xy_connect_var') else False,
                'xy_show_mean': self.xy_show_mean_var.get() if hasattr(self, 'xy_show_mean_var') else True,
                'xy_show_mean_errorbars': self.xy_show_mean_errorbars_var.get() if hasattr(self, 'xy_show_mean_errorbars_var') else True,
                'xy_draw_band': self.xy_draw_band_var.get() if hasattr(self, 'xy_draw_band_var') else False,
                
                # Plot titles/labels
                'plot_title': self.plot_title_var.get() if hasattr(self, 'plot_title_var') else "",
                'x_axis_label': self.x_axis_label_var.get() if hasattr(self, 'x_axis_label_var') else "",
                'y_axis_label': self.y_axis_label_var.get() if hasattr(self, 'y_axis_label_var') else ""
            }
            
            # Save to file
            with open(file_path, 'w') as f:
                json.dump(project, f, indent=2)
                
            messagebox.showinfo("Success", f"Project saved to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save project: {str(e)}")
            
    def load_project(self):
        """Load a saved project file."""
        file_path = filedialog.askopenfilename(
            filetypes=[("ExPlot Projects", "*.explt"), ("All files", "*.*")],
            title="Load Project"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Load project file
            with open(file_path, 'r') as f:
                project = json.load(f)
                
            # Check version compatibility
            file_version = project.get('version', '0.0.0')
            if file_version.split('.')[0] != VERSION.split('.')[0]:
                if not messagebox.askyesno("Version Mismatch", 
                                      f"Project was created with version {file_version} and may not be compatible with current version {VERSION}. Continue anyway?"):
                    return
            
            # Check if we have embedded data in the project
            has_embedded_data = 'data' in project and project['data'] is not None
            
            # Check if the original Excel file exists
            excel_file = project.get('excel_file')
            original_file_exists = excel_file and os.path.exists(excel_file)
            
            # Always use embedded data if available
            use_original_file = False
            
            if has_embedded_data:
                # Proceed with embedded data without asking
                use_original_file = False
            elif original_file_exists:
                # Only use original file if no embedded data is available
                use_original_file = True
            else:
                # No data source available
                messagebox.showwarning(
                    "No Data Source", 
                    "Neither embedded data nor the original Excel file are available. "
                    "Please select a new data source."
                )
                if self.open_file():
                    # If user successfully selected a new file, continue with that
                    use_original_file = True
                else:
                    # User cancelled file selection
                    return
                
            # Load data from appropriate source
            if use_original_file:
                # Use original Excel file
                self.process_excel_file(excel_file)
                
                # Select the right sheet
                sheet_name = project.get('sheet_name')
                if sheet_name and hasattr(self, 'selected_sheet') and sheet_name in self.sheet_options:
                    self.selected_sheet.set(sheet_name)
                    self.on_sheet_selected()
            elif has_embedded_data:
                # Use embedded data
                try:
                    # Reconstruct DataFrame from embedded data
                    data_info = project['data']
                    columns = data_info.get('columns', [])
                    
                    # Get data records - these could be stored in different formats depending on version
                    records = data_info.get('records')
                    if records is None:
                        # Try legacy format
                        legacy_data = data_info.get('data', {})
                        if isinstance(legacy_data, dict):
                            # Handle column-oriented dict format
                            df_data = {}
                            for col in columns:
                                if col in legacy_data:
                                    df_data[col] = legacy_data[col]
                                else:
                                    # Handle missing columns
                                    df_data[col] = [None] * (len(list(legacy_data.values())[0]) if legacy_data else 0)
                            self.df = pd.DataFrame(df_data)
                        elif isinstance(legacy_data, list):
                            # Handle list format
                            self.df = pd.DataFrame(legacy_data, columns=columns)
                        else:
                            # Cannot process this format
                            raise ValueError(f"Unrecognized data format in project file")
                    else:
                        # Handle standard records format
                        self.df = pd.DataFrame.from_records(records)
                    
                    # Restore index if available
                    index = data_info.get('index')
                    if index is not None:
                        self.df.index = index
                        
                    # Set up UI for embedded data
                    self.current_excel_file = "[Embedded Data]"  # Mark as embedded data
                    self.excel_file = "[Embedded Data]"
                    
                    # Create mock sheet selection for embedded data
                    self.sheet_options = ['EmbeddedData']
                    if hasattr(self, 'sheet_dropdown'):
                        self.sheet_dropdown['values'] = self.sheet_options
                    
                    if hasattr(self, 'selected_sheet'):
                        self.selected_sheet.set('EmbeddedData')
                    elif hasattr(self, 'sheet_var'):
                        self.sheet_var.set('EmbeddedData')
                    
                    # Update columns in UI
                    self.update_columns()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load embedded data: {str(e)}")
                    return
            
            # Set column selections
            x_column = project.get('x_column')
            y_column = project.get('y_column')
            hue_column = project.get('hue_column')
            selected_columns = project.get('selected_columns', [])
            
            # Set X-axis column
            if x_column and hasattr(self, 'xaxis_var') and x_column in self.df.columns:
                self.xaxis_var.set(x_column)
                
            # Set Y-axis column (if applicable)
            if y_column and hasattr(self, 'selected_y_col') and y_column in self.df.columns:
                self.selected_y_col.set(y_column)
                
            # Set group/hue column
            if hasattr(self, 'group_var'):
                if hue_column and hue_column in self.df.columns:
                    self.group_var.set(hue_column)
                else:
                    self.group_var.set('None')
                
            # Restore the selected value columns in the Basic tab
            if hasattr(self, 'value_vars') and selected_columns:
                # Set checkboxes for the selected columns that exist in the current dataframe
                existing_columns = set(self.df.columns)
                for var, col in self.value_vars:
                    if col in selected_columns and col in existing_columns:
                        var.set(True)
                    else:
                        var.set(False)
                        
            # Restore X-axis label renaming and reordering
            if hasattr(self, 'xaxis_renames') and 'xaxis_renames' in project:
                # Load the saved renames dictionary
                saved_renames = project.get('xaxis_renames', {})
                if saved_renames:
                    # Convert keys to appropriate types if needed (JSON serialization may have converted numeric keys to strings)
                    self.xaxis_renames = {}
                    for k, v in saved_renames.items():
                        # Try to convert numeric strings back to numbers if they were originally numbers
                        try:
                            if '.' in k:
                                # Try as float
                                num_key = float(k)
                                self.xaxis_renames[num_key] = v
                            else:
                                # Try as integer
                                num_key = int(k)
                                self.xaxis_renames[num_key] = v
                        except ValueError:
                            # Not a number, use as is
                            self.xaxis_renames[k] = v
            
            if hasattr(self, 'xaxis_order') and 'xaxis_order' in project:
                # Load the saved order list
                self.xaxis_order = project.get('xaxis_order', [])
                
            # Apply plot settings
            if 'plot_kind' in project and hasattr(self, 'plot_kind_var'):
                self.plot_kind_var.set(project['plot_kind'])
                
            if 'use_stats' in project and hasattr(self, 'use_stats_var'):
                self.use_stats_var.set(project['use_stats'])
                
            if 'errorbar_type' in project and hasattr(self, 'errorbar_type_var'):
                self.errorbar_type_var.set(project['errorbar_type'])
                
            if 'errorbar_black' in project and hasattr(self, 'errorbar_black_var'):
                self.errorbar_black_var.set(project['errorbar_black'])
                
            if 'show_stripplot' in project and hasattr(self, 'show_stripplot_var'):
                self.show_stripplot_var.set(project['show_stripplot'])
                
            if 'strip_black' in project and hasattr(self, 'strip_black_var'):
                self.strip_black_var.set(project['strip_black'])
                
            # Statistics settings
            if 'ttest_type' in project and hasattr(self, 'ttest_type_var'):
                self.ttest_type_var.set(project['ttest_type'])
                
            if 'ttest_alternative' in project and hasattr(self, 'ttest_alternative_var'):
                self.ttest_alternative_var.set(project['ttest_alternative'])
                
            if 'anova_type' in project and hasattr(self, 'anova_type_var'):
                self.anova_type_var.set(project['anova_type'])
                
            if 'posthoc_type' in project and hasattr(self, 'posthoc_type_var'):
                self.posthoc_type_var.set(project['posthoc_type'])
                
            if 'alpha_level' in project and hasattr(self, 'alpha_level_var'):
                self.alpha_level_var.set(project['alpha_level'])
                
            # Plot appearance
            if 'plot_width' in project and hasattr(self, 'plot_width_var'):
                self.plot_width_var.set(project['plot_width'])
                
            if 'plot_height' in project and hasattr(self, 'plot_height_var'):
                self.plot_height_var.set(project['plot_height'])
                
            if 'bar_gap_multiplier' in project and hasattr(self, 'bar_gap_multiplier_var'):
                self.bar_gap_multiplier_var.set(project['bar_gap_multiplier'])
                
            if 'xlogscale' in project and hasattr(self, 'xlogscale_var'):
                self.xlogscale_var.set(project['xlogscale'])
                
            if 'xlog_base' in project and hasattr(self, 'xlog_base_var'):
                self.xlog_base_var.set(project['xlog_base'])
                
            if 'logscale' in project and hasattr(self, 'logscale_var'):
                self.logscale_var.set(project['logscale'])
                
            if 'ylog_base' in project and hasattr(self, 'ylog_base_var'):
                self.ylog_base_var.set(project['ylog_base'])
            
            # XY plot settings
            if 'xy_marker_symbol' in project and hasattr(self, 'xy_marker_symbol_var'):
                self.xy_marker_symbol_var.set(project['xy_marker_symbol'])
                
            if 'xy_marker_size' in project and hasattr(self, 'xy_marker_size_var'):
                self.xy_marker_size_var.set(project['xy_marker_size'])
                
            if 'xy_filled' in project and hasattr(self, 'xy_filled_var'):
                self.xy_filled_var.set(project['xy_filled'])
                
            if 'xy_line_style' in project and hasattr(self, 'xy_line_style_var'):
                self.xy_line_style_var.set(project['xy_line_style'])
                
            if 'xy_line_black' in project and hasattr(self, 'xy_line_black_var'):
                self.xy_line_black_var.set(project['xy_line_black'])
                
            if 'xy_connect' in project and hasattr(self, 'xy_connect_var'):
                self.xy_connect_var.set(project['xy_connect'])
                
            if 'xy_show_mean' in project and hasattr(self, 'xy_show_mean_var'):
                self.xy_show_mean_var.set(project['xy_show_mean'])
                
            if 'xy_show_mean_errorbars' in project and hasattr(self, 'xy_show_mean_errorbars_var'):
                self.xy_show_mean_errorbars_var.set(project['xy_show_mean_errorbars'])
                
            if 'xy_draw_band' in project and hasattr(self, 'xy_draw_band_var'):
                self.xy_draw_band_var.set(project['xy_draw_band'])
                
            # Plot titles/labels
            if 'plot_title' in project and hasattr(self, 'plot_title_var'):
                self.plot_title_var.set(project['plot_title'])
                
            if 'x_axis_label' in project and hasattr(self, 'x_axis_label_var'):
                self.x_axis_label_var.set(project['x_axis_label'])
                
            if 'y_axis_label' in project and hasattr(self, 'y_axis_label_var'):
                self.y_axis_label_var.set(project['y_axis_label'])
                
            # Generate the plot
            if hasattr(self, 'plot_button'):
                self.plot_button.invoke()
                
            messagebox.showinfo("Success", "Project loaded successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project: {str(e)}")
            
    def process_excel_file(self, file_path):
        try:
            # Store file paths
            self.excel_file = file_path
            self.current_excel_file = file_path  # Store path for project saving/loading
            
            if file_path.lower().endswith('.csv'):
                # Handle CSV files
                self.df = pd.read_csv(file_path, dtype=object)
                # Create a single sheet for CSV
                self.sheet_options = ['Sheet1']
                self.sheet_var.set('Sheet1')
                
                # Initialize both raw and modified dataframes
                self.raw_df = self.df.copy()
                self.modified_df = self.df.copy()
                
                # Initialize modification tracking variables
                self.xaxis_renames = {}
                self.excluded_x_values = set()
                self.xaxis_order = []
                self.current_sheet_type = 'external'
                
                # Update columns
                self.update_columns(reset_labels=True)
            else:
                # Handle Excel files
                xls = pd.ExcelFile(file_path)
                all_sheets = xls.sheet_names
                
                # Filter out sheets that start with underscore
                visible_sheets = [sheet for sheet in all_sheets if not str(sheet).startswith('_')]
                self.sheet_options = visible_sheets
                
                # Add special options for raw and modified data
                if hasattr(self, 'raw_df') and self.raw_df is not None:
                    self.sheet_options = ['Raw Embedded Data', 'Modified Embedded Data'] + self.sheet_options
                
                # Update the sheet dropdown
                if hasattr(self, 'sheet_dropdown'):
                    self.sheet_dropdown['values'] = self.sheet_options
                
                # Set the sheet (prefer 'export' if it exists)
                if "export" in self.sheet_options:
                    sheet_name = "export"
                else:
                    sheet_name = self.sheet_options[0]
                
                # Set the sheet variable
                self.sheet_var.set(sheet_name)
                
                # Load the selected sheet
                self.load_sheet()
                
        except Exception as e:
            raise Exception(f"Error processing file: {str(e)}")
    
    def open_file(self):
        """Open an Excel file through a file dialog.
        
        Returns:
            bool: True if a file was successfully loaded, False otherwise.
        """
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")],
            title="Open Data File"
        )
        
        if file_path:
            try:
                self.process_excel_file(file_path)
                return True
            except Exception as e:
                messagebox.showerror("Error", f"Failed to open file: {str(e)}")
                return False
        return False
                
    def load_example_data(self):
        """Load the example data file included with the application."""
        # Construct the path to the example_data.xlsx file
        script_dir = os.path.dirname(os.path.abspath(__file__))
        example_data_path = os.path.join(script_dir, "example_data.xlsx")
        
        try:
            print(f"Attempting to load example data from: {example_data_path}")
            if os.path.exists(example_data_path):
                print("Example data file exists, processing...")
                try:
                    self.process_excel_file(example_data_path)
                    print("Example data loaded successfully")
                except Exception as e:
                    error_msg = f"Error processing example data file: {str(e)}"
                    print(error_msg)
                    messagebox.showerror("Error", error_msg)
            else:
                error_msg = f"Example data file not found at {example_data_path}"
                print(error_msg)
                messagebox.showerror("Error", error_msg)
        except Exception as e:
            error_msg = f"Failed to load example data: {str(e)}"
            print(error_msg)
            messagebox.showerror("Error", error_msg)
            
    def export_to_excel(self):
        """Export the current data to an Excel file, preserving any modifications.
        This includes reordering or renaming of x values.
        """
        if not hasattr(self, 'df') or self.df is None or self.df.empty:
            messagebox.showerror("Error", "No data to export. Please load data first.")
            return
            
        # Ask the user for a file to save to
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Export Data to Excel"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Export the DataFrame to Excel
            self.df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Data successfully exported to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def save_graph(self, file_format='png'):
        """Save the current graph as an image file.
        
        Args:
            file_format (str): The format to save the graph in ('png' or 'pdf')
        """
        if not hasattr(self, 'fig') or self.fig is None:
            messagebox.showerror("Error", "No graph to save. Please generate a graph first.")
            return
            
        # Set up file dialog options based on format
        if file_format == 'pdf':
            defaultextension = ".pdf"
            filetypes = [("PDF files", "*.pdf"), ("All files", "*.*")]
            title = "Save as PDF"
        else:  # Default to PNG
            defaultextension = ".png"
            filetypes = [("PNG files", "*.png"), ("All files", "*.*")]
            title = "Save as PNG"
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=defaultextension,
            filetypes=filetypes,
            title=title
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Ensure the file has the correct extension
            if not file_path.lower().endswith(f".{file_format}"):
                file_path = f"{os.path.splitext(file_path)[0]}.{file_format}"
            
            # Save with appropriate settings based on format
            if file_format == 'pdf':
                self.fig.savefig(file_path, format='pdf', bbox_inches='tight')
                message = f"PDF saved to {file_path}"
            else:  # PNG
                self.fig.savefig(file_path, format='png', dpi=300, bbox_inches='tight')
                message = f"PNG image saved to {file_path}"
                
            messagebox.showinfo("Success", message)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save graph: {str(e)}")

    def show_settings(self):
        window = tk.Toplevel(self.root)
        window.title("Default Settings")
        window.geometry("600x550")
        
        # Create notebook for tabs
        notebook = ttk.Notebook(window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        general_tab = ttk.Frame(notebook)
        plot_settings_tab = ttk.Frame(notebook)
        stats_tab = ttk.Frame(notebook)
        appearance_tab = ttk.Frame(notebook)
        bar_graph_tab = ttk.Frame(notebook)
        xy_plot_tab = ttk.Frame(notebook)
        colors_tab = ttk.Frame(notebook)
        
        notebook.add(general_tab, text='General')
        notebook.add(plot_settings_tab, text='Plot Settings')
        notebook.add(stats_tab, text='Statistics')
        notebook.add(appearance_tab, text='Appearance')
        notebook.add(bar_graph_tab, text='Bar Graph')
        notebook.add(xy_plot_tab, text='XY Plot')
        notebook.add(colors_tab, text='Colors')
        
        # Variables to hold settings
        # General tab
        self.settings_plot_kind_var = tk.StringVar(value=self.plot_kind_var.get())
        self.settings_start_maximized_var = tk.BooleanVar(value=self.start_maximized_var.get() if hasattr(self, 'start_maximized_var') else True)
        
        # Plot Settings tab
        self.settings_show_stripplot_var = tk.BooleanVar(value=self.show_stripplot_var.get())
        self.settings_strip_black_var = tk.BooleanVar(value=self.strip_black_var.get())
        self.settings_errorbar_type_var = tk.StringVar(value=self.errorbar_type_var.get())
        self.settings_errorbar_black_var = tk.BooleanVar(value=self.errorbar_black_var.get())
        self.settings_errorbar_capsize_var = tk.StringVar(value=self.errorbar_capsize_var.get())
        
        # Bar Graph tab
        self.settings_bar_outline_var = tk.BooleanVar(value=self.bar_outline_var.get())
        self.settings_upward_errorbar_var = tk.BooleanVar(value=self.upward_errorbar_var.get())
        
        # Statistics tab
        self.settings_use_stats_var = tk.BooleanVar(value=self.use_stats_var.get())
        self.settings_ttest_type_var = tk.StringVar(value=self.ttest_type_var.get())
        self.settings_ttest_alternative_var = tk.StringVar(value=self.ttest_alternative_var.get())
        self.settings_anova_type_var = tk.StringVar(value=self.anova_type_var.get())
        self.settings_alpha_level_var = tk.StringVar(value=self.alpha_level_var.get())
        self.settings_posthoc_type_var = tk.StringVar(value=self.posthoc_type_var.get())
        
        # Bar Graph tab
        self.settings_bar_gap_multiplier_var = tk.DoubleVar(value=self.bar_gap_multiplier_var.get())
        
        # Appearance tab
        self.settings_linewidth = tk.DoubleVar(value=self.linewidth.get())
        self.settings_plot_width_var = tk.DoubleVar(value=self.plot_width_var.get())
        self.settings_plot_height_var = tk.DoubleVar(value=self.plot_height_var.get())
        self.settings_preview_dpi_var = tk.IntVar(value=self.preview_dpi.get() if hasattr(self, 'preview_dpi') else 175)

        try:
            self.settings_ybreak_marker_style_var = tk.StringVar(value=self.ybreak_marker_style_var.get())
        except Exception:
            self.settings_ybreak_marker_style_var = tk.StringVar(value="Connected")

        try:
            gap_val = 0.07
            if hasattr(self, 'ybreak_gap_entry'):
                gap_str = self.ybreak_gap_entry.get().strip()
                if gap_str:
                    if '-' in gap_str and not gap_str.startswith('-'):
                        gap_str = gap_str.split('-')[0]
                    gap_val = float(gap_str)
            gap_val = max(0.0, min(0.3, gap_val))
        except Exception:
            gap_val = 0.07
        self.settings_ybreak_gap_var = tk.DoubleVar(value=gap_val)
        
        # XY Plot tab
        self.settings_xy_marker_symbol_var = tk.StringVar(value=self.xy_marker_symbol_var.get())
        self.settings_xy_marker_size_var = tk.DoubleVar(value=self.xy_marker_size_var.get())
        self.settings_xy_filled_var = tk.BooleanVar(value=self.xy_filled_var.get())
        self.settings_xy_line_style_var = tk.StringVar(value=self.xy_line_style_var.get())
        self.settings_xy_line_black_var = tk.BooleanVar(value=self.xy_line_black_var.get())
        self.settings_xy_connect_var = tk.BooleanVar(value=self.xy_connect_var.get())
        self.settings_xy_show_mean_var = tk.BooleanVar(value=self.xy_show_mean_var.get())
        self.settings_xy_show_mean_errorbars_var = tk.BooleanVar(value=self.xy_show_mean_errorbars_var.get())
        self.settings_xy_draw_band_var = tk.BooleanVar(value=self.xy_draw_band_var.get())
        
        # Colors tab
        self.settings_single_color_var = tk.StringVar(value=self.single_color_var.get() if hasattr(self, 'single_color_var') else list(self.custom_colors.keys())[0])
        self.settings_palette_var = tk.StringVar(value=self.palette_var.get() if hasattr(self, 'palette_var') else list(self.custom_palettes.keys())[0])
        self.settings_outline_color_var = tk.StringVar(value=self.outline_color_var.get())
        
        # General Tab Content
        ttk.Label(general_tab, text="Default Plot Type:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(general_tab, textvariable=self.settings_plot_kind_var, values=["bar", "box", "violin", "xy"], width=15, state="readonly").grid(row=0, column=1, sticky="w", padx=10, pady=10)

        ttk.Checkbutton(general_tab, text="Start maximized", variable=self.settings_start_maximized_var).grid(row=1, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        # Plot Settings Tab Content
        ttk.Checkbutton(plot_settings_tab, text="Show stripplot", variable=self.settings_show_stripplot_var).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Checkbutton(plot_settings_tab, text="Use black stripplot", variable=self.settings_strip_black_var).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        ttk.Label(plot_settings_tab, text="Error bar type:", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(plot_settings_tab, textvariable=self.settings_errorbar_type_var, values=["SD", "SEM"], width=15, state="readonly").grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Checkbutton(plot_settings_tab, text="Black error bars", variable=self.settings_errorbar_black_var).grid(row=3, column=0, sticky="w", padx=10, pady=5)
        
        ttk.Label(plot_settings_tab, text="Error bar capsize:", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(plot_settings_tab, textvariable=self.settings_errorbar_capsize_var, values=["Default", "None", "Small", "Medium", "Large"], width=15, state="readonly").grid(row=4, column=1, sticky="w", padx=10, pady=10)
        
        # Statistics Tab Content
        # Use statistics checkbox at the top
        ttk.Checkbutton(stats_tab, text="Use statistics by default", variable=self.settings_use_stats_var).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        ttk.Label(stats_tab, text="t-test type:", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        ttest_options = ["Student's t-test (unpaired, equal variances)", "Welch's t-test (unpaired, unequal variances)", "Paired t-test", "Mann-Whitney U test (non-parametric)", "Wilcoxon signed-rank test (non-parametric)"]
        ttk.Combobox(stats_tab, textvariable=self.settings_ttest_type_var, values=ttest_options, width=35, state="readonly").grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(stats_tab, text="T-test Alternative:", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttest_alternative_options = ["two-sided", "less", "greater"]
        ttk.Combobox(stats_tab, textvariable=self.settings_ttest_alternative_var, values=ttest_alternative_options, width=35, state="readonly").grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(stats_tab, text="ANOVA type:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        anova_options = ["One-way ANOVA", "Welch's ANOVA", "Repeated measures ANOVA", "Kruskal-Wallis H test (non-parametric)", "Friedman test (non-parametric)"]
        ttk.Combobox(stats_tab, textvariable=self.settings_anova_type_var, values=anova_options, width=35, state="readonly").grid(row=3, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(stats_tab, text="Alpha level:", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        alpha_options = ["0.05", "0.01", "0.001", "0.0001"]
        ttk.Combobox(stats_tab, textvariable=self.settings_alpha_level_var, values=alpha_options, width=35, state="readonly").grid(row=4, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(stats_tab, text="Post-hoc test:", anchor="w").grid(row=5, column=0, sticky="w", padx=10, pady=10)
        posthoc_options = ["Tukey's HSD", "Tamhane's T2", "Scheffe's test", "Dunn's test", "Conover's test (non-parametric)", "Nemenyi test (non-parametric)"]
        ttk.Combobox(stats_tab, textvariable=self.settings_posthoc_type_var, values=posthoc_options, width=35, state="readonly").grid(row=5, column=1, sticky="w", padx=10, pady=10)
        
        # Appearance Tab Content
        ttk.Label(appearance_tab, text="Line width:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.5, textvariable=self.settings_linewidth, width=5).grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(appearance_tab, text="Plot width (inches):", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.1, textvariable=self.settings_plot_width_var, width=5).grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(appearance_tab, text="Plot height (inches):", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.1, textvariable=self.settings_plot_height_var, width=5).grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(appearance_tab, text="Preview DPI:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=50, to=300, increment=25, textvariable=self.settings_preview_dpi_var, width=5).grid(row=3, column=1, sticky="w", padx=10, pady=10)

        ttk.Label(appearance_tab, text="Y-break gap:", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.0, to=0.3, increment=0.01, textvariable=self.settings_ybreak_gap_var, width=6).grid(row=4, column=1, sticky="w", padx=10, pady=10)

        ttk.Label(appearance_tab, text="Y-break marker:", anchor="w").grid(row=5, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(appearance_tab, textvariable=self.settings_ybreak_marker_style_var, values=["Connected", "Diagonal"], width=12, state="readonly").grid(row=5, column=1, sticky="w", padx=10, pady=10)
        
        # Bar Graph Tab Content
        ttk.Checkbutton(bar_graph_tab, text="Draw bar outlines", variable=self.settings_bar_outline_var).grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Checkbutton(bar_graph_tab, text="Upward-only error bars", variable=self.settings_upward_errorbar_var).grid(row=1, column=0, sticky="w", padx=10, pady=10)
        
        ttk.Label(bar_graph_tab, text="Bar gap multiplier:", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(bar_graph_tab, from_=0.5, to=1.0, increment=0.05, textvariable=self.settings_bar_gap_multiplier_var, width=5).grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        # XY Plot Tab Content
        ttk.Label(xy_plot_tab, text="Marker Symbol:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(xy_plot_tab, textvariable=self.settings_xy_marker_symbol_var, values=["o", "s", "^", "D", "v", "P", "X", "+", "x", "*", "."], width=5, state="readonly").grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Label(xy_plot_tab, text="Marker Size:", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(xy_plot_tab, from_=1, to=15, increment=0.5, textvariable=self.settings_xy_marker_size_var, width=5).grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Checkbutton(xy_plot_tab, text="Filled Symbols", variable=self.settings_xy_filled_var).grid(row=2, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        
        ttk.Label(xy_plot_tab, text="Line Style:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(xy_plot_tab, textvariable=self.settings_xy_line_style_var, values=["solid", "dashed", "dotted", "dashdot"], width=10, state="readonly").grid(row=3, column=1, sticky="w", padx=10, pady=10)
        
        ttk.Checkbutton(xy_plot_tab, text="Lines in Black", variable=self.settings_xy_line_black_var).grid(row=4, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        ttk.Checkbutton(xy_plot_tab, text="Connect Mean with Lines", variable=self.settings_xy_connect_var).grid(row=5, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        ttk.Checkbutton(xy_plot_tab, text="Show Mean Values", variable=self.settings_xy_show_mean_var).grid(row=6, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        ttk.Checkbutton(xy_plot_tab, text="With Error Bars", variable=self.settings_xy_show_mean_errorbars_var).grid(row=7, column=0, sticky="w", padx=30, pady=5, columnspan=2)
        ttk.Checkbutton(xy_plot_tab, text="Draw Bands", variable=self.settings_xy_draw_band_var).grid(row=8, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        
        # Colors Tab Content
        # Reset colors/palettes section
        reset_frame = ttk.Frame(colors_tab)
        reset_frame.grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 20))
        
        ttk.Label(reset_frame, text="Reset options:", font=(None, 10, 'bold')).pack(side=tk.LEFT, padx=5)
        
        def reset_colors():
            if messagebox.askyesno("Reset Colors", "This will delete your custom colors and cannot be undone. Continue?"):
                if os.path.exists(self.custom_colors_file):
                    os.remove(self.custom_colors_file)
                self.load_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                messagebox.showinfo("Reset Colors", "Colors have been reset to default.")
                
        def reset_palettes():
            if messagebox.askyesno("Reset Palettes", "This will delete your custom palettes and cannot be undone. Continue?"):
                if os.path.exists(self.custom_palettes_file):
                    os.remove(self.custom_palettes_file)
                self.load_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                messagebox.showinfo("Reset Palettes", "Color palettes have been reset to default.")
        
        ttk.Button(reset_frame, text="Reset Colors", command=reset_colors, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(reset_frame, text="Reset Palettes", command=reset_palettes, width=15).pack(side=tk.LEFT, padx=5)
        
        # Single Color section
        ttk.Label(colors_tab, text="Single Data Color:", anchor="w", font=(None, 10, 'bold')).grid(row=1, column=0, sticky="w", padx=10, pady=(10, 5))
        self.settings_single_color_dropdown = ttk.Combobox(colors_tab, textvariable=self.settings_single_color_var, values=list(self.custom_colors.keys()), width=20, state="readonly")
        self.settings_single_color_dropdown.grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        # Create a frame for the color preview
        single_color_preview_frame = ttk.Frame(colors_tab)
        single_color_preview_frame.grid(row=2, column=0, sticky="w", padx=10, pady=5)
        self.settings_single_color_preview = tk.Canvas(single_color_preview_frame, width=60, height=20, highlightthickness=1, bg='white')
        self.settings_single_color_preview.pack()
        
        # Function to update single color preview
        def update_settings_single_color_preview(event=None):
            self.settings_single_color_preview.delete('all')
            name = self.settings_single_color_var.get()
            hexcode = self.custom_colors.get(name)
            if hexcode:
                self.settings_single_color_preview.create_rectangle(10, 2, 50, 18, fill=hexcode, outline='black')
            else:
                # If color not found, reset to first available color
                if self.custom_colors:
                    self.settings_single_color_var.set(list(self.custom_colors.keys())[0])
                    update_settings_single_color_preview()
        
        self.settings_single_color_dropdown.bind('<<ComboboxSelected>>', update_settings_single_color_preview)
        update_settings_single_color_preview()
        
        # Group Palette section
        ttk.Label(colors_tab, text="Group Palette:", anchor="w", font=(None, 10, 'bold')).grid(row=3, column=0, sticky="w", padx=10, pady=(20, 5))
        self.settings_palette_dropdown = ttk.Combobox(colors_tab, textvariable=self.settings_palette_var, values=list(self.custom_palettes.keys()), width=20, state="readonly")
        self.settings_palette_dropdown.grid(row=4, column=0, sticky="w", padx=10, pady=5)
        
        # Create a frame for the palette preview
        palette_preview_frame = ttk.Frame(colors_tab)
        palette_preview_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=5)
        self.settings_palette_preview = tk.Canvas(palette_preview_frame, height=20, highlightthickness=1, bg='white')
        self.settings_palette_preview.pack(fill='x')
        
        # Function to update palette preview
        def update_settings_palette_preview(event=None):
            self.settings_palette_preview.delete('all')
            name = self.settings_palette_var.get()
            colors = self.custom_palettes.get(name, [])
            if colors:
                for i, hexcode in enumerate(colors):
                    x0 = 5 + i*14
                    x1 = x0 + 12
                    self.settings_palette_preview.create_rectangle(x0, 2, x1, 18, fill=hexcode, outline='black')
            else:
                # If palette not found, reset to first available palette
                if self.custom_palettes:
                    self.settings_palette_var.set(list(self.custom_palettes.keys())[0])
                    update_settings_palette_preview()
        
        self.settings_palette_dropdown.bind('<<ComboboxSelected>>', update_settings_palette_preview)
        update_settings_palette_preview()
        
        # Outline Color section
        ttk.Label(colors_tab, text="Outline Color:", anchor="w", font=(None, 10, 'bold')).grid(row=6, column=0, sticky="w", padx=10, pady=(20, 5))
        outline_frame = ttk.Frame(colors_tab)
        outline_frame.grid(row=7, column=0, sticky="w", padx=10, pady=5)
        
        ttk.Radiobutton(outline_frame, text="As set", variable=self.settings_outline_color_var, value="as_set").pack(anchor="w", pady=2)
        ttk.Radiobutton(outline_frame, text="Black", variable=self.settings_outline_color_var, value="black").pack(anchor="w", pady=2)
        ttk.Radiobutton(outline_frame, text="Gray", variable=self.settings_outline_color_var, value="gray").pack(anchor="w", pady=2)
        ttk.Radiobutton(outline_frame, text="White", variable=self.settings_outline_color_var, value="white").pack(anchor="w", pady=2)
        
        # Buttons at the bottom
        button_frame = ttk.Frame(window)
        button_frame.pack(pady=10, fill='x')
        
        def save_settings():
            # Update main variables from settings first
            self.bar_outline_var.set(self.settings_bar_outline_var.get())
            self.bar_gap_multiplier_var.set(self.settings_bar_gap_multiplier_var.get())
            # Update color settings
            self.single_color_var.set(self.settings_single_color_var.get())
            self.palette_var.set(self.settings_palette_var.get())
            self.outline_color_var.set(self.settings_outline_color_var.get())
            # Update preview DPI
            if hasattr(self, 'preview_dpi'):
                self.preview_dpi.set(self.settings_preview_dpi_var.get())

            if hasattr(self, 'start_maximized_var') and hasattr(self, 'settings_start_maximized_var'):
                self.start_maximized_var.set(self.settings_start_maximized_var.get())

            try:
                if hasattr(self, 'ybreak_marker_style_var') and hasattr(self, 'settings_ybreak_marker_style_var'):
                    self.ybreak_marker_style_var.set(self.settings_ybreak_marker_style_var.get())
                    if hasattr(self, 'ybreak_marker_style_user_set_var'):
                        self.ybreak_marker_style_user_set_var.set(True)
            except Exception:
                pass

            try:
                if hasattr(self, 'ybreak_gap_entry') and hasattr(self, 'settings_ybreak_gap_var'):
                    gap = float(self.settings_ybreak_gap_var.get())
                    gap = max(0.0, min(0.3, gap))
                    self.ybreak_gap_entry.delete(0, tk.END)
                    self.ybreak_gap_entry.insert(0, f"{gap:g}")
            except Exception:
                pass
            # Then save preferences
            self.save_user_preferences()
            messagebox.showinfo("Settings Saved", "Your preferences have been saved.")
            window.destroy()
            
        def reset_all_preferences():
            if messagebox.askyesno("Reset All Preferences", "This will reset all preferences to default values and cannot be undone. Continue?"):
                if os.path.exists(self.default_settings_file):
                    os.remove(self.default_settings_file)
                messagebox.showinfo("Reset Preferences", "All preferences have been reset to default values.")
                window.destroy()
        
        ttk.Button(button_frame, text="Save Settings", command=save_settings, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Reset All Preferences", command=reset_all_preferences, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=window.destroy, width=10).pack(side=tk.RIGHT, padx=5)

    def show_about(self):
        messagebox.showinfo("About ExPlot", f"ExPlot\nVersion: {self.version}\n\nA tool for plotting Excel data.")
        
    def show_package_info(self):
        """Display information about packages used in the application for scientific publications."""
        window = tk.Toplevel(self.root)
        window.title("Package Information for Publications")
        window.geometry("800x600")
        
        # Create notebook with tabs for different package categories
        notebook = ttk.Notebook(window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create frames for each category
        plotting_frame = ttk.Frame(notebook)
        stats_frame = ttk.Frame(notebook)
        fitting_frame = ttk.Frame(notebook)
        
        notebook.add(plotting_frame, text="Plotting")
        notebook.add(stats_frame, text="Statistics")
        notebook.add(fitting_frame, text="Fitting")
        
        # Get package versions
        try:
            import pandas as pd
            pandas_version = pd.__version__
        except:
            pandas_version = "Not available"
            
        try:
            import matplotlib
            matplotlib_version = matplotlib.__version__
        except:
            matplotlib_version = "Not available"
            
        try:
            import seaborn as sns
            seaborn_version = sns.__version__
        except:
            seaborn_version = "Not available"
            
        try:
            import numpy as np
            numpy_version = np.__version__
        except:
            numpy_version = "Not available"
            
        try:
            import scipy
            scipy_version = scipy.__version__
        except:
            scipy_version = "Not available"
            
        try:
            import pingouin as pg
            pingouin_version = pg.__version__
        except:
            pingouin_version = "Not available"
            
        try:
            import scikit_posthocs as sp
            scikit_posthocs_version = sp.__version__
        except:
            scikit_posthocs_version = "Not available"
        
        # Plotting packages information
        plotting_text = tk.Text(plotting_frame, wrap=tk.WORD, padx=10, pady=10)
        plotting_text.pack(fill='both', expand=True)
        plotting_text.insert(tk.END, "Packages for Data Visualization and Plotting:\n\n")
        plotting_text.insert(tk.END, f"1. Matplotlib (version {matplotlib_version})\n")
        plotting_text.insert(tk.END, "   Essential functions: pyplot, Figure, Axes, FigureCanvasTkAgg\n")
        plotting_text.insert(tk.END, "   Purpose: Core plotting library providing comprehensive visualization capabilities\n\n")
        
        plotting_text.insert(tk.END, f"2. Seaborn (version {seaborn_version})\n")
        plotting_text.insert(tk.END, "   Essential functions: color_palette\n")
        plotting_text.insert(tk.END, "   Purpose: High-level interface for creating statistical graphics with enhanced aesthetics\n\n")
        
        plotting_text.insert(tk.END, f"3. Pandas (version {pandas_version})\n")
        plotting_text.insert(tk.END, "   Essential functions: DataFrame, read_excel\n")
        plotting_text.insert(tk.END, "   Purpose: Data manipulation and analysis, loading Excel data\n\n")
        
        plotting_text.configure(state='disabled')  # Make read-only
        
        # Statistics packages information
        stats_text = tk.Text(stats_frame, wrap=tk.WORD, padx=10, pady=10)
        stats_text.pack(fill='both', expand=True)
        stats_text.insert(tk.END, "Packages for Statistical Analysis:\n\n")
        
        stats_text.insert(tk.END, f"1. SciPy (version {scipy_version})\n")
        stats_text.insert(tk.END, "   Essential functions: stats.ttest_ind, stats.ttest_rel, stats.f_oneway\n")
        stats_text.insert(tk.END, "   Purpose: Provides statistical functions for t-tests, ANOVA, and other statistical tests\n\n")
        
        stats_text.insert(tk.END, f"2. Pingouin (version {pingouin_version})\n")
        stats_text.insert(tk.END, "   Essential functions: welch_anova, pairwise_tests, rm_anova\n")
        stats_text.insert(tk.END, "   Purpose: Advanced statistical analyses including Welch's ANOVA and repeated measures ANOVA\n\n")
        
        stats_text.insert(tk.END, f"3. Scikit-posthocs (version {scikit_posthocs_version})\n")
        stats_text.insert(tk.END, "   Essential functions: posthoc_tukey, posthoc_tamhane, posthoc_games_howell, posthoc_dunn\n")
        stats_text.insert(tk.END, "   Purpose: Post-hoc tests following ANOVA (Tukey's HSD, Tamhane's T2, Games-Howell, Dunn's test)\n\n")
        
        stats_text.insert(tk.END, f"4. NumPy (version {numpy_version})\n")
        stats_text.insert(tk.END, "   Essential functions: array, mean, std, nan_to_num\n")
        stats_text.insert(tk.END, "   Purpose: Numerical operations and array handling for statistical computations\n\n")
        
        stats_text.configure(state='disabled')  # Make read-only
        
        # Fitting packages information
        fitting_text = tk.Text(fitting_frame, wrap=tk.WORD, padx=10, pady=10)
        fitting_text.pack(fill='both', expand=True)
        fitting_text.insert(tk.END, "Packages for Curve Fitting and Modeling:\n\n")
        
        fitting_text.insert(tk.END, f"1. SciPy (version {scipy_version})\n")
        fitting_text.insert(tk.END, "   Essential functions: optimize.curve_fit\n")
        fitting_text.insert(tk.END, "   Purpose: Non-linear least squares fitting for custom model functions\n\n")
        
        fitting_text.insert(tk.END, f"2. NumPy (version {numpy_version})\n")
        fitting_text.insert(tk.END, "   Essential functions: polyfit, poly1d\n")
        fitting_text.insert(tk.END, "   Purpose: Polynomial fitting and evaluation\n\n")
        
        fitting_text.configure(state='disabled')  # Make read-only
        
        # Add note about scientific publications in the fitting frame
        fitting_text.configure(state='normal')  # Temporarily enable for adding more text
        fitting_text.insert(tk.END, "\n\nNote for Scientific Publications:\n")
        fitting_text.insert(tk.END, "When citing this software in scientific publications, please include the essential packages ")
        fitting_text.insert(tk.END, "listed above with their version numbers and mention any specific statistical tests used ")
        fitting_text.insert(tk.END, "(e.g., Welch's ANOVA with Games-Howell post-hoc tests).")
        fitting_text.configure(state='disabled')  # Make read-only again
        
        # Add a button to copy information to clipboard
        def copy_to_clipboard():
            package_info = f"\"ExPlot\" (version {self.version}) was used for data visualization and statistical analysis. "
            package_info += f"It utilizes the following Python packages: matplotlib {matplotlib_version}, seaborn {seaborn_version}, "
            package_info += f"pandas {pandas_version}, numpy {numpy_version}, scipy {scipy_version}, pingouin {pingouin_version}, "
            package_info += f"and scikit-posthocs {scikit_posthocs_version}.\n"
            
            # Copy to clipboard
            window.clipboard_clear()
            window.clipboard_append(package_info)
            messagebox.showinfo("Copied", "Package information copied to clipboard!")
            
        button_frame = ttk.Frame(window)
        button_frame.pack(fill='x', padx=10, pady=10)
        
        copy_button = ttk.Button(button_frame, text="Copy Citation Text", command=copy_to_clipboard)
        copy_button.pack(side='right')
        
    def open_label_formatter(self, axis_type):
        """Open a dialog for formatting axis labels with rich text capabilities."""
        # Get current label text
        entry_widget = self.xlabel_entry if axis_type == 'x' else self.ylabel_entry
        current_text = entry_widget.get()
        
        # Create dialog window
        window = tk.Toplevel(self.root)
        window.title(f"Format {'X' if axis_type == 'x' else 'Y'}-Axis Label")
        window.geometry("650x650")
        window.transient(self.root)
        window.grab_set()
        
        # Create main frame with padding
        main_frame = ttk.Frame(window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Create preview frame
        preview_frame = ttk.LabelFrame(main_frame, text="Preview")
        preview_frame.pack(fill='x', padx=5, pady=5)
        
        # Create canvas for preview
        preview_fig = plt.Figure(figsize=(3, 3), dpi=75)
        preview_ax = preview_fig.add_subplot(111)
        preview_canvas = FigureCanvasTkAgg(preview_fig, master=preview_frame)
        preview_canvas.get_tk_widget().pack(fill='both', expand=True)
        preview_ax.set_xticks([])
        preview_ax.set_yticks([])
        
        # Function to update preview
        def update_preview(text=None):
            if text is None:
                text = text_entry.get("1.0", tk.END).strip()
            preview_ax.clear()
            preview_ax.set_xticks([])
            preview_ax.set_yticks([])
            if axis_type == 'x':
                preview_ax.set_xlabel(text, fontsize=12)
            else:
                preview_ax.set_ylabel(text, fontsize=12)
            preview_fig.tight_layout()
            preview_canvas.draw()
        
        # Create editing frame
        edit_frame = ttk.LabelFrame(main_frame, text="Edit Label Text")
        edit_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Text entry for the label
        text_entry = tk.Text(edit_frame, height=3, width=50, wrap=tk.WORD)
        text_entry.pack(fill='both', expand=True, padx=5, pady=5)
        text_entry.insert("1.0", current_text)
        
        # Bind text change to preview update
        text_entry.bind("<KeyRelease>", lambda e: update_preview())
        
        # Create formatting buttons frame
        format_frame = ttk.LabelFrame(main_frame, text="Formatting Options")
        format_frame.pack(fill='x', padx=5, pady=5)
        
        # Create two rows for buttons to ensure they all fit
        buttons_row1 = ttk.Frame(format_frame)
        buttons_row1.pack(fill='x', padx=5, pady=5)
        
        buttons_row2 = ttk.Frame(format_frame)
        buttons_row2.pack(fill='x', padx=5, pady=5)
        
        # First row of formatting buttons
        ttk.Button(buttons_row1, text="Bold", width=12,
                  command=lambda: insert_format_tag(r"\mathbf{", "}")).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(buttons_row1, text="Italic", width=12,
                  command=lambda: insert_format_tag(r"\it{", "}")).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(buttons_row1, text="Superscript", width=12,
                  command=lambda: insert_format_tag("^{", "}")).pack(side=tk.LEFT, padx=5, pady=3)
        
        # Second row of formatting buttons
        ttk.Button(buttons_row2, text="Subscript", width=12,
                  command=lambda: insert_format_tag("_{", "}")).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(buttons_row2, text="Greek μ", width=12,
                  command=lambda: insert_text(r"\mu")).pack(side=tk.LEFT, padx=5, pady=3)
        ttk.Button(buttons_row2, text="°C", width=12,
                  command=lambda: insert_text(r"^{\circ}C")).pack(side=tk.LEFT, padx=5, pady=3)
        
        # Mathematical mode frame
        math_frame = ttk.LabelFrame(main_frame, text="Mathematical Mode")
        math_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(math_frame, text="For formatting to work, you must wrap parts that use formatting in $ signs.").pack(pady=2)
        example_text = "$\\mathbf{Bold}$ normal $\\it{italic}$ $\\mu$mol$^{-1}$"
        ttk.Label(math_frame, text=f"Example: {example_text}").pack(pady=2)
        
        # Insert formatting function
        def insert_format_tag(prefix, suffix):
            try:
                # Get selection or insert position
                if text_entry.tag_ranges(tk.SEL):
                    start = text_entry.index(tk.SEL_FIRST)
                    end = text_entry.index(tk.SEL_LAST)
                    selected_text = text_entry.get(start, end)
                    
                    # Check if we need to add dollar signs
                    text_before = text_entry.get("1.0", start)
                    text_after = text_entry.get(end, tk.END)
                    
                    needs_dollars = True
                    if "$" in text_before and text_before.count("$") % 2 == 1:
                        # We're already in math mode
                        needs_dollars = False
                    
                    # Apply formatting
                    formatted_text = selected_text
                    if needs_dollars:
                        formatted_text = f"${prefix}{selected_text}{suffix}$"
                    else:
                        formatted_text = f"{prefix}{selected_text}{suffix}"
                        
                    text_entry.delete(start, end)
                    text_entry.insert(start, formatted_text)
                else:
                    # No selection, insert at cursor
                    cursor_pos = text_entry.index(tk.INSERT)
                    text_entry.insert(cursor_pos, f"${prefix}text{suffix}$")
                    
                    # Select the word "text" for easy replacement
                    start_pos = f"{cursor_pos}+{1 + len(prefix)}c"
                    end_pos = f"{start_pos}+4c"  # "text" is 4 characters
                    text_entry.mark_set(tk.INSERT, start_pos)
                    text_entry.tag_add(tk.SEL, start_pos, end_pos)
                    text_entry.focus_set()
                
                update_preview()
            except Exception as e:
                print(f"Error in insert_format_tag: {e}")
        
        # Insert text function
        def insert_text(text):
            cursor_pos = text_entry.index(tk.INSERT)
            
            # Check if we're in math mode
            text_before = text_entry.get("1.0", cursor_pos)
            text_after = text_entry.get(cursor_pos, tk.END)
            
            in_math_mode = text_before.count("$") % 2 == 1
            
            if in_math_mode:
                text_entry.insert(cursor_pos, text)
            else:
                text_entry.insert(cursor_pos, f"${text}$")
            
            update_preview()
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', padx=5, pady=10)
        
        # Apply and cancel buttons
        def apply_format():
            new_text = text_entry.get("1.0", tk.END).strip()
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, new_text)
            window.destroy()
        
        ttk.Button(button_frame, text="Apply", command=apply_format).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=window.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Initial preview
        update_preview(current_text)
        
        # Set focus to text entry
        text_entry.focus_set()
        
    def show_stats_info(self):
        """Display information about statistical tests and when to use them."""
        window = tk.Toplevel(self.root)
        window.title("Statistical Tests Information")
        window.geometry("800x600")
        
        # Create notebook with tabs for different test categories
        notebook = ttk.Notebook(window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        t_test_frame = ttk.Frame(notebook)
        anova_frame = ttk.Frame(notebook)
        posthoc_frame = ttk.Frame(notebook)
        general_frame = ttk.Frame(notebook)
        
        notebook.add(general_frame, text='General Guidelines')
        notebook.add(t_test_frame, text='t-tests')
        notebook.add(anova_frame, text='ANOVA')
        notebook.add(posthoc_frame, text='Post-hoc Tests')
        
        # Function to create formatted text widgets
        def create_text_widget(parent):
            text = tk.Text(parent, wrap='word', padx=10, pady=10)
            text.pack(fill='both', expand=True)
            scrollbar = ttk.Scrollbar(text, command=text.yview)
            text.configure(yscrollcommand=scrollbar.set)
            scrollbar.pack(side='right', fill='y')
            text.tag_configure('heading', font=(None, 12, 'bold'))
            text.tag_configure('subheading', font=(None, 11, 'bold'))
            text.tag_configure('normal', font=(None, 10))
            return text
        
        # General guidelines
        general_text = create_text_widget(general_frame)
        general_text.insert('end', 'When to Use Which Statistical Test\n', 'heading')
        general_text.insert('end', '\nThis guide helps you choose the appropriate statistical test for your data.\n\n', 'normal')
        
        general_text.insert('end', 'Quick Decision Guide:\n', 'subheading')
        general_text.insert('end', '\n', 'normal')
        general_text.insert('end', '1. Comparing TWO groups:\n', 'subheading')
        general_text.insert('end', '   • Data normally distributed → t-test (Welch\'s recommended)\n', 'normal')
        general_text.insert('end', '   • Data NOT normally distributed → Mann-Whitney U test\n', 'normal')
        general_text.insert('end', '   • Paired/matched samples → Paired t-test or Wilcoxon signed-rank\n\n', 'normal')
        
        general_text.insert('end', '2. Comparing THREE OR MORE groups:\n', 'subheading')
        general_text.insert('end', '   • Data normally distributed → ANOVA + post-hoc test\n', 'normal')
        general_text.insert('end', '   • Data NOT normally distributed → Kruskal-Wallis + Dunn\'s test\n\n', 'normal')
        
        general_text.insert('end', 'Parametric vs. Non-parametric Tests:\n', 'subheading')
        general_text.insert('end', '\n• Parametric tests (t-test, ANOVA) assume your data follows a normal distribution.\n', 'normal')
        general_text.insert('end', '  They are more powerful when assumptions are met.\n\n', 'normal')
        general_text.insert('end', '• Non-parametric tests (Mann-Whitney, Wilcoxon, Dunn\'s) make no assumptions about\n', 'normal')
        general_text.insert('end', '  the distribution. Use them when:\n', 'normal')
        general_text.insert('end', '  - Data is skewed or has outliers\n', 'normal')
        general_text.insert('end', '  - Sample sizes are very small (n < 10)\n', 'normal')
        general_text.insert('end', '  - Data is ordinal (ranked) rather than continuous\n\n', 'normal')
        
        general_text.insert('end', 'Important Notes on Small Sample Sizes:\n', 'subheading')
        general_text.insert('end', '\n• Non-parametric tests with small samples have LIMITED resolution.\n', 'normal')
        general_text.insert('end', '• Example: Mann-Whitney U with n=4 vs n=4 has a minimum p-value of 0.029.\n', 'normal')
        general_text.insert('end', '  This means you cannot achieve p < 0.029 regardless of how different the groups are.\n', 'normal')
        general_text.insert('end', '• If all your comparisons show identical p-values, this is likely because all\n', 'normal')
        general_text.insert('end', '  groups show complete separation (the minimum possible p-value).\n', 'normal')
        general_text.insert('end', '• Consider increasing sample size or using parametric tests if appropriate.\n', 'normal')
        general_text.configure(state='disabled')  # Make read-only
        
        # t-tests information
        t_test_text = create_text_widget(t_test_frame)
        t_test_text.insert('end', 'Two-Group Comparison Tests\n', 'heading')
        t_test_text.insert('end', '\nThese tests compare measurements between exactly two groups.\n\n', 'normal')
        
        t_test_text.insert('end', '═══ PARAMETRIC TESTS (assume normal distribution) ═══\n\n', 'subheading')
        
        t_test_text.insert('end', 'Student\'s t-test (unpaired, equal variances):\n', 'subheading')
        t_test_text.insert('end', '• Best for: Two independent groups with similar spread (variance)\n', 'normal')
        t_test_text.insert('end', '• Example: Comparing cell counts between two different cell lines\n', 'normal')
        t_test_text.insert('end', '• Assumptions: Normal distribution, equal variances, independent samples\n', 'normal')
        t_test_text.insert('end', '• Note: Rarely used in practice—Welch\'s is usually preferred\n\n', 'normal')
        
        t_test_text.insert('end', 'Welch\'s t-test (unpaired, unequal variances):\n', 'subheading')
        t_test_text.insert('end', '• Best for: Two independent groups (default choice for most comparisons)\n', 'normal')
        t_test_text.insert('end', '• Example: Control vs. Treatment groups with different variability\n', 'normal')
        t_test_text.insert('end', '• Assumptions: Normal distribution, independent samples\n', 'normal')
        t_test_text.insert('end', '• Advantage: Does NOT require equal variances—more robust\n', 'normal')
        t_test_text.insert('end', '• ✓ RECOMMENDED as default for unpaired comparisons\n\n', 'normal')
        
        t_test_text.insert('end', 'Paired t-test:\n', 'subheading')
        t_test_text.insert('end', '• Best for: Matched or repeated measurements on the same subjects\n', 'normal')
        t_test_text.insert('end', '• Example: Before vs. After treatment in the same patients\n', 'normal')
        t_test_text.insert('end', '• Example: Left eye vs. Right eye measurements\n', 'normal')
        t_test_text.insert('end', '• Requirement: Equal number of data points in both groups\n', 'normal')
        t_test_text.insert('end', '• Advantage: More powerful than unpaired tests for paired data\n\n', 'normal')
        
        t_test_text.insert('end', '═══ NON-PARAMETRIC TESTS (no distribution assumptions) ═══\n\n', 'subheading')
        
        t_test_text.insert('end', 'Mann-Whitney U test (unpaired):\n', 'subheading')
        t_test_text.insert('end', '• Best for: Two independent groups when data is NOT normally distributed\n', 'normal')
        t_test_text.insert('end', '• Also called: Wilcoxon rank-sum test (different from signed-rank!)\n', 'normal')
        t_test_text.insert('end', '• Example: Comparing scores when data is skewed or has outliers\n', 'normal')
        t_test_text.insert('end', '• How it works: Compares RANKS of values, not the values themselves\n', 'normal')
        t_test_text.insert('end', '• ⚠ IMPORTANT: With small samples, p-values have limited resolution!\n', 'normal')
        t_test_text.insert('end', '    - n=4 vs n=4: minimum p-value = 0.029 (cannot go lower)\n', 'normal')
        t_test_text.insert('end', '    - n=5 vs n=5: minimum p-value = 0.008\n', 'normal')
        t_test_text.insert('end', '    - n=10 vs n=10: minimum p-value = 0.00006\n\n', 'normal')
        
        t_test_text.insert('end', 'Wilcoxon signed-rank test (paired):\n', 'subheading')
        t_test_text.insert('end', '• Best for: Paired/matched data when differences are NOT normally distributed\n', 'normal')
        t_test_text.insert('end', '• Example: Before vs. After when the differences are skewed\n', 'normal')
        t_test_text.insert('end', '• How it works: Ranks the absolute differences, then compares signed ranks\n', 'normal')
        t_test_text.insert('end', '• Requirement: Equal number of data points in both groups\n', 'normal')
        t_test_text.insert('end', '• ⚠ Same small-sample limitations as Mann-Whitney U\n\n', 'normal')
        
        t_test_text.insert('end', '═══ ALTERNATIVE HYPOTHESIS OPTIONS ═══\n\n', 'subheading')
        
        t_test_text.insert('end', '• two-sided: Are the groups different? (most common, use when unsure)\n', 'normal')
        t_test_text.insert('end', '• greater: Is group 1 larger than group 2?\n', 'normal')
        t_test_text.insert('end', '• less: Is group 1 smaller than group 2?\n', 'normal')
        t_test_text.insert('end', '\nNote: One-sided tests are more powerful but require prior hypothesis.\n', 'normal')
        t_test_text.configure(state='disabled')  # Make read-only
        
        # ANOVA information
        anova_text = create_text_widget(anova_frame)
        anova_text.insert('end', 'Multi-Group Comparison Tests (ANOVA)\n', 'heading')
        anova_text.insert('end', '\nANOVA (Analysis of Variance) tests whether means differ across 3+ groups.\n', 'normal')
        anova_text.insert('end', 'It answers: "Is there ANY difference among these groups?"\n\n', 'normal')
        
        anova_text.insert('end', '═══ PARAMETRIC ANOVA TYPES ═══\n\n', 'subheading')
        
        anova_text.insert('end', 'One-way ANOVA:\n', 'subheading')
        anova_text.insert('end', '• Best for: Comparing 3+ independent groups with similar variances\n', 'normal')
        anova_text.insert('end', '• Example: Comparing expression levels across 4 different cell types\n', 'normal')
        anova_text.insert('end', '• Assumptions: Normal distribution, equal variances, independent samples\n', 'normal')
        anova_text.insert('end', '• Output: F-statistic and p-value for overall group difference\n\n', 'normal')
        
        anova_text.insert('end', 'Welch\'s ANOVA:\n', 'subheading')
        anova_text.insert('end', '• Best for: Comparing 3+ independent groups with UNEQUAL variances\n', 'normal')
        anova_text.insert('end', '• Example: Treatment groups where one group is more variable\n', 'normal')
        anova_text.insert('end', '• Assumptions: Normal distribution, independent samples\n', 'normal')
        anova_text.insert('end', '• Advantage: Does NOT require equal variances—more robust\n', 'normal')
        anova_text.insert('end', '• ✓ RECOMMENDED as default for most multi-group comparisons\n\n', 'normal')
        
        anova_text.insert('end', 'Repeated Measures ANOVA:\n', 'subheading')
        anova_text.insert('end', '• Best for: Multiple measurements on the SAME subjects\n', 'normal')
        anova_text.insert('end', '• Example: Drug response measured at 0h, 6h, 12h, 24h in same patients\n', 'normal')
        anova_text.insert('end', '• Example: Same mice tested under 3 different conditions\n', 'normal')
        anova_text.insert('end', '• Advantage: Much more powerful because it accounts for within-subject variation\n', 'normal')
        anova_text.insert('end', '• Requirement: Balanced design (same subjects in all conditions)\n\n', 'normal')
        
        anova_text.insert('end', '═══ NON-PARAMETRIC ALTERNATIVE ═══\n\n', 'subheading')
        
        anova_text.insert('end', 'Kruskal-Wallis test:\n', 'subheading')
        anova_text.insert('end', '• Best for: 3+ groups when data is NOT normally distributed\n', 'normal')
        anova_text.insert('end', '• How it works: Compares ranks instead of actual values\n', 'normal')
        anova_text.insert('end', '• Follow with Dunn\'s test for pairwise comparisons\n\n', 'normal')
        
        anova_text.insert('end', '═══ IMPORTANT NOTES ═══\n\n', 'subheading')
        
        anova_text.insert('end', '⚠ ANOVA only tells you IF differences exist, not WHERE.\n', 'normal')
        anova_text.insert('end', '  A significant ANOVA means at least one group differs from another,\n', 'normal')
        anova_text.insert('end', '  but it does NOT tell you which specific groups are different.\n\n', 'normal')
        anova_text.insert('end', '→ Always follow significant ANOVA with POST-HOC TESTS to identify\n', 'normal')
        anova_text.insert('end', '  which specific pairs of groups differ (see Post-hoc Tests tab).\n', 'normal')
        anova_text.configure(state='disabled')  # Make read-only
        
        # Post-hoc tests information
        posthoc_text = create_text_widget(posthoc_frame)
        posthoc_text.insert('end', 'Post-hoc Tests (Pairwise Comparisons)\n', 'heading')
        posthoc_text.insert('end', '\nAfter ANOVA shows significant differences, post-hoc tests identify WHICH\n', 'normal')
        posthoc_text.insert('end', 'specific pairs of groups differ from each other.\n\n', 'normal')
        
        posthoc_text.insert('end', '═══ WHY POST-HOC TESTS? ═══\n\n', 'subheading')
        posthoc_text.insert('end', 'When comparing multiple groups, running many t-tests inflates false positive\n', 'normal')
        posthoc_text.insert('end', 'rate (Type I error). Post-hoc tests correct for multiple comparisons.\n\n', 'normal')
        posthoc_text.insert('end', 'Example: With 5 groups, there are 10 possible pairwise comparisons.\n', 'normal')
        posthoc_text.insert('end', 'At α=0.05, you\'d expect ~0.5 false positives by chance alone!\n\n', 'normal')
        
        posthoc_text.insert('end', '═══ PARAMETRIC POST-HOC TESTS ═══\n\n', 'subheading')
        
        posthoc_text.insert('end', 'Tukey\'s HSD (Honestly Significant Difference):\n', 'subheading')
        posthoc_text.insert('end', '• Best for: Balanced designs with equal sample sizes and variances\n', 'normal')
        posthoc_text.insert('end', '• Tests ALL possible pairwise comparisons simultaneously\n', 'normal')
        posthoc_text.insert('end', '• Controls family-wise error rate (probability of ANY false positive)\n', 'normal')
        posthoc_text.insert('end', '• Good balance between power and Type I error control\n', 'normal')
        posthoc_text.insert('end', '• Use after: Standard one-way ANOVA\n\n', 'normal')
        
        posthoc_text.insert('end', 'Games-Howell:\n', 'subheading')
        posthoc_text.insert('end', '• Best for: Unequal sample sizes AND unequal variances\n', 'normal')
        posthoc_text.insert('end', '• Does not assume equal variances or sample sizes\n', 'normal')
        posthoc_text.insert('end', '• More powerful than Tamhane\'s T2 in most situations\n', 'normal')
        posthoc_text.insert('end', '• ✓ RECOMMENDED after Welch\'s ANOVA\n\n', 'normal')
        
        posthoc_text.insert('end', 'Tamhane\'s T2:\n', 'subheading')
        posthoc_text.insert('end', '• Best for: Unequal variances across groups\n', 'normal')
        posthoc_text.insert('end', '• Very conservative—less likely to find significance\n', 'normal')
        posthoc_text.insert('end', '• Good when you want to be extra careful about false positives\n', 'normal')
        posthoc_text.insert('end', '• Use after: Welch\'s ANOVA when being conservative\n\n', 'normal')
        
        posthoc_text.insert('end', 'Scheffé\'s test:\n', 'subheading')
        posthoc_text.insert('end', '• Best for: Complex contrasts beyond simple pairwise comparisons\n', 'normal')
        posthoc_text.insert('end', '• Most conservative test—least likely to find significance\n', 'normal')
        posthoc_text.insert('end', '• Can test ANY combination of group comparisons\n', 'normal')
        posthoc_text.insert('end', '• Use when: You want to compare group combinations (e.g., A vs B+C)\n\n', 'normal')
        
        posthoc_text.insert('end', '═══ NON-PARAMETRIC POST-HOC TEST ═══\n\n', 'subheading')
        
        posthoc_text.insert('end', 'Dunn\'s test:\n', 'subheading')
        posthoc_text.insert('end', '• Best for: Non-normally distributed data\n', 'normal')
        posthoc_text.insert('end', '• Works on ranks, not raw values\n', 'normal')
        posthoc_text.insert('end', '• Use after: Kruskal-Wallis test (non-parametric ANOVA)\n', 'normal')
        posthoc_text.insert('end', '• Includes Bonferroni correction for multiple comparisons\n\n', 'normal')
        
        posthoc_text.insert('end', '═══ QUICK SELECTION GUIDE ═══\n\n', 'subheading')
        posthoc_text.insert('end', '• Equal variances, equal n → Tukey\'s HSD\n', 'normal')
        posthoc_text.insert('end', '• Unequal variances or unequal n → Games-Howell\n', 'normal')
        posthoc_text.insert('end', '• Want to be very conservative → Tamhane\'s T2 or Scheffé\n', 'normal')
        posthoc_text.insert('end', '• Non-normal data → Dunn\'s test\n', 'normal')
        posthoc_text.configure(state='disabled')  # Make read-only
        
        # Close button at bottom
        ttk.Button(window, text="Close", command=window.destroy).pack(pady=10)

    def show_statistical_details(self):
        """Show a window with detailed statistical results.
        Now checks if statistics have been generated before displaying results.
        """
        # Initialize posthoc_matrices to prevent NameError
        posthoc_matrices = {}
        
        # Check if statistics have been generated yet
        if not hasattr(self, 'latest_pvals') or not self.latest_pvals:
            messagebox.showinfo("No Statistics", "No statistics have been generated yet. Please use the 'Generate Statistics' button first.")
            return
            
        # Create the statistical details window with a wider default size
        window = tk.Toplevel(self.root)
        window.title("Statistical Details")
        window.geometry("700x700")  # Increased width from 500 to 700
        
        # Create a frame to hold the text widget and scrollbar
        frame = ttk.Frame(window)
        frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Add a scrollbar
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side='right', fill='y')
        
        # Create the text widget with the scrollbar
        details_text = tk.Text(frame, wrap='word', height=30, width=100, yscrollcommand=scrollbar.set)
        details_text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=details_text.yview)
        # Set a monospaced font for table alignment
        import tkinter.font as tkfont
        monospace_font = None
        for fname in ("TkFixedFont", "Courier", "Menlo", "Consolas", "Monaco", "Liberation Mono"):
            if fname in tkfont.families():
                monospace_font = fname
                break
        if monospace_font is None:
            monospace_font = "Courier"  # fallback
        try:
            details_text.config(font=(monospace_font, 10))  # Set font to monospace for better table alignment
        except Exception as e:
            print(f"[DEBUG] Failed to set monospace font: {e}")
        details_text.insert(tk.END, "Statistical Details\n\n")
        # Legend
        # Get the current alpha level from dropdown to show correct significance thresholds
        try:
            alpha = float(self.alpha_level_var.get())
        except (ValueError, AttributeError):
            alpha = 0.05  # Default if not set or invalid
        
        # Get the current alpha level for dynamic thresholds
        try:
            alpha = float(self.alpha_level_var.get())
        except (ValueError, AttributeError):
            alpha = 0.05  # Default if not set or invalid
            
        # Define the same thresholds as in the statannotations configuration
        thresholds = [
            (alpha/5000, '****'),  # 4 stars threshold (e.g., 0.00001 at alpha=0.05)
            (alpha/50, '***'),    # 3 stars threshold (e.g., 0.001 at alpha=0.05)
            (alpha/5, '**'),      # 2 stars threshold (e.g., 0.01 at alpha=0.05)
            (alpha, '*'),         # 1 star threshold (alpha itself)
            (1, 'ns')             # Not significant
        ]
        
        # Display the significance levels with the actual threshold values
        details_text.insert(tk.END, "Significance levels:\n")
        prev_threshold = 1.0
        for threshold, symbol in sorted(thresholds, reverse=True):
            if symbol == 'ns':
                details_text.insert(tk.END, f"     {symbol} p > {alpha:.5g}\n")
            else:
                details_text.insert(tk.END, f"{symbol:>5} p ≤ {threshold:.5g}\n")
                prev_threshold = threshold
        details_text.insert(tk.END, "\n")
        
        # Show which alpha level is currently being used
        details_text.insert(tk.END, f"Current alpha level: {alpha}\n\n")
        # Try to reconstruct the most recent statistical tests
        # We'll use the existing p-values that were calculated for annotations
        # DO NOT clear the p-values dictionary here
        if not hasattr(self, 'df') or self.df is None:
            details_text.insert(tk.END, "No data loaded.\n")
            return
        try:
            # Import required libraries at the beginning of the function
            import pandas as pd
            import numpy as np
            import itertools
            from scipy import stats
            
            x_col = self.xaxis_var.get()
            group_col = self.group_var.get()
            value_cols = [col for var, col in self.value_vars if var.get() and col != x_col]
            if not x_col or not value_cols:
                details_text.insert(tk.END, "No plot or insufficient columns selected.\n")
                return
                
            # Check if we have a working dataframe
            if not hasattr(self, 'df_work') or self.df_work is None:
                details_text.insert(tk.END, "Error calculating statistics: Working dataframe is not available.\n")
                details_text.insert(tk.END, "Please regenerate the plot and then use Generate Statistics.")
                return
                
            # Use self.df_work which has the processed data
            df_plot = self.df_work.copy()
            
            if self.xaxis_renames:
                df_plot[x_col] = df_plot[x_col].map(self.xaxis_renames).fillna(df_plot[x_col])
            if self.xaxis_order:
                df_plot[x_col] = pd.Categorical(df_plot[x_col], categories=self.xaxis_order, ordered=True)
            plot_kind = self.plot_kind_var.get()
            swap_axes = self.swap_axes_var.get()
            # Check if we have any statistics calculated
            if not hasattr(self, 'latest_pvals') or not self.latest_pvals:
                details_text.insert(tk.END, "No statistics have been calculated for this plot.\n")
                details_text.insert(tk.END, "\nUse the 'Generate Statistics' button to calculate statistics and see detailed results.")
                return
            try:
                import pingouin as pg
            except ImportError:
                pg = None
            try:
                import scikit_posthocs as sp
            except ImportError:
                sp = None
             # Collect results for all selected value columns
            # Only process value columns that exist in the DataFrame and are not empty
            valid_value_cols = [col for col in value_cols if col in df_plot.columns and not df_plot[col].dropna().empty]
            if not valid_value_cols:
                details_text.insert(tk.END, "No valid statistics could be calculated.\n")
            import traceback
            for val_col in valid_value_cols:
                    # Detect ungrouped data the same way as calculate_statistics
                    if not group_col or group_col == 'None' or (group_col and len(df_plot[group_col].dropna().unique()) <= 1):
                        # Ungrouped Data: show statistical tests based on x-axis categories
                        # Import pandas for DataFrame operations
                        import pandas as pd
                        import numpy as np
                        
                        # Get x categories (these are the values we're comparing)
                        x_categories = df_plot[x_col].dropna().unique() if x_col in df_plot else []
                        n_x_categories = len(x_categories)
                        
                        # Define n_groups for ungrouped data (implicitly 1 group)
                        n_groups = 1
                        
                        if n_x_categories <= 1:
                            details_text.insert(tk.END, "Only one category: no statistical test performed.\n")
                        elif n_x_categories == 2:
                            # Two-sample t-test between categories
                            cat1, cat2 = x_categories
                            
                            # Get the selected t-test type and alternative from UI
                            ttest_type = self.ttest_type_var.get()
                            alternative = self.ttest_alternative_var.get()
                            
                            # Convert data to numeric format first
                            df_numeric = df_plot.copy()
                            df_numeric[val_col] = pd.to_numeric(df_numeric[val_col], errors='coerce')
                            
                            # Get data for each category
                            data1 = df_numeric[df_numeric[x_col] == cat1][val_col].dropna()
                            data2 = df_numeric[df_numeric[x_col] == cat2][val_col].dropna()
                            
                            details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n\n")
                            
                            # Check if we have p-values from previous calculations
                            p_val = None
                            
                            # First try string-based keys
                            key = self.stat_key(cat1, cat2)
                            if key in self.latest_pvals:
                                p_val = self.latest_pvals[key]
                            
                            # Then try numeric indices (0 and 1 for two categories)
                            if p_val is None:
                                # Try both regular tuples and numpy integer tuples
                                possible_keys = [
                                    (0, 1), (1, 0),  # Regular integer tuples
                                    (np.int64(0), np.int64(1)),  # numpy integer tuples
                                    (np.int64(1), np.int64(0))
                                ]
                                
                                for test_key in possible_keys:
                                    if test_key in self.latest_pvals:
                                        p_val = self.latest_pvals[test_key]
                                        break
                            
                            if p_val is not None:
                                # Get descriptive statistics for each group
                                n1 = len(data1)
                                n2 = len(data2)
                                mean1 = data1.mean()
                                mean2 = data2.mean()
                                std1 = data1.std()
                                std2 = data2.std()
                                sem1 = std1 / np.sqrt(n1) if n1 > 0 else 0
                                sem2 = std2 / np.sqrt(n2) if n2 > 0 else 0
                                
                                # Get the test details if available
                                test_key = None
                                for k in [self.stat_key(cat1, cat2), (0, 1), (1, 0), (np.int64(0), np.int64(1)), (np.int64(1), np.int64(0))]:
                                    if k in self.latest_test_info:
                                        test_key = k
                                        break
                                
                                # Display significance and p-value
                                sig = self.pval_to_annotation(p_val)
                                details_text.insert(tk.END, f"P-value: {p_val:.4g} {sig}\n\n")
                                
                                # Get the current error bar type from the GUI
                                error_type = self.errorbar_type_var.get() if hasattr(self, 'errorbar_type_var') else "SD"
                                
                                # Display descriptive statistics
                                details_text.insert(tk.END, "Group Statistics:\n")
                                details_text.insert(tk.END, "-" * 50 + "\n")
                                
                                # Show header based on selected error type
                                if error_type == "SD":
                                    details_text.insert(tk.END, f"{'Group':<15}{'n':<8}{'Mean':<12}{'SD':<12}\n")
                                    details_text.insert(tk.END, "-" * 50 + "\n")
                                    details_text.insert(tk.END, f"{str(cat1):<15}{n1:<8}{mean1:.4f}{' ':<4}{std1:.4f}\n")
                                    details_text.insert(tk.END, f"{str(cat2):<15}{n2:<8}{mean2:.4f}{' ':<4}{std2:.4f}\n")
                                else:  # SEM
                                    details_text.insert(tk.END, f"{'Group':<15}{'n':<8}{'Mean':<12}{'SEM':<12}\n")
                                    details_text.insert(tk.END, "-" * 50 + "\n")
                                    details_text.insert(tk.END, f"{str(cat1):<15}{n1:<8}{mean1:.4f}{' ':<4}{sem1:.4f}\n")
                                    details_text.insert(tk.END, f"{str(cat2):<15}{n2:<8}{mean2:.4f}{' ':<4}{sem2:.4f}\n")
                                details_text.insert(tk.END, "-" * 50 + "\n\n")
                                
                                # Show detailed test information if available
                                if test_key and test_key in self.latest_test_info:
                                    test_info = self.latest_test_info[test_key]
                                    test_used = test_info.get('test_used', 'Unknown test')
                                    statistic = test_info.get('statistic')
                                    df = test_info.get('df')
                                    
                                    details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n")
                                    details_text.insert(tk.END, f"Actual test performed: {test_used}\n")
                                    
                                    if statistic is not None and df is not None:
                                        details_text.insert(tk.END, f"t = {statistic:.3f}, df = {df}, p = {p_val:.4g} {sig}\n")
                                    elif statistic is not None:
                                        details_text.insert(tk.END, f"t = {statistic:.3f}, p = {p_val:.4g} {sig}\n")
                                
                                details_text.insert(tk.END, "\nNote: These statistics are the same as those used for plot annotations.\n")
                            else:
                                # Calculate new p-value if we don't have one stored
                                if len(data1) > 0 and len(data2) > 0:
                                    try:
                                        if ttest_type == "Independent t-test":
                                            stat, p_val = stats.ttest_ind(data1, data2, alternative=alternative)
                                        else:  # Paired t-test
                                            if len(data1) != len(data2):
                                                details_text.insert(tk.END, "Error: Paired t-test requires equal number of samples in both groups.\n")
                                                return
                                            stat, p_val = stats.ttest_rel(data1, data2, alternative=alternative)
                                        
                                        # Store the new p-value
                                        key = self.stat_key(cat1, cat2)
                                        self.latest_pvals[key] = p_val
                                        
                                        # Get descriptive statistics for each group
                                        n1 = len(data1)
                                        n2 = len(data2)
                                        mean1 = data1.mean()
                                        mean2 = data2.mean()
                                        std1 = data1.std()
                                        std2 = data2.std()
                                        sem1 = std1 / np.sqrt(n1) if n1 > 0 else 0
                                        sem2 = std2 / np.sqrt(n2) if n2 > 0 else 0
                                        
                                        # Display significance and p-value
                                        sig = self.pval_to_annotation(p_val)
                                        details_text.insert(tk.END, f"P-value: {p_val:.4g} {sig}\n\n")
                                        
                                        # Get the current error bar type from the GUI
                                        error_type = self.errorbar_type_var.get() if hasattr(self, 'errorbar_type_var') else "SD"
                                        
                                        # Display descriptive statistics
                                        details_text.insert(tk.END, "Group Statistics:\n")
                                        details_text.insert(tk.END, "-" * 50 + "\n")
                                        
                                        # Show header based on selected error type
                                        if error_type == "SD":
                                            details_text.insert(tk.END, f"{'Group':<15}{'n':<8}{'Mean':<12}{'SD':<12}\n")
                                            details_text.insert(tk.END, "-" * 50 + "\n")
                                            details_text.insert(tk.END, f"{str(cat1):<15}{n1:<8}{mean1:.4f}{' ':<4}{std1:.4f}\n")
                                            details_text.insert(tk.END, f"{str(cat2):<15}{n2:<8}{mean2:.4f}{' ':<4}{std2:.4f}\n")
                                        else:  # SEM
                                            details_text.insert(tk.END, f"{'Group':<15}{'n':<8}{'Mean':<12}{'SEM':<12}\n")
                                            details_text.insert(tk.END, "-" * 50 + "\n")
                                            details_text.insert(tk.END, f"{str(cat1):<15}{n1:<8}{mean1:.4f}{' ':<4}{sem1:.4f}\n")
                                            details_text.insert(tk.END, f"{str(cat2):<15}{n2:<8}{mean2:.4f}{' ':<4}{sem2:.4f}\n")
                                        details_text.insert(tk.END, "-" * 50 + "\n\n")
                                        
                                        # Show test information
                                        details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n")
                                        if ttest_type == "Independent t-test":
                                            details_text.insert(tk.END, f"t = {stat:.3f}, p = {p_val:.4g} {sig}\n")
                                    except Exception as e:
                                        details_text.insert(tk.END, f"Error calculating t-test: {str(e)}\n")
                                else:
                                    details_text.insert(tk.END, "No valid numeric data available for statistical test.\n")
                        elif n_x_categories > 2:
                            # For 3+ categories, display ANOVA + post-hoc test results
                            # Get the stored information about the statistical tests
                            anova_type = self.latest_stats.get('anova_type', "One-way ANOVA")
                            posthoc_type = self.latest_stats.get('posthoc_type', "Tukey's HSD")
                            
                            
                            # Display the main ANOVA result if available
                            anova_results = self.latest_stats.get('anova_results', None)
                            if anova_results is not None:
                                try:
                                    # ANOVA results format may vary based on the test type
                                    if 'p-unc' in anova_results.columns:  # pingouin format
                                        anova_p = anova_results['p-unc'].iloc[0]
                                    elif 'p' in anova_results.columns:  # alternate format
                                        anova_p = anova_results['p'].iloc[0]
                                    else:
                                        anova_p = None
                                        
                                    if anova_p is not None:
                                        details_text.insert(tk.END, f"\n============ ANOVA Results Summary ============\n")
                                        details_text.insert(tk.END, f"Test type: {anova_type}\n")
                                        details_text.insert(tk.END, f"Main ANOVA result: p = {anova_p:.4g} {self.pval_to_annotation(anova_p, alpha=alpha)}\n")
                                        details_text.insert(tk.END, "")
                                        
                                        # Add guidance for non-significant ANOVA results
                                        if anova_p > 0.05:
                                            details_text.insert(tk.END, "\n⚠️ WARNING: The overall ANOVA p-value is not significant (p > 0.05).\n")
                                            details_text.insert(tk.END, "When the main ANOVA is not significant, post-hoc test results should be interpreted with caution\n")
                                            details_text.insert(tk.END, "as there may not be meaningful differences between any groups.\n\n")
                                except Exception as e:
                                    print(f"[DEBUG] Error displaying ANOVA p-value: {e}")
                            
                            # Check if we have p-values from previous calculations
                            print(f"[DEBUG] Statistical Details - latest_pvals keys: {list(self.latest_pvals.keys())}")
                            print(f"[DEBUG] Statistical Details - x_categories: {x_categories}")
                            
                            # Always try to display p-values since they're calculated elsewhere
                            has_pvals = False
                            # Display the p-value matrix
                            p_matrix = pd.DataFrame(index=x_categories, columns=x_categories)
                            
                            # Check key structures in latest_pvals
                            key_lengths = [len(k) if isinstance(k, tuple) else 0 for k in self.latest_pvals.keys()]
                            has_triplet_keys = any(length == 3 for length in key_lengths)
                            has_numeric_keys = any(isinstance(k[0], (int, float, np.integer, np.floating)) 
                                                 for k in self.latest_pvals.keys() 
                                                 if isinstance(k, tuple) and len(k) >= 2)
                            
                            print(f"[DEBUG] Has numeric keys: {has_numeric_keys}")
                            print(f"[DEBUG] Has triplet keys: {has_triplet_keys}")
                            print(f"[DEBUG] Key lengths: {key_lengths}")
                                
                            # Fill matrix with p-values
                            for i, g1 in enumerate(x_categories):
                                for j, g2 in enumerate(x_categories):
                                    if g1 == g2:
                                        p_matrix.loc[g1, g2] = float('nan')  # Diagonal is not applicable
                                    else:
                                        p_val = None
                                        
                                        # First, try standard string-based keys
                                        key1 = self.stat_key(g1, g2)  # Standard key
                                        key2 = (g1, g2)  # Direct tuple
                                        key3 = (g2, g1)  # Reversed tuple

                                        print(f"[DEBUG] Looking for p-value for {g1} vs {g2}: keys {key1}, {key2}, {key3}")
                                        if key1 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key1: {key1}")
                                            p_val = self.latest_pvals[key1]
                                            has_pvals = True
                                        elif key2 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key2: {key2}")
                                            p_val = self.latest_pvals[key2]
                                            has_pvals = True
                                        elif key3 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key3: {key3}")
                                            p_val = self.latest_pvals[key3]
                                            has_pvals = True
                                            
                                        # If not found yet and we have numeric keys, try with indices
                                        if p_val is None and has_numeric_keys:
                                            # Try both orderings of numeric indices (int)
                                            if (i, j) in self.latest_pvals:
                                                print(f"[DEBUG] Found using numeric indices: ({i}, {j})")
                                                p_val = self.latest_pvals[(i, j)]
                                                has_pvals = True
                                            elif (j, i) in self.latest_pvals:
                                                print(f"[DEBUG] Found using numeric indices: ({j}, {i})")
                                                p_val = self.latest_pvals[(j, i)]
                                                has_pvals = True
                                            # Also try with float indices
                                            elif (float(i), float(j)) in self.latest_pvals:
                                                print(f"[DEBUG] Found using float indices: ({float(i)}, {float(j)})")
                                                p_val = self.latest_pvals[(float(i), float(j))]
                                                has_pvals = True
                                            elif (float(j), float(i)) in self.latest_pvals:
                                                print(f"[DEBUG] Found using float indices: ({float(j)}, {float(i)})")
                                                p_val = self.latest_pvals[(float(j), float(i))]
                                                has_pvals = True
                                        
                                        # If still not found, look for numpy numeric indices (int64, float64, etc.)
                                        if p_val is None and has_numeric_keys:
                                            for key, val in self.latest_pvals.items():
                                                # First check if this is a triplet key for metric comparisons
                                                if (isinstance(key, tuple) and len(key) == 3 and
                                                    isinstance(key[0], (np.integer, np.floating)) and
                                                    isinstance(key[1], str) and isinstance(key[2], str)):
                                                    # This is a metric comparison key (idx, metric1, metric2)
                                                    # Skip it here - these are handled separately
                                                    continue
                                                    
                                                # Handle regular numeric indices in pairs
                                                if (isinstance(key, tuple) and len(key) == 2 and
                                                   (isinstance(key[0], (np.integer, np.floating)) or 
                                                    isinstance(key[1], (np.integer, np.floating)))):
                                                    # Check if this key matches our indices (allowing for float/int conversion)
                                                    try:
                                                        key0_matches_i = (float(key[0]) == float(i))
                                                        key0_matches_j = (float(key[0]) == float(j))
                                                        key1_matches_i = (float(key[1]) == float(i))
                                                        key1_matches_j = (float(key[1]) == float(j))
                                                        
                                                        if (key0_matches_i and key1_matches_j) or (key0_matches_j and key1_matches_i):
                                                            print(f"[DEBUG] Found using numpy numeric: {key}")
                                                            p_val = val
                                                            has_pvals = True
                                                            break
                                                    except ValueError:
                                                        # Skip keys that can't be converted to float
                                                        continue
                                        
                                        # Before giving up, check triplet keys that might contain metric comparisons
                                        if p_val is None and has_triplet_keys:
                                            # Extract all available metric pairs from triplet keys
                                            metric_pairs = []
                                            x_values = []
                                            for key in self.latest_pvals.keys():
                                                if isinstance(key, tuple) and len(key) == 3:
                                                    x_val = key[0]
                                                    metric1 = key[1]
                                                    metric2 = key[2]
                                                    x_values.append(x_val)
                                                    if (metric1, metric2) not in metric_pairs and (metric2, metric1) not in metric_pairs:
                                                        metric_pairs.append((metric1, metric2))
                                            
                                            if metric_pairs and x_values:
                                                # Check for matching x-indices for this category pair
                                                potential_i = None
                                                potential_j = None
                                                
                                                # Try to match string category to numeric index
                                                for idx, x_val in enumerate(x_categories):
                                                    if x_val == g1 or str(x_val) == str(i):
                                                        potential_i = idx
                                                    if x_val == g2 or str(x_val) == str(j):
                                                        potential_j = idx
                                                
                                                if potential_i is not None and potential_j is not None:
                                                    # Try all metric pairs with all potential indices
                                                    for pair in metric_pairs:
                                                        metric1, metric2 = pair
                                                        # Try forward and reversed tuples with np.int64
                                                        test_keys = [
                                                            (np.int64(potential_i), metric1, metric2),
                                                            (np.int64(potential_j), metric1, metric2)
                                                        ]
                                                        
                                                        for test_key in test_keys:
                                                            if test_key in self.latest_pvals:
                                                                p_val = self.latest_pvals[test_key]
                                                                has_pvals = True
                                                                print(f"[DEBUG] Found using triplet key: {test_key} -> {p_val}")
                                                                break
                                                        
                                                        if p_val is not None:
                                                            break
                                        
                                        # Store the p-value in the matrix if found
                                        if p_val is not None:
                                            p_matrix.loc[g1, g2] = p_val
                                        else:
                                            p_matrix.loc[g1, g2] = float('nan')
                                
                            # For multiple y-axis columns, we handle statistics differently
                            # to avoid redundant analyses and ensure consistent results
                            metrics_analysis_done = False  # Track if we've already done the metrics analysis
                            
                            # Check if we have triplet keys with metric comparisons
                            if has_triplet_keys:
                                # Get the value columns that were selected in the UI
                                value_cols = [col for var, col in self.value_vars if var.get()]
                                value_cols = [col for col in value_cols if col]
                                self.debug(f"Multiple metrics detected: {value_cols}")
                                
                                # Clear the text widget and start with a fresh header
                                details_text.delete(1.0, tk.END)
                                self._add_significance_legend(details_text)
                                
                                # Collect all unique metric pairs for comparison
                                metric_pairs = []
                                metric_results = {}
                                
                                for key, val in self.latest_pvals.items():
                                    if isinstance(key, tuple) and len(key) == 3:
                                        x_val = key[0]
                                        metric1 = key[1]
                                        metric2 = key[2]
                                        metric_pair = (metric1, metric2)
                                        
                                        if metric_pair not in metric_pairs and (metric2, metric1) not in metric_pairs:
                                            metric_pairs.append(metric_pair)
                                        
                                        # Store with x_category name if possible
                                        x_index = None
                                        try:
                                            x_index = int(float(x_val))
                                            if 0 <= x_index < len(x_categories):
                                                x_name = x_categories[x_index]
                                            else:
                                                x_name = str(x_val)
                                        except:
                                            x_name = str(x_val)
                                            
                                        result_key = (x_name, metric1, metric2)
                                        metric_results[result_key] = val
                                
                                if metric_pairs and not metrics_analysis_done:
                                    metrics_analysis_done = True  # Mark that we've performed the metrics analysis
                                    
                                    # Get the selected ANOVA and post-hoc test types from UI settings
                                    anova_type = self.anova_type_var.get()
                                    posthoc_type = self.posthoc_type_var.get()
                                    
                                    details_text.insert(tk.END, f"\n============ Multiple Y-Axis Columns Analysis ============\n")
                                    details_text.insert(tk.END, f"Comparing {len(value_cols)} metrics across {len(x_categories)} categories\n")
                                    
                                    # Create an appropriate melted dataframe for statistical analysis
                                    df_analysis = None
                                    
                                    # First try to make a fresh melted DataFrame for better statistics
                                    try:
                                        # Get the necessary columns
                                        if hasattr(self, 'xaxis_var') and len(value_cols) > 1 and self.xaxis_var.get() in df_plot.columns:
                                            x_col = self.xaxis_var.get()
                                            
                                            # Create a subset of the dataframe with just what we need
                                            cols_to_keep = [x_col] + value_cols
                                            df_subset = df_plot[cols_to_keep].copy()
                                            
                                            # Melt the data into long format
                                            id_vars = [x_col]
                                            df_analysis = pd.melt(df_subset, 
                                                                id_vars=id_vars, 
                                                                value_vars=value_cols, 
                                                                var_name='Measurement', 
                                                                value_name='MeltedValue')
                                            
                                            # Ensure values are numeric
                                            df_analysis['MeltedValue'] = pd.to_numeric(df_analysis['MeltedValue'], errors='coerce')
                                            df_analysis = df_analysis.dropna(subset=['MeltedValue'])
                                            
                                            self.debug(f"Created fresh analysis DataFrame, shape: {df_analysis.shape}")
                                    except Exception as e_frame:
                                        self.debug(f"Error creating analysis DataFrame: {e_frame}\n{traceback.format_exc()}")
                                        df_analysis = None
                                    
                                    # Perform the statistical analysis
                                    if df_analysis is not None and len(df_analysis) > 0:
                                        try:
                                            # Import statistical functions
                                            from explot_stats import run_anova, run_posthoc, run_ttest
                                            
                                            # Get unique metrics
                                            metrics = df_analysis['Measurement'].unique()
                                            
                                            # Use t-test for exactly 2 y-columns, ANOVA for 3+ columns
                                            if len(metrics) == 2:
                                                # For 2 metrics, use t-test instead of ANOVA
                                                test_type = "Welch's t-test (unpaired, unequal variances)"
                                                
                                                # Extract the two metrics
                                                metric1, metric2 = metrics
                                                
                                                # Run the t-test
                                                p_val, t_result = run_ttest(df_analysis, 'MeltedValue', metric1, metric2, 'Measurement', test_type)
                                                
                                                # Display t-test result
                                                details_text.insert(tk.END, f"T-test result: p = {p_val:.4f} {p_val:.2e}\n\n")
                                                
                                                # Add a note about interpretation
                                                if p_val > 0.05:
                                                    details_text.insert(tk.END, "⚠️ NOTE: The t-test p-value is not significant (p > 0.05).\n")
                                                    details_text.insert(tk.END, "This suggests there is no statistically significant difference between the two metrics.\n\n")
                                                else:
                                                    details_text.insert(tk.END, "The t-test p-value is significant (p ≤ 0.05).\n")
                                                    details_text.insert(tk.END, "This suggests there is a statistically significant difference between the two metrics.\n\n")
                                                
                                                # Show test type information
                                                details_text.insert(tk.END, f"Test Used: {test_type} between the two metrics\n\n")
                                                
                                                # Create a simple posthoc-like result for consistent downstream processing
                                                posthoc = pd.DataFrame(data=[[np.nan, p_val], [p_val, np.nan]], 
                                                                   index=[metric1, metric2], 
                                                                   columns=[metric1, metric2])
                                            else:
                                                # Run ANOVA on metrics (3+ columns)
                                                anova_result = run_anova(df_analysis, 'MeltedValue', 'Measurement', anova_type=anova_type)
                                                
                                                if anova_result is not None:
                                                    # Display main ANOVA result
                                                    if 'p-unc' in anova_result.columns:
                                                        p_val = anova_result['p-unc'].iloc[0]
                                                    else:
                                                        p_val = anova_result.iloc[0, -1]
                                                        
                                                    details_text.insert(tk.END, f"Main ANOVA result: p = {p_val:.4f} {p_val:.2e}\n\n")
                                                    
                                                    # Add warning if not significant
                                                    if p_val > 0.05:
                                                        details_text.insert(tk.END, "⚠️ WARNING: The overall ANOVA p-value is not significant (p > 0.05).\n")
                                                        details_text.insert(tk.END, "When the main ANOVA is not significant, post-hoc test results should be interpreted with caution\n")
                                                        details_text.insert(tk.END, "as there may not be meaningful differences between any groups.\n\n")
                                                    
                                                    # Show test type information
                                                    details_text.insert(tk.END, f"Test Used: {anova_type} + {posthoc_type} across {len(metrics)} metrics\n\n")
                                                
                                                # Run post-hoc tests
                                                posthoc = run_posthoc(df_analysis, 'MeltedValue', 'Measurement', posthoc_type=posthoc_type)
                                                
                                                if posthoc is not None:
                                                    # Format and display the post-hoc results matrix
                                                    details_text.insert(tk.END, f"Post-hoc {posthoc_type} test results:\n")
                                                    details_text.insert(tk.END, "-" * 60 + "\n")
                                                    details_text.insert(tk.END, self.format_pvalue_matrix(posthoc) + '\n')
                                                    
                                                    # Show nicely formatted significance indicators
                                                    details_text.insert(tk.END, "\nSignificance indicators for pairwise comparisons:\n")
                                                    details_text.insert(tk.END, "-" * 60 + "\n")
                                                    
                                                    # Only show upper triangle to avoid duplication
                                                    for i, m1 in enumerate(posthoc.index):
                                                        for j, m2 in enumerate(posthoc.columns):
                                                            if i < j:  # Upper triangle only
                                                                try:
                                                                    p_val = posthoc.loc[m1, m2]
                                                                    if pd.notna(p_val):
                                                                        sig = self.pval_to_annotation(p_val)
                                                                        
                                                                        # Format p-value for display
                                                                        if abs(p_val) < 1e-10:
                                                                            p_text = "p < 1e-10"
                                                                        elif p_val < 0.0001:
                                                                            p_text = f"p = {p_val:.2e}"
                                                                        else:
                                                                            p_text = f"p = {p_val:.4f}"
                                                                        
                                                                        details_text.insert(tk.END, f"{m1} vs {m2}: {p_text} {sig}\n")
                                                                except Exception as e:
                                                                    self.debug(f"Error displaying posthoc p-value: {e}")
                                                    
                                                    # Add note about consistency with plot annotations
                                                    details_text.insert(tk.END, "\nNote: These statistics are the same as those used for plot annotations.\n")
                                                else:
                                                    details_text.insert(tk.END, "Error running post-hoc tests. Unable to calculate pairwise comparisons.\n")
                                                    
                                                    # This handles both the case where t-test results or ANOVA results are None
                                                    if len(metrics) == 2:
                                                        details_text.insert(tk.END, "Error running t-test. Unable to determine significance.\n")
                                                    else:
                                                        details_text.insert(tk.END, "Error running ANOVA. Unable to determine overall significance.\n")
                                                    # Fall back to showing the raw p-values
                                                    self._show_old_metric_comparisons(details_text, metric_pairs, metric_results, x_categories)
                                        except Exception as e_stats:
                                            self.debug(f"Error in ANOVA analysis: {e_stats}\n{traceback.format_exc()}")
                                            details_text.insert(tk.END, f"Error in statistical analysis: {e_stats}\n\n")
                                            # Fall back to showing the raw p-values
                                            self._show_old_metric_comparisons(details_text, metric_pairs, metric_results, x_categories)
                                    else:
                                        # If we couldn't create a proper analysis dataframe, use the old method
                                        details_text.insert(tk.END, "Could not create proper analysis dataframe. Showing raw p-values:\n\n")
                                        self._show_old_metric_comparisons(details_text, metric_pairs, metric_results, x_categories)
                                    
                                    # Skip the rest of the statistical display to avoid duplications
                                    # Set has_pvals to indicate we have successfully showed the statistics
                                    has_pvals = True
                                    
                                    # Return early to avoid showing redundant statistical details
                                    return
                            
                            # Format and display p-value matrix if we found values
                            if not has_triplet_keys:  # Skip this section entirely if we're dealing with metric comparisons
                                if has_pvals:
                                    details_text.insert(tk.END, f"Test Used: {anova_type} + {posthoc_type} across {n_x_categories} categories\n")
                                    
                                    # We'll show the formatted p-value matrix in the post-hoc test section below
                                    # This avoids duplication of the same information
                                
                                # Note about using same p-values as annotations
                                details_text.insert(tk.END, "\nNote: These statistics are the same as those used for plot annotations.\n")
                            elif not has_triplet_keys:  # Only show error if we haven't already displayed metric comparisons
                                details_text.insert(tk.END, "No matching p-values found in latest_pvals dictionary.\n")
                                details_text.insert(tk.END, f"Keys available: {list(self.latest_pvals.keys())[:5]}\n")
                    else:
                        # Grouped Data: multiple groups in a dataframe
                        unique_groups = df_plot[group_col].dropna().unique() if group_col in df_plot else []
                        x_categories = df_plot[x_col].dropna().unique() if x_col in df_plot else []
                        n_groups = len(unique_groups)
                        n_x_categories = len(x_categories)
                        
                        # We'll show the detailed results in the ANOVA and post-hoc sections below
                        # So we'll skip the individual comparisons here
                        
                        # Initialize posthoc_matrices if not already done
                        if 'posthoc_matrices' not in locals():
                            posthoc_matrices = {}
                        
                        # Initialize desc_stats if not already done
                        if 'desc_stats' not in locals():
                            desc_stats = {}
                        
                        # Initialize error_label if not already done
                        if 'error_label' not in locals():
                            error_label = "SEM"  # Default value
                        
                        # Now we'll let the code continue to the ANOVA and post-hoc sections below
                        # The rest of the statistical details will be handled there
                        pass
                        
                    # End of grouped data section
                    
                    # The code will now continue to the ANOVA and post-hoc test sections below
                    # which will display the results in a clean, tabular format
                    
                    # Display available keys for debugging if no results found
                    if not hasattr(self, 'latest_pvals') or not self.latest_pvals:
                        details_text.insert(tk.END, "No p-values stored. Please generate statistics first.\n")
                    
                    # Handle the case of a single group with multiple x-categories
                    if n_groups == 1:
                        if n_x_categories == 1:
                            details_text.insert(tk.END, "Only one category: no statistical test performed.\n")
                        elif n_x_categories == 2:
                            # Two-sample t-test between categories
                            cat1, cat2 = x_categories
                            # Convert data to numeric, handling potential errors
                            def convert_to_numeric(series):
                                try:
                                    return pd.to_numeric(series, errors='coerce').dropna()
                                except Exception as e:
                                    details_text.insert(tk.END, f"Error converting data to numeric: {e}\n")
                                    return pd.Series(dtype=float)

                            df_cat1 = convert_to_numeric(df_plot[df_plot[x_col] == cat1][val_col])
                            df_cat2 = convert_to_numeric(df_plot[df_plot[x_col] == cat2][val_col])

                            # Check if we have enough valid numeric data
                            if len(df_cat1) == 0 or len(df_cat2) == 0:
                                details_text.insert(tk.END, f"Insufficient numeric data for t-test between {cat1} and {cat2}\n")
                                return

                            # Get the selected t-test type and alternative from UI
                            ttest_type = self.ttest_type_var.get()
                            alternative = self.ttest_alternative_var.get()
                            
                            details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n")
                            try:
                                # Use the explot_stats module for consistency
                                from explot_stats import run_ttest
                                
                                # Create a temporary DataFrame for the t-test
                                temp_df = pd.DataFrame({
                                    'group': [cat1] * len(df_cat1) + [cat2] * len(df_cat2),
                                    'value': np.concatenate([df_cat1, df_cat2])
                                })
                                
                                # Use run_ttest from explot_stats module
                                p_val, ttest_results = run_ttest(
                                    temp_df, 'value', cat1, cat2, 'group', 
                                    test_type=ttest_type, alternative=alternative
                                )
                                
                                # Display detailed test information
                                if isinstance(ttest_results, dict):
                                    # Use the actual test that was performed
                                    if 'test_used' in ttest_results:
                                        actual_test = ttest_results['test_used']
                                        details_text.insert(tk.END, f"Actual test performed: {actual_test}\n")
                                    
                                    # Show test statistics
                                    p_val = ttest_results.get('p-val', 1.0)
                                    t_val = ttest_results.get('t', 'N/A')
                                    df_val = ttest_results.get('df', 'N/A')
                                    p_annotation = self.pval_to_annotation(p_val)
                                    details_text.insert(tk.END, f"t = {t_val if t_val == 'N/A' else f'{t_val:.4g}'}, ")
                                    details_text.insert(tk.END, f"df = {df_val if df_val == 'N/A' else f'{df_val:.4g}'}, ")
                                    details_text.insert(tk.END, f"p = {p_val:.4g} {p_annotation}\n")
                                else:
                                    # Fallback for backward compatibility
                                    p_annotation = self.pval_to_annotation(p_val)
                                    details_text.insert(tk.END, f"p = {p_val:.4g} {p_annotation}\n")
                            except Exception as e:
                                details_text.insert(tk.END, f"T-test failed: {e}\n")
                            
                            # End of t-test section
                            
                        elif n_x_categories > 2 and pg is not None:
                            # Get the selected ANOVA test type from settings
                            anova_type = self.anova_type_var.get()
                            posthoc_type = self.posthoc_type_var.get()
                            
                            # Check if we have stored ANOVA results from previous calculation
                            stored_anova_type = self.latest_stats.get('anova_type', anova_type)
                            stored_posthoc_type = self.latest_stats.get('posthoc_type', posthoc_type)
                            
                            details_text.insert(tk.END, f"Test Used: {stored_anova_type} across x-axis categories\n")
                            anova_success = False
                            
                            try:
                                # Convert data to numeric format first
                                df_long = df_plot.melt(
                                    id_vars=[x_col], 
                                    value_vars=[val_col], 
                                    var_name='Condition', 
                                    value_name='MeltedValue'
                                )
                                
                                # Explicitly convert to numeric and drop NA values
                                df_long['MeltedValue'] = pd.to_numeric(df_long['MeltedValue'], errors='coerce')
                                df_long = df_long.dropna(subset=['MeltedValue'])
                                
                                # Check if we have enough valid numeric data
                                if len(df_long) == 0:
                                    details_text.insert(tk.END, "ANOVA failed: No valid numeric data after conversion\n")
                                else:
                                    # Use stored ANOVA results if available, otherwise calculate new ones
                                    aov = self.latest_stats.get('anova_results', None)
                                    if aov is None:
                                        # Perform the ANOVA based on the selected type
                                        from explot_stats import run_anova
                                        
                                        # Create a long format dataframe for the ANOVA
                                        df_long = pd.DataFrame({
                                            x_col: np.concatenate([
                                                [cat] * len(df_plot[df_plot[x_col] == cat][val_col].dropna()) 
                                                for cat in x_categories
                                            ]),
                                            'MeltedValue': np.concatenate([
                                                df_plot[df_plot[x_col] == cat][val_col].dropna().values 
                                                for cat in x_categories
                                            ])
                                        })
                                        
                                        try:
                                            # Run ANOVA using the explot_stats module
                                            if anova_type == "Repeated measures ANOVA":
                                                # Need subject column for repeated measures ANOVA
                                                details_text.insert(tk.END, "Using results from prior calculation.\n")
                                            else:
                                                # Run ANOVA using the explot_stats module
                                                aov = run_anova(df_long, 'MeltedValue', x_col, anova_type)
                                                details_text.insert(tk.END, str(aov) + '\n')
                                                anova_success = True
                                        except ImportError as e_import:
                                            # Fall back to direct calculations if module not available
                                            try:
                                                if anova_type == "Welch's ANOVA":
                                                    aov = pg.welch_anova(
                                                        data=df_long, 
                                                        dv='MeltedValue', 
                                                        between=x_col
                                                    )
                                                elif anova_type == "Repeated measures ANOVA":
                                                    # For repeated measures, we need a subject identifier
                                                    df_long['Subject'] = np.arange(len(df_long))
                                                    aov = pg.rm_anova(
                                                        data=df_long, 
                                                        dv='MeltedValue', 
                                                        within=x_col,
                                                        subject='Subject',
                                                        detailed=True
                                                    )
                                                else:  # Regular one-way ANOVA
                                                    aov = pg.anova(
                                                        data=df_long, 
                                                        dv='MeltedValue', 
                                                        between=x_col,
                                                        detailed=True
                                                    )
                                                
                                                details_text.insert(tk.END, str(aov) + '\n')
                                                anova_success = True
                                                
                                            except Exception as e2:
                                                details_text.insert(tk.END, f"ANOVA calculation failed: {e2}\n")
                                                anova_success = False
                                        except Exception as e_df:
                                            details_text.insert(tk.END, f"Error creating data frame for ANOVA: {e_df}\n")
                                            anova_success = False
                                    if aov is not None:
                                        anova_success = True
                                        
                            except Exception as e:
                                details_text.insert(tk.END, f"ANOVA processing failed: {e}\n")
                                anova_success = False
                                
                            # Only run post-hoc test if ANOVA was successful
                            if anova_success:
                                # Use stored post-hoc results if available, otherwise calculate new ones
                                posthoc = self.latest_stats.get('posthoc_results', None)
                                
                                # Only proceed if we have posthoc results or can calculate them
                                if posthoc is None:
                                    try:
                                        # Try to import and run the posthoc analysis
                                        from explot_stats import run_posthoc
                                        posthoc = run_posthoc(df_long, 'MeltedValue', x_col, posthoc_type)
                                        
                                        # Ensure index and columns are strings for consistency
                                        posthoc.index = posthoc.index.astype(str)
                                        posthoc.columns = posthoc.columns.astype(str)
                                        
                                        # Store these values in latest_pvals for display
                                        # Check if df_long[x_col] is a DataFrame or Series
                                        try:
                                            if isinstance(df_long[x_col], pd.DataFrame):
                                                # If it's a DataFrame, get unique values from the first column
                                                groups = df_long[x_col].iloc[:, 0].unique()
                                            else:
                                                # Otherwise, assume it's a Series
                                                groups = df_long[x_col].unique()
                                                
                                            for i, g1 in enumerate(groups):
                                                for j, g2 in enumerate(groups):
                                                    if i != j:  # Skip diagonal
                                                        # Store p-value for this pair
                                                        pval = posthoc.loc[g1, g2]
                                                        key = self.stat_key(g1, g2)
                                                        self.latest_pvals[key] = pval
                                                        self.latest_stats['pvals'][key] = pval
                                        except Exception as e:
                                            import traceback
                                            self.debug(f"Error accessing groups: {e}\n{traceback.format_exc()}")
                                            # Create an empty set of groups as fallback
                                            groups = []
                                    except ImportError:
                                        details_text.insert(tk.END, "Error: The explot_stats module is required for post-hoc analysis.\n")
                                        details_text.insert(tk.END, "Please install the module or contact the developer.\n")
                                        # Skip the rest of the post-hoc analysis if we can't import the module
                                        posthoc = None
                                
                                # Only proceed with displaying results if we have valid posthoc data
                                if posthoc is not None:
                                    try:
                                        details_text.insert(tk.END, f"\nPost-hoc {stored_posthoc_type} test results:\n")
                                        details_text.insert(tk.END, "-" * 60 + "\n")
                                        # Format the posthoc matrix with improved handling of very small p-values
                                        details_text.insert(tk.END, self.format_pvalue_matrix(posthoc) + '\n')
                                        
                                        # Add a more readable version with significance indicators
                                        details_text.insert(tk.END, "\nSignificance indicators for pairwise comparisons:\n")
                                        details_text.insert(tk.END, "-" * 60 + "\n")
                                        
                                        # Track if any extremely small p-values were found
                                        has_very_small_pvals = False
                                        
                                        for idx1, group1 in enumerate(posthoc.index):
                                            for idx2, group2 in enumerate(posthoc.columns):
                                                if idx1 < idx2:  # Only show each comparison once (upper triangle)
                                                    try:
                                                        p_val = posthoc.loc[group1, group2]
                                                        if pd.notna(p_val):
                                                            sig = self.pval_to_annotation(p_val)
                                                            
                                                            # Format p-value for display
                                                            if abs(p_val) < 1e-10:  # Extremely small p-value
                                                                p_text = "p < 1e-10"
                                                                has_very_small_pvals = True
                                                            elif p_val < 0.0001:  # Small p-value
                                                                p_text = f"p = {p_val:.2e}"
                                                            else:  # Regular p-value
                                                                p_text = f"p = {p_val:.4f}"
                                                            
                                                            details_text.insert(tk.END, f"{group1} vs {group2}: {p_text} {sig}\n")
                                                    except Exception as e:
                                                        print(f"[DEBUG] Error formatting post-hoc result for {group1} vs {group2}: {e}")
                                                        continue
                                        
                                        # Add note about extremely small p-values if any were found
                                        if has_very_small_pvals:
                                            details_text.insert(tk.END, "\nNote: p < 1e-10 indicates an extremely small p-value\n")
                                            details_text.insert(tk.END, "that is effectively zero in floating point precision.\n")
                                    except Exception as e:
                                        details_text.insert(tk.END, f"Error displaying post-hoc results: {e}\n")
                                else:
                                    details_text.insert(tk.END, "ANOVA/post-hoc pipeline requires required packages.\n")

                    # Handle multiple groups case (more than 1 group)
                    if group_col and group_col.strip() != '' and group_col != 'None':
                        unique_groups = df_plot[group_col].dropna().unique() if group_col in df_plot else []
                        x_categories = df_plot[x_col].dropna().unique() if x_col in df_plot else []
                        n_groups = len(unique_groups)
                        if n_groups > 1:  # If we have more than 1 group (multiple group comparison)
                            base_groups = [g for g in df_plot[x_col].dropna().unique()]
                            hue_groups = [g for g in df_plot[group_col].dropna().unique()]
                        n_hue = len(hue_groups)
                        
                        # Show which test is being used
                        if n_hue > 2 and pg is not None and sp is not None:
                            # Get the selected ANOVA and post-hoc test types from UI
                            anova_type = self.anova_type_var.get()
                            posthoc_type = self.posthoc_type_var.get()
                            details_text.insert(tk.END, f"Test Used: {anova_type} + {posthoc_type} for multiple groups\n")
                        elif n_hue == 2:
                            # Get the selected t-test type and alternative from UI
                            ttest_type = self.ttest_type_var.get()
                            alternative = self.ttest_alternative_var.get()
                            details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) for multiple groups\n")
                            
                            # Display t-test results for each x-category (exactly 2 groups case)
                            details_text.insert(tk.END, "\n== T-TEST RESULTS FOR EACH CATEGORY ==\n\n")
                            
                            # For each x-axis category, perform a t-test between the two groups
                            for x_cat in x_categories:
                                details_text.insert(tk.END, f"\nCategory: {x_cat}\n{'-'*40}\n")
                                
                                # Get data for each group in this category
                                df_cat = df_plot[df_plot[x_col] == x_cat]
                                if len(df_cat) == 0:
                                    details_text.insert(tk.END, "No data available for this category\n")
                                    continue
                                
                                # Get the two groups
                                if len(hue_groups) != 2:
                                    details_text.insert(tk.END, f"Expected exactly 2 groups, found {len(hue_groups)}\n")
                                    continue
                                    
                                group1, group2 = hue_groups[0], hue_groups[1]
                                
                                # Look for p-value in latest_pvals using various key formats
                                p_val = None
                                key_formats = [
                                    (x_cat, group1, group2),
                                    (x_cat, group2, group1),
                                    str((x_cat, group1, group2)),
                                    str((x_cat, group2, group1)),
                                ]
                                
                                # Try numpy int64 keys if x_cat is numeric
                                try:
                                    x_idx = list(x_categories).index(x_cat)
                                    key_formats.extend([
                                        (np.int64(x_idx), group1, group2),
                                        (np.int64(x_idx), group2, group1)
                                    ])
                                except (ValueError, TypeError):
                                    pass
                                
                                # Check all key formats
                                for key in key_formats:
                                    if key in self.latest_pvals:
                                        p_val = self.latest_pvals[key]
                                        break
                                
                                # Extract group data for this category
                                def convert_to_numeric(series):
                                    try:
                                        return pd.to_numeric(series, errors='coerce').dropna()
                                    except Exception as e:
                                        details_text.insert(tk.END, f"Error converting data to numeric: {e}\n")
                                        return pd.Series(dtype=float)
                                
                                vals1 = convert_to_numeric(df_cat[df_cat[group_col] == group1][val_col])
                                vals2 = convert_to_numeric(df_cat[df_cat[group_col] == group2][val_col])
                                
                                # Skip if insufficient data
                                if len(vals1) == 0 or len(vals2) == 0:
                                    details_text.insert(tk.END, f"Insufficient data for {group1} vs {group2}\n")
                                    continue
                                
                                # Display t-test result
                                if p_val is not None:
                                    sig = self.pval_to_annotation(p_val)
                                    details_text.insert(tk.END, f"P-value: {p_val:.4g} {sig}\n")
                                    
                                    # Calculate statistics
                                    means = [vals1.mean(), vals2.mean()]
                                    stds = [vals1.std(ddof=1), vals2.std(ddof=1)]
                                    ns = [len(vals1), len(vals2)]
                                    sems = [std / np.sqrt(n) for std, n in zip(stds, ns)]
                                    
                                    # Get error type from settings (SD or SEM)
                                    error_type = self.errorbar_type_var.get().lower()
                                    errors = stds if error_type == 'sd' else sems
                                    error_label = 'SD' if error_type == 'sd' else 'SEM'
                                    
                                    # Display the means and errors in a table
                                    details_text.insert(tk.END, f"\n{'Group':<10} {'Mean':<10} {error_label+' ±':<10} {'n':<5}\n")
                                    details_text.insert(tk.END, "-"*40 + "\n")
                                    for group, mean, error, n in zip([group1, group2], means, errors, ns):
                                        details_text.insert(tk.END, f"{group:<10} {mean:<10.4g} {error:<10.4g} {n:<5}\n")
                                else:
                                    details_text.insert(tk.END, f"No p-value found for {group1} vs {group2} in category {x_cat}\n")
                        else:
                            details_text.insert(tk.END, "Multiple groups present, but no suitable test could be performed.\n")
                            continue
                                                
                        # Get the selected ANOVA and post-hoc test types from UI
                        anova_type = self.anova_type_var.get()
                        posthoc_type = self.posthoc_type_var.get()
                        
                        # Store all p-value matrices for each category
                        all_pval_matrices = {}
                        anova_results = {}
                        
                        # Process each x-axis category and perform separate ANOVA tests
                        for i, g in enumerate(base_groups):
                            df_sub = df_plot[df_plot[x_col] == g]
                            pairs = list(itertools.combinations(hue_groups, 2))
                            
                            # For ANOVA + post-hoc test, build a matrix
                            if n_hue > 2:
                                pval_matrix = pd.DataFrame(index=hue_groups, columns=hue_groups, dtype=object)
                                
                                # Perform a separate ANOVA test for this x-category
                                try:
                                    # First ensure the data is numeric
                                    df_sub_clean = df_sub.copy()
                                    # Convert the value column to numeric, making a safe copy
                                    try:
                                        df_sub_clean[val_col] = pd.to_numeric(df_sub_clean[val_col], errors='coerce')
                                        # Drop any rows with NaN values after conversion
                                        df_sub_clean = df_sub_clean.dropna(subset=[val_col])
                                    except Exception as e:
                                        print(f"[DEBUG] Error converting values to numeric: {e}")
                                    
                                    # Verify we have enough data for ANOVA
                                    if len(df_sub_clean) < 3 or len(df_sub_clean[group_col].unique()) < 2:
                                        print(f"[DEBUG] Not enough data for ANOVA in category {g}")
                                        aov = None
                                    else:
                                        # Use explot_stats module for consistent ANOVA calculations
                                        try:
                                            # Import directly here to avoid circular imports
                                            from explot_stats import run_anova
                                            
                                            # Run ANOVA using the explot_stats module
                                            if anova_type == "Repeated measures ANOVA":
                                                # For repeated measures, we need a subject identifier
                                                df_sub_clean['Subject'] = np.arange(len(df_sub_clean))
                                                # Run repeated measures ANOVA
                                                aov = run_anova(df_sub_clean, val_col, group_col, anova_type, subject_col='Subject')
                                            else:
                                                # Run one-way or Welch's ANOVA 
                                                aov = run_anova(df_sub_clean, val_col, group_col, anova_type)
                                            
                                            # Debug output
                                            if 'p-unc' in aov.columns:
                                                print(f"[DEBUG] ANOVA for {g}: p = {aov['p-unc'].iloc[0]}")
                                            elif 'p' in aov.columns:
                                                print(f"[DEBUG] ANOVA for {g}: p = {aov['p'].iloc[0]}")
                                                
                                        except ImportError:
                                            # Fall back to direct calculations if module not available
                                            details_text.insert(tk.END, "Error: The explot_stats module is required for statistical analysis.\n")
                                            details_text.insert(tk.END, "Please install the module or contact the developer.\n")
                                            raise
                                    
                                    # Use only explot_stats module for post-hoc tests
                                    try:
                                        from explot_stats import run_posthoc
                                        # Use the same cleaned numeric dataframe that we used for ANOVA
                                        print(f"[DEBUG] Running {posthoc_type} post-hoc test for {g}")
                                        # Double check that the data is numeric before passing it to post-hoc test
                                        test_df = df_sub_clean.copy()
                                        test_df[val_col] = pd.to_numeric(test_df[val_col], errors='coerce')
                                        test_df = test_df.dropna(subset=[val_col])
                                        posthoc = run_posthoc(test_df, val_col, group_col, posthoc_type)
                                        print(f"[DEBUG] Successfully completed {posthoc_type} post-hoc test for {g}")
                                        
                                        # Convert index and columns to string for consistency if needed
                                        if posthoc is not None and not posthoc.empty:
                                            if not isinstance(posthoc.index[0], str):
                                                posthoc.index = posthoc.index.astype(str)
                                            if not isinstance(posthoc.columns[0], str):
                                                posthoc.columns = posthoc.columns.astype(str)
                                    except ImportError:
                                        details_text.insert(tk.END, "Error: The explot_stats module is required for post-hoc tests.\n")
                                        details_text.insert(tk.END, "Please install the module or contact the developer.\n")
                                        raise
                                    except Exception as e:
                                        print(f"[DEBUG] Post-hoc test {posthoc_type} failed: {str(e)}")
                                        details_text.insert(tk.END, f"Post-hoc test {posthoc_type} failed: {str(e)}\n")
                                        posthoc = None
                                        
                                    # If posthoc test failed, create a fallback with t-tests
                                    if posthoc is None or posthoc.empty:
                                        # Create an empty DataFrame for the post-hoc results
                                        posthoc = pd.DataFrame(index=df_sub_clean[group_col].unique(), columns=df_sub_clean[group_col].unique())
                                        
                                        # Manually calculate pairwise p-values using t-tests
                                        print(f"[DEBUG] Fallback: Using pairwise t-tests for {g} since post-hoc test failed")
                                        for h1 in df_sub_clean[group_col].unique():
                                            for h2 in df_sub_clean[group_col].unique():
                                                if h1 != h2:  # Skip the diagonal
                                                    try:
                                                        # Get data for each group using the cleaned numeric dataframe
                                                        data1 = df_sub_clean[df_sub_clean[group_col] == h1][val_col]
                                                        data2 = df_sub_clean[df_sub_clean[group_col] == h2][val_col]
                                                        
                                                        # Skip if insufficient data
                                                        if len(data1) < 2 or len(data2) < 2:
                                                            print(f"[DEBUG] Insufficient data for t-test in {g}: {h1} vs {h2}")
                                                            continue
                                                            
                                                        # Perform Welch's t-test
                                                        from scipy import stats
                                                        t_stat, p_val = stats.ttest_ind(data1, data2, equal_var=False)
                                                        
                                                        # Store in the posthoc matrix
                                                        posthoc.loc[h1, h2] = p_val
                                                        print(f"[DEBUG] Fallback t-test for {g}: {h1} vs {h2}: p = {p_val:.4g}")
                                                    except Exception as e:
                                                        print(f"[DEBUG] Error in fallback t-test: {e}")
                                                        
                                        # Check if we successfully calculated any p-values
                                        if posthoc.isnull().all().all():
                                            print(f"[DEBUG] All fallback t-tests failed for {g}")
                                            posthoc = None
                                        else:
                                            # Fill diagonal with 1.0 (same group comparison)
                                            for h in posthoc.index:
                                                posthoc.loc[h, h] = 1.0
                                            print(f"[DEBUG] Fallback post-hoc result matrix created successfully")
                                except Exception as outer_e:
                                    details_text.insert(tk.END, f"Statistical analysis failed: {outer_e}\n")
                                    posthoc = None
                                # Store ANOVA results for this category for later display
                                anova_results[g] = {}
                                
                                if aov is not None and not aov.empty:
                                    # Extract p-value from ANOVA result
                                    anova_p = None
                                    df_between = aov['DF'].iloc[0] if 'DF' in aov.columns else 'N/A'
                                    df_error = aov['DF'].iloc[-1] if 'DF' in aov.columns and len(aov) > 1 else 'N/A'
                                    
                                    if 'p-unc' in aov.columns:
                                        anova_p = aov['p-unc'].iloc[0]
                                    elif 'p' in aov.columns:
                                        anova_p = aov['p'].iloc[0]
                                    
                                    # Calculate effect size (eta squared) if possible
                                    ss_between = aov['SS'].iloc[0] if 'SS' in aov.columns else 'N/A'
                                    ss_total = aov['SS'].sum() if 'SS' in aov.columns else 'N/A'
                                    eta_sq = ss_between / ss_total if isinstance(ss_between, (int, float)) and ss_total != 0 else 'N/A'
                                    
                                    # Store ANOVA results for display later
                                    f_stat = aov['F'].iloc[0] if 'F' in aov.columns else 'N/A'
                                    anova_results[g]['f_stat'] = f_stat
                                    anova_results[g]['df_between'] = df_between
                                    anova_results[g]['df_error'] = df_error
                                    anova_results[g]['p_value'] = anova_p
                                    anova_results[g]['eta_sq'] = eta_sq
                                    anova_results[g]['sig'] = self.pval_to_annotation(anova_p, alpha=alpha) if anova_p is not None else ''
                                else:
                                    # If ANOVA calculation failed, provide default values
                                    anova_results[g]['f_stat'] = 'N/A'
                                    anova_results[g]['df_between'] = 'N/A'
                                    anova_results[g]['df_error'] = 'N/A'
                                    # Use the p-values we already have from the post-hoc tests
                                    key_pattern = f"({g}, "
                                    relevant_pvals = [v for k, v in self.latest_pvals.items() 
                                                    if isinstance(k, tuple) and str(k).startswith(key_pattern)]
                                    
                                    if relevant_pvals:
                                        # Use the minimum p-value from post-hoc tests as a conservative estimate
                                        anova_p = min(relevant_pvals)
                                        anova_results[g]['p_value'] = anova_p
                                        anova_results[g]['sig'] = self.pval_to_annotation(anova_p, alpha=alpha)
                                    else:
                                        anova_results[g]['p_value'] = 'N/A'
                                        anova_results[g]['sig'] = ''
                                        
                                    anova_results[g]['eta_sq'] = 'N/A'
                                
                                # Store post-hoc test results for display later
                                
                                # Prepare p-value matrix with clean formatting
                                pval_matrix = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                
                                # Fill matrix with p-values
                                for h1 in hue_groups:
                                    for h2 in hue_groups:
                                        if h1 == h2:
                                            pval_matrix.loc[h1, h2] = '—'  # Diagonal - same group
                                        elif posthoc is not None:
                                            try:
                                                # Access p-value from posthoc result
                                                # First check if both indices exist in the posthoc result
                                                if h2 in posthoc.columns and h1 in posthoc.index:
                                                    pval_val = posthoc.loc[h1, h2]
                                                elif h1 in posthoc.columns and h2 in posthoc.index:
                                                    pval_val = posthoc.loc[h2, h1]
                                                else:
                                                    pval_val = float('nan')
                                                
                                                # Format as simple p-value with significance indicator
                                                if not np.isnan(pval_val):
                                                    sig = self.pval_to_annotation(pval_val, alpha=alpha)
                                                    pval_matrix.loc[h1, h2] = f"{pval_val:.4g} {sig}"
                                                else:
                                                    pval_matrix.loc[h1, h2] = 'n/a'
                                            except Exception as e:
                                                pval_matrix.loc[h1, h2] = 'error'
                                        else:
                                            pval_matrix.loc[h1, h2] = ''
                                
                                # Store the p-value matrix for later display
                                all_pval_matrices[g] = pval_matrix.copy()
                                
                            # End of loop for each x-axis category
                        
                        # Now display the results in a more concise format
                        # 1. First show the ANOVA results summary for each category
                        details_text.insert(tk.END, "\n" + "="*80 + "\n")
                        details_text.insert(tk.END, f"ANOVA RESULTS SUMMARY: {anova_type}\n")
                        details_text.insert(tk.END, f"Alpha level: {alpha}\n")
                        details_text.insert(tk.END, "="*80 + "\n\n")
                        
                        details_text.insert(tk.END, f"{'Category':<15} {'F-statistic':<15} {'DF':<10} {'p-value':<15} {'Sig.':<5} {'η²':<10}\n")
                        details_text.insert(tk.END, "-"*70 + "\n")
                        
                        for g, results in anova_results.items():
                            f_stat = results['f_stat']
                            df = f"{results['df_between']}, {results['df_error']}"
                            p_val = results['p_value']
                            sig = results['sig']
                            eta_sq = results['eta_sq']
                            
                            # Format the F-statistic
                            if isinstance(f_stat, (int, float)):
                                f_str = f"{f_stat:.3f}"
                            else:
                                f_str = str(f_stat)
                            
                            # Format the p-value
                            if isinstance(p_val, (int, float)):
                                p_str = f"{p_val:.4g}" if p_val < 0.001 else f"{p_val:.4f}"
                            else:
                                p_str = str(p_val)
                                
                            # Format the eta-squared value
                            if isinstance(eta_sq, (int, float)) and eta_sq != 'N/A':
                                eta_str = f"{float(eta_sq):.3f}"
                            else:
                                eta_str = str(eta_sq)
                            
                            details_text.insert(tk.END, f"{g:<15} {f_str:<15} {df:<10} {p_str:<15} {sig:<5} {eta_str:<10}\n")
                        
                        details_text.insert(tk.END, "\n\n")
                        
                        # 2. Now display the post-hoc test results for each category
                        details_text.insert(tk.END, "="*80 + "\n")
                        details_text.insert(tk.END, f"POST-HOC TEST RESULTS: {posthoc_type}\n")
                        details_text.insert(tk.END, "="*80 + "\n\n")
                        
                        # Display post-hoc matrices for each category
                        for g, matrix in all_pval_matrices.items():
                            details_text.insert(tk.END, f"Category: {g}\n")
                            details_text.insert(tk.END, "-"*40 + "\n")
                            
                            # Check if we have ANOVA results for this group
                            if g in anova_results and anova_results[g]['p_value'] != 'N/A':
                                anova_p = anova_results[g]['p_value']
                                if isinstance(anova_p, (int, float)):
                                    anova_sig = self.pval_to_annotation(anova_p, alpha=alpha)
                                    p_str = f"{anova_p:.4g}" if anova_p < 0.001 else f"{anova_p:.4f}"
                                    details_text.insert(tk.END, f"ANOVA p-value: {p_str} {anova_sig}\n")
                                    
                                    # If ANOVA is not significant, add a warning
                                    if anova_p > alpha:
                                        details_text.insert(tk.END, "NOTE: ANOVA is not significant, interpret post-hoc tests with caution\n")
                            
                            # Display the matrix with formatted p-values
                            matrix_str = self.format_pvalue_matrix(matrix)
                            details_text.insert(tk.END, matrix_str + "\n\n")
                        
                        # 3. Only show descriptive statistics section if not 2 groups case (which already displays stats)
                        if n_hue != 2:
                            details_text.insert(tk.END, "="*80 + "\n")
                            details_text.insert(tk.END, "DESCRIPTIVE STATISTICS\n")
                            details_text.insert(tk.END, "="*80 + "\n\n")
                            
                            # Determine if we should show SEM or SD based on user setting
                            error_type = self.errorbar_type_var.get().lower()
                            error_label = "SEM" if error_type == "sem" else "SD"
                            
                            # Create a dictionary to store descriptive statistics for each category and group
                            desc_stats = {}
                            
                            for g in base_groups:
                                desc_stats[g] = {}
                                df_sub = df_plot[df_plot[x_col] == g]
                                
                                for group in hue_groups:
                                    group_data = df_sub[df_sub[group_col] == group][val_col].dropna()
                                    if len(group_data) > 0:
                                        mean = group_data.mean()
                                        if error_type == "sem":
                                            error = group_data.sem()
                                        else:  # SD
                                            error = group_data.std(ddof=1)
                                        n = len(group_data)
                                        
                                        desc_stats[g][group] = {
                                            'mean': mean,
                                            'error': error,
                                            'n': n
                                        }
                            
                            # Display descriptive statistics in a table format
                            for g in base_groups:
                                details_text.insert(tk.END, f"Category: {g}\n")
                                details_text.insert(tk.END, "-"*40 + "\n")
                                details_text.insert(tk.END, f"{'Group':<15} {'Mean':<10} {error_label + ' ±':<10} {'n':<5}\n")
                                details_text.insert(tk.END, "-"*40 + "\n")
                                
                                for group in hue_groups:
                                    if g in desc_stats and group in desc_stats[g]:
                                        stats = desc_stats[g][group]
                                        details_text.insert(tk.END, f"{group:<15} {stats['mean']:<10.4g} {stats['error']:<10.4g} {stats['n']:<5}\n")
                                    else:
                                        details_text.insert(tk.END, f"{group:<15} {'No data':<10} {'-':<10} {'-':<5}\n")
                            
                            # No additional newline needed here
                            
                        # Add a section to display overall test information
                        details_text.insert(tk.END, "\n" + "="*80 + "\n")
                        details_text.insert(tk.END, "NOTES:\n")
                        details_text.insert(tk.END, "="*80 + "\n\n")
                        
                        # Display appropriate test information based on number of groups
                        if n_hue == 2:
                            details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) for multiple groups\n")
                        else:
                            details_text.insert(tk.END, f"Test Used: {anova_type} with {posthoc_type} post-hoc tests\n")
                            
                        details_text.insert(tk.END, f"Alpha level: {alpha}\n\n")
                        if any([results.get('p_value', 1.0) > alpha for results in anova_results.values()]):
                            details_text.insert(tk.END, "\n⚠️ WARNING: Some ANOVA p-values are not significant. ")
                            details_text.insert(tk.END, "When the main ANOVA is not significant, post-hoc test results should be interpreted with caution.\n")
                            
                        # Handle case for non-ANOVA tests or two groups
                        else:
                            # Skip if we already displayed the special section for exactly 2 groups
                            if n_hue == 2:
                                # We've already shown detailed results in the t-test section above
                                pass
                            else:
                                # For other cases or if posthoc not available, show the pairwise list
                                for h1, h2 in pairs:
                                    # Convert data to numeric, handling potential errors
                                    def convert_to_numeric(series):
                                        try:
                                            return pd.to_numeric(series, errors='coerce').dropna()
                                        except Exception as e:
                                            details_text.insert(tk.END, f"Error converting data to numeric: {e}\n")
                                            return pd.Series(dtype=float)
                                    
                                    vals1 = convert_to_numeric(df_sub[df_sub[group_col] == h1][val_col])
                                    vals2 = convert_to_numeric(df_sub[df_sub[group_col] == h2][val_col])
                                    
                                    # Check if we have enough valid numeric data
                                    if len(vals1) == 0 or len(vals2) == 0:
                                        details_text.insert(tk.END, f"Insufficient numeric data for t-test between {h1} and {h2}\n")
                                        continue
                                    # Get the selected t-test type and alternative
                                    ttest_type = self.ttest_type_var.get()
                                    alternative = self.ttest_alternative_var.get()
                                    
                                    try:
                                        # Create a temporary DataFrame for the t-test
                                        temp_df = pd.DataFrame({
                                            'group': [h1] * len(vals1) + [h2] * len(vals2),
                                            'value': np.concatenate([vals1, vals2])
                                        })
                                        
                                        # Calculate statistics
                                        means = [vals1.mean(), vals2.mean()]
                                        stds = [vals1.std(ddof=1), vals1.std(ddof=1)]
                                        ns = [len(vals1), len(vals2)]
                                        sems = [std / np.sqrt(n) for std, n in zip(stds, ns)]
                                        
                                        # Get error type from settings (SD or SEM)
                                        error_type = self.errorbar_type_var.get().lower()
                                        errors = stds if error_type == 'sd' else sems
                                        error_label = 'SD' if error_type == 'sd' else 'SEM'
                                        
                                        # Perform the appropriate t-test using only explot_stats
                                        try:
                                            from explot_stats import run_ttest
                                            p_val, ttest_result = run_ttest(
                                                temp_df, 'value', h1, h2, 'group',
                                                test_type=ttest_type, alternative=alternative
                                            )
                                            
                                            # Extract test statistic and degrees of freedom for display
                                            if hasattr(ttest_result, 'statistic'):
                                                test_stat = ttest_result.statistic
                                            else:
                                                test_stat = float('nan')
                                            
                                            # Calculate degrees of freedom based on test type
                                            if ttest_type == "Paired t-test" and len(vals1) == len(vals2):
                                                df = len(vals1) - 1  # n-1 for paired t-test
                                            elif ttest_type == "Student's t-test (unpaired, equal variances)":
                                                df = len(vals1) + len(vals2) - 2  # n1 + n2 - 2 for Student's t-test
                                            else:  # Welch's t-test (default)
                                                # Welch-Satterthwaite equation for degrees of freedom
                                                var1, var2 = np.var(vals1, ddof=1), np.var(vals2, ddof=1)
                                                n1, n2 = len(vals1), len(vals2)
                                                df = ((var1/n1 + var2/n2)**2) / ((var1/n1)**2/(n1-1) + (var2/n2)**2/(n2-1))
                                            
                                            # Store the p-value for this comparison
                                            key = (g, h1, h2) if (g, h1, h2) not in self.latest_pvals else (g, h2, h1)
                                            self.latest_pvals[key] = p_val
                                            
                                        except ImportError:
                                            details_text.insert(tk.END, "Error: The explot_stats module is required for t-tests.\n")
                                            details_text.insert(tk.END, "Please install the module or contact the developer.\n")
                                            raise
                                            
                                    except Exception as e:
                                        details_text.insert(tk.END, f"Error in t-test: {str(e)}\n")
                                        continue
                                        
                                    # Format test statistics
                                    test_name = ttest_type.split('(')[0].strip()
                                    p_annotation = self.pval_to_annotation(p_val)
                                    
                                    # First line: Test statistics (compact format)
                                    if 'test_stat' in locals() and 'df' in locals():
                                        test_name = "Welch's t-test" if 'Welch' in ttest_type else 't-test'
                                        details_text.insert(tk.END, 
                                            f"{g}: {h1} vs {h2} - {test_name}: "
                                            f"p = {p_val:.2g} {p_annotation}, "
                                            f"t = {test_stat:.2f}, "
                                            f"df = {df:.1f}\n"
                                        )
                                    
                                    # Second table: Means and errors (compact format)
                                    details_text.insert(tk.END, f"{'Group':<8} {'Mean':<10} {error_label+' ±':<10} {'n':<4}\n")
                                    details_text.insert(tk.END, f"{'-' * 35}\n")
                                    for group, mean, error, n in zip([h1, h2], means, errors, ns):
                                        details_text.insert(tk.END, f"{group:<8} {mean:<10.4g} {error:<10.4g} {n:<4}\n")
                                    details_text.insert(tk.END, "\n")
        except Exception as e:
            details_text.insert(tk.END, f"Error calculating statistics: {e}\n")
            details_text.insert(tk.END, traceback.format_exc())
        finally:
            details_text.config(state='disabled')
            ttk.Button(window, text='Close', command=window.destroy).pack(pady=8)


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
        
    def _load_theme_settings(self):
        """Load theme settings from JSON file."""
        if os.path.exists(self.theme_settings_file):
            try:
                with open(self.theme_settings_file, 'r') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Could not load theme settings: {e}")
        return {}

    def _save_theme_settings(self, theme_name, dark_mode):
        """Save theme settings to JSON file."""
        try:
            with open(self.theme_settings_file, 'w') as f:
                json.dump({
                    'theme_name': theme_name,
                    'dark_mode': dark_mode,
                    'timestamp': time.time()
                }, f, indent=2)
        except IOError as e:
            print(f"Warning: Could not save theme settings: {e}")

    def load_user_preferences(self):
        """Load user preferences from JSON files and apply them to the application."""
        # Load theme settings first
        theme_settings = self._load_theme_settings()
        
        # Initialize default preferences
        default_preferences = {
            # Initialize theme settings from saved file or use defaults
            'theme_name': theme_settings.get('theme_name', 'light'),
            'dark_mode': theme_settings.get('dark_mode', False),
            'plot_kind': 'bar',
            'show_stripplot': True,
            'strip_black': True,
            'errorbar_type': 'SD',
            'errorbar_black': True,
            'errorbar_capsize': 'Default',
            'bar_outline': False,
            'outline_color': 'as_set',  # Default outline color setting
            'upward_errorbar': True,  # Use upward-only error bars by default
            'bar_gap_multiplier': 0.75,  # Default bar gap multiplier
            'use_stats': False,  # Default 'Use statistics' setting
            'ttest_type': "Welch's t-test (unpaired, unequal variances)",
            'ttest_alternative': 'two-sided',
            'anova_type': "Welch's ANOVA",
            'alpha_level': "0.05",
            'posthoc_type': "Tamhane's T2",
            'linewidth': 1.0,
            'plot_width': 1.5,
            'plot_height': 1.5,
            'xy_marker_symbol': 'o',
            'xy_marker_size': 5,
            'xy_filled': True,
            'xy_line_style': 'solid',
            'xy_line_black': False,
            'xy_connect': False,
            'xy_show_mean': True,
            'xy_show_mean_errorbars': True,
            'xy_draw_band': False,
            # Default color settings
            'single_color': None,  # Will be set after colors are loaded
            'palette': None,  # Will be set after palettes are loaded
            # Theme settings
            'theme_name': 'light',
            'dark_mode': False,
            # Preview settings
            'preview_dpi': 175,  # Default DPI for preview images
            'start_maximized': True,
            'ybreak_marker_style': "Connected",
            'ybreak_marker_style_user_set': False,
            'ybreak_marker_style_user_set_schema': 1,
            'ybreak_gap': 0.07
        }
        
        # Set default colors after loading custom colors
        if hasattr(self, 'custom_colors') and self.custom_colors and default_preferences['single_color'] is None:
            default_preferences['single_color'] = list(self.custom_colors.keys())[0]
        if hasattr(self, 'custom_palettes') and self.custom_palettes and default_preferences['palette'] is None:
            default_preferences['palette'] = list(self.custom_palettes.keys())[0]
        
        # Check if user preferences file exists
        if os.path.exists(self.default_settings_file):
            try:
                with open(self.default_settings_file, 'r') as f:
                    user_prefs = json.load(f)
                # Update default_preferences with user's saved preferences
                default_preferences.update(user_prefs)
                
                # Validate color settings exist in current colors/palettes
                if 'single_color' in user_prefs and user_prefs['single_color'] not in self.custom_colors:
                    default_preferences['single_color'] = list(self.custom_colors.keys())[0]
                if 'palette' in user_prefs and user_prefs['palette'] not in self.custom_palettes:
                    default_preferences['palette'] = list(self.custom_palettes.keys())[0]
            except Exception as e:
                messagebox.showwarning("Error Loading Preferences", f"Could not load preferences: {str(e)}")
        
        # Apply the preferences to the UI
        self._apply_user_preferences(default_preferences)
    
    def _apply_user_preferences(self, preferences):
        """Apply the loaded preferences to UI elements."""
        # Apply theme settings first, before other UI elements
        if 'theme_name' in preferences and 'dark_mode' in preferences:
            theme_name = preferences['theme_name']
            dark_mode = preferences['dark_mode']
            self.theme_name = theme_name
            self.dark_mode = dark_mode
            
            # Apply the theme if the style object is available
            if hasattr(self, 'style'):
                try:
                    # For custom themes (nord, nordic), use _change_theme to properly set them up
                    if theme_name.lower() in ['nord', 'nordic']:
                        from launch import _change_theme
                        _change_theme(self.style, theme_name, dark_mode, self, update_menu=False, silent=True)
                    else:
                        # For standard ttkbootstrap themes
                        self.style.theme_use(theme_name)
                    
                    # Update the theme-dependent styles
                    self.update_theme_dependent_styles()
                    
                except Exception as e:
                    print(f"Warning: Could not apply theme {theme_name}: {e}")
                    import traceback
                    traceback.print_exc()
        
        # General tab settings
        if hasattr(self, 'plot_kind_var') and 'plot_kind' in preferences:
            self.plot_kind_var.set(preferences['plot_kind'])
        
        # Plot Settings
        if hasattr(self, 'show_stripplot_var') and 'show_stripplot' in preferences:
            self.show_stripplot_var.set(preferences['show_stripplot'])
        if hasattr(self, 'strip_black_var') and 'strip_black' in preferences:
            self.strip_black_var.set(preferences['strip_black'])
        if hasattr(self, 'errorbar_type_var') and 'errorbar_type' in preferences:
            self.errorbar_type_var.set(preferences['errorbar_type'])
        if hasattr(self, 'errorbar_black_var') and 'errorbar_black' in preferences:
            self.errorbar_black_var.set(preferences['errorbar_black'])
        if hasattr(self, 'errorbar_capsize_var') and 'errorbar_capsize' in preferences:
            self.errorbar_capsize_var.set(preferences['errorbar_capsize'])
        if hasattr(self, 'bar_outline_var') and 'bar_outline' in preferences:
            self.bar_outline_var.set(preferences['bar_outline'])
        if hasattr(self, 'outline_color_var') and 'outline_color' in preferences:
            self.outline_color_var.set(preferences['outline_color'])
        if hasattr(self, 'upward_errorbar_var') and 'upward_errorbar' in preferences:
            self.upward_errorbar_var.set(preferences['upward_errorbar'])
        if hasattr(self, 'bar_gap_multiplier_var') and 'bar_gap_multiplier' in preferences:
            self.bar_gap_multiplier_var.set(preferences['bar_gap_multiplier'])
            
        # Statistics tab
        if hasattr(self, 'use_stats_var') and 'use_stats' in preferences:
            self.use_stats_var.set(preferences['use_stats'])
        if hasattr(self, 'ttest_type_var') and 'ttest_type' in preferences:
            self.ttest_type_var.set(preferences['ttest_type'])
        if hasattr(self, 'ttest_alternative_var') and 'ttest_alternative' in preferences:
            self.ttest_alternative_var.set(preferences['ttest_alternative'])
        if hasattr(self, 'anova_type_var') and 'anova_type' in preferences:
            self.anova_type_var.set(preferences['anova_type'])
        if hasattr(self, 'alpha_level_var') and 'alpha_level' in preferences:
            self.alpha_level_var.set(preferences['alpha_level'])
        if hasattr(self, 'posthoc_type_var') and 'posthoc_type' in preferences:
            self.posthoc_type_var.set(preferences['posthoc_type'])
            
        # Appearance tab
        if hasattr(self, 'linewidth') and 'linewidth' in preferences:
            self.linewidth.set(preferences['linewidth'])
        if 'plot_width' in preferences and hasattr(self, 'plot_width_var'):
            self.plot_width_var.set(preferences['plot_width'])
        if 'plot_height' in preferences and hasattr(self, 'plot_height_var'):
            self.plot_height_var.set(preferences['plot_height'])
        if hasattr(self, 'preview_dpi') and 'preview_dpi' in preferences:
            self.preview_dpi.set(preferences['preview_dpi'])

        if hasattr(self, 'start_maximized_var') and 'start_maximized' in preferences:
            self.start_maximized_var.set(preferences['start_maximized'])

        if hasattr(self, 'ybreak_marker_style_var') and 'ybreak_marker_style' in preferences:
            user_set = bool(preferences.get('ybreak_marker_style_user_set', False))
            schema = int(preferences.get('ybreak_marker_style_user_set_schema', 0) or 0)
            if schema < 1:
                user_set = False

            if not user_set:
                self.ybreak_marker_style_var.set('Connected')
                if hasattr(self, 'ybreak_marker_style_user_set_var'):
                    self.ybreak_marker_style_user_set_var.set(False)
            else:
                self.ybreak_marker_style_var.set(preferences['ybreak_marker_style'])
                if hasattr(self, 'ybreak_marker_style_user_set_var'):
                    self.ybreak_marker_style_user_set_var.set(True)

        if hasattr(self, 'ybreak_gap_entry') and 'ybreak_gap' in preferences:
            try:
                gap = float(preferences['ybreak_gap'])
                gap = max(0.0, min(0.3, gap))
                self.ybreak_gap_entry.delete(0, tk.END)
                self.ybreak_gap_entry.insert(0, f"{gap:g}")
            except Exception:
                pass
            
        # Color settings
        if 'single_color' in preferences and hasattr(self, 'single_color_var'):
            if preferences['single_color'] in self.custom_colors:
                self.single_color_var.set(preferences['single_color'])
        if 'palette' in preferences and hasattr(self, 'palette_var'):
            if preferences['palette'] in self.custom_palettes:
                self.palette_var.set(preferences['palette'])
                
        # Theme settings
        # Load theme settings
        if 'theme_name' in preferences:
            self.theme_name = preferences['theme_name']
        if 'dark_mode' in preferences:
            self.dark_mode = preferences['dark_mode']
            
        # Save theme settings to dedicated file
        if hasattr(self, 'theme_name') and hasattr(self, 'dark_mode'):
            self._save_theme_settings(self.theme_name, self.dark_mode)
        
        # XY Plot tab
        if hasattr(self, 'xy_marker_symbol_var') and 'xy_marker_symbol' in preferences:
            self.xy_marker_symbol_var.set(preferences['xy_marker_symbol'])
        if hasattr(self, 'xy_marker_size_var') and 'xy_marker_size' in preferences:
            self.xy_marker_size_var.set(preferences['xy_marker_size'])
        if hasattr(self, 'xy_filled_var') and 'xy_filled' in preferences:
            self.xy_filled_var.set(preferences['xy_filled'])
        if hasattr(self, 'xy_line_style_var') and 'xy_line_style' in preferences:
            self.xy_line_style_var.set(preferences['xy_line_style'])
        if hasattr(self, 'xy_line_black_var') and 'xy_line_black' in preferences:
            self.xy_line_black_var.set(preferences['xy_line_black'])
        if hasattr(self, 'xy_connect_var') and 'xy_connect' in preferences:
            self.xy_connect_var.set(preferences['xy_connect'])
        if hasattr(self, 'xy_show_mean_var') and 'xy_show_mean' in preferences:
            self.xy_show_mean_var.set(preferences['xy_show_mean'])
        if hasattr(self, 'xy_show_mean_errorbars_var') and 'xy_show_mean_errorbars' in preferences:
            self.xy_show_mean_errorbars_var.set(preferences['xy_show_mean_errorbars'])
        if hasattr(self, 'xy_draw_band_var') and 'xy_draw_band' in preferences:
            self.xy_draw_band_var.set(preferences['xy_draw_band'])
    
    def save_user_preferences(self, silent=False):
        """
        Save current UI settings as user preferences to a JSON file.
        
        Args:
            silent: If True, suppresses the success message (used for automatic saves)
        """
        preferences = {}
        
        # Always save current color settings (not just from the settings dialog)
        if hasattr(self, 'single_color_var') and self.single_color_var.get() in self.custom_colors:
            preferences['single_color'] = self.single_color_var.get()
        if hasattr(self, 'palette_var') and self.palette_var.get() in self.custom_palettes:
            preferences['palette'] = self.palette_var.get()
        if hasattr(self, 'outline_color_var'):
            preferences['outline_color'] = self.outline_color_var.get()
        if hasattr(self, 'bar_gap_multiplier_var'):
            preferences['bar_gap_multiplier'] = self.bar_gap_multiplier_var.get()

        if hasattr(self, 'start_maximized_var'):
            preferences['start_maximized'] = self.start_maximized_var.get()

        if hasattr(self, 'ybreak_marker_style_var'):
            preferences['ybreak_marker_style'] = self.ybreak_marker_style_var.get()
            if hasattr(self, 'ybreak_marker_style_user_set_var'):
                preferences['ybreak_marker_style_user_set'] = bool(self.ybreak_marker_style_user_set_var.get())
            else:
                preferences['ybreak_marker_style_user_set'] = False
            preferences['ybreak_marker_style_user_set_schema'] = 1

        if hasattr(self, 'ybreak_gap_entry'):
            try:
                gap_str = self.ybreak_gap_entry.get().strip()
                if gap_str:
                    if '-' in gap_str and not gap_str.startswith('-'):
                        gap_str = gap_str.split('-')[0]
                    gap = float(gap_str)
                    gap = max(0.0, min(0.3, gap))
                    preferences['ybreak_gap'] = gap
            except Exception:
                pass
            
        # Save theme settings to both preferences and dedicated file
        if hasattr(self, 'theme_name') and hasattr(self, 'dark_mode'):
            preferences['theme_name'] = self.theme_name
            preferences['dark_mode'] = self.dark_mode
            self._save_theme_settings(self.theme_name, self.dark_mode)
        
        # General tab settings
        if hasattr(self, 'settings_plot_kind_var'):
            preferences['plot_kind'] = self.settings_plot_kind_var.get()
            
        # Plot Settings
        if hasattr(self, 'settings_show_stripplot_var'):
            preferences['show_stripplot'] = self.settings_show_stripplot_var.get()
        if hasattr(self, 'settings_strip_black_var'):
            preferences['strip_black'] = self.settings_strip_black_var.get()
        if hasattr(self, 'settings_errorbar_type_var'):
            preferences['errorbar_type'] = self.settings_errorbar_type_var.get()
        if hasattr(self, 'settings_errorbar_black_var'):
            preferences['errorbar_black'] = self.settings_errorbar_black_var.get()
        if hasattr(self, 'settings_errorbar_capsize_var'):
            preferences['errorbar_capsize'] = self.settings_errorbar_capsize_var.get()
            
        # Bar Graph tab
        if hasattr(self, 'settings_bar_outline_var'):
            preferences['bar_outline'] = self.settings_bar_outline_var.get()
        if hasattr(self, 'settings_upward_errorbar_var'):
            preferences['upward_errorbar'] = self.settings_upward_errorbar_var.get()
        if hasattr(self, 'settings_bar_gap_multiplier_var'):
            preferences['bar_gap_multiplier'] = self.settings_bar_gap_multiplier_var.get()
            
        # Colors tab settings
        if hasattr(self, 'settings_single_color_var'):
            preferences['single_color'] = self.settings_single_color_var.get()
        if hasattr(self, 'settings_palette_var'):
            preferences['palette'] = self.settings_palette_var.get()
        if hasattr(self, 'settings_outline_color_var'):
            preferences['outline_color'] = self.settings_outline_color_var.get()
            
        # Statistics tab
        if hasattr(self, 'settings_use_stats_var'):
            preferences['use_stats'] = self.settings_use_stats_var.get()
        if hasattr(self, 'settings_ttest_type_var'):
            preferences['ttest_type'] = self.settings_ttest_type_var.get()
        if hasattr(self, 'settings_ttest_alternative_var'):
            preferences['ttest_alternative'] = self.settings_ttest_alternative_var.get()
        if hasattr(self, 'settings_anova_type_var'):
            preferences['anova_type'] = self.settings_anova_type_var.get()
        if hasattr(self, 'settings_alpha_level_var'):
            preferences['alpha_level'] = self.settings_alpha_level_var.get()
        if hasattr(self, 'settings_posthoc_type_var'):
            preferences['posthoc_type'] = self.settings_posthoc_type_var.get()
            
            self.xy_show_mean_errorbars_var.set(preferences.get('xy_show_mean_errorbars', True))
        if hasattr(self, 'xy_draw_band_var') and 'xy_draw_band' in preferences:
            self.xy_draw_band_var.set(preferences['xy_draw_band'])
            
        # Appearance tab
        if hasattr(self, 'settings_linewidth'):
            preferences['linewidth'] = self.settings_linewidth.get()
        if hasattr(self, 'settings_plot_width_var'):
            plot_width = self.settings_plot_width_var.get()
            preferences['plot_width'] = plot_width
            self.plot_width_var.set(plot_width)  # Apply to main app immediately
        if hasattr(self, 'settings_plot_height_var'):
            plot_height = self.settings_plot_height_var.get()
            preferences['plot_height'] = plot_height
            self.plot_height_var.set(plot_height)  # Apply to main app immediately
        if hasattr(self, 'preview_dpi'):
            preferences['preview_dpi'] = self.preview_dpi.get()
            
        # XY Plot tab
        if hasattr(self, 'settings_xy_marker_symbol_var'):
            preferences['xy_marker_symbol'] = self.settings_xy_marker_symbol_var.get()
        if hasattr(self, 'settings_xy_marker_size_var'):
            preferences['xy_marker_size'] = self.settings_xy_marker_size_var.get()
        if hasattr(self, 'settings_xy_filled_var'):
            preferences['xy_filled'] = self.settings_xy_filled_var.get()
        if hasattr(self, 'settings_xy_line_style_var'):
            preferences['xy_line_style'] = self.settings_xy_line_style_var.get()
        if hasattr(self, 'settings_xy_line_black_var'):
            preferences['xy_line_black'] = self.settings_xy_line_black_var.get()
        if hasattr(self, 'settings_xy_connect_var'):
            preferences['xy_connect'] = self.settings_xy_connect_var.get()
        if hasattr(self, 'settings_xy_show_mean_var'):
            preferences['xy_show_mean'] = self.settings_xy_show_mean_var.get()
        if hasattr(self, 'settings_xy_show_mean_errorbars_var'):
            preferences['xy_show_mean_errorbars'] = self.settings_xy_show_mean_errorbars_var.get()
        if hasattr(self, 'settings_xy_draw_band_var'):
            preferences['xy_draw_band'] = self.settings_xy_draw_band_var.get()
            
        # Save preferences to file
        try:
            # Ensure config directory exists
            os.makedirs(os.path.dirname(self.default_settings_file), exist_ok=True)
            
            # Save preferences
            with open(self.default_settings_file, 'w') as f:
                json.dump(preferences, f, indent=2)
                
            # Apply preferences to application variables
            self.plot_kind_var.set(preferences.get('plot_kind', 'bar'))
            self.show_stripplot_var.set(preferences.get('show_stripplot', True))
            self.strip_black_var.set(preferences.get('strip_black', True))
            self.errorbar_type_var.set(preferences.get('errorbar_type', 'SD'))
            self.errorbar_black_var.set(preferences.get('errorbar_black', True))
            self.errorbar_capsize_var.set(preferences.get('errorbar_capsize', 'Default'))
            self.use_stats_var.set(preferences.get('use_stats', False))
            self.ttest_type_var.set(preferences.get('ttest_type', "Welch's t-test (unpaired, unequal variances)"))
            self.ttest_alternative_var.set(preferences.get('ttest_alternative', 'two-sided'))
            self.anova_type_var.set(preferences.get('anova_type', "Welch's ANOVA"))
            self.alpha_level_var.set(preferences.get('alpha_level', "0.05"))
            self.posthoc_type_var.set(preferences.get('posthoc_type', "Tamhane's T2"))
            self.linewidth.set(preferences.get('linewidth', 1.0))
            self.plot_width_var.set(preferences.get('plot_width', 1.5))
            self.plot_height_var.set(preferences.get('plot_height', 1.5))
            self.xy_marker_symbol_var.set(preferences.get('xy_marker_symbol', 'o'))
            self.xy_marker_size_var.set(preferences.get('xy_marker_size', 5))
            self.xy_filled_var.set(preferences.get('xy_filled', True))
            self.xy_line_style_var.set(preferences.get('xy_line_style', 'solid'))
            self.xy_line_black_var.set(preferences.get('xy_line_black', False))
            self.xy_connect_var.set(preferences.get('xy_connect', False))
            self.xy_show_mean_var.set(preferences.get('xy_show_mean', True))
            self.xy_show_mean_errorbars_var.set(preferences.get('xy_show_mean_errorbars', True))
            self.xy_draw_band_var.set(preferences.get('xy_draw_band', False))
            self.bar_outline_var.set(preferences.get('bar_outline', False))
            
            # Only show success message if not in silent mode
            if not silent:
                messagebox.showinfo("Settings Saved", "Your preferences have been saved and applied.")
        except Exception as e:
            if not silent:  # Only show error dialog if not in silent mode
                messagebox.showerror("Error Saving Preferences", f"Could not save preferences: {str(e)}")
            print(f"Error saving preferences: {str(e)}")

    def setup_ui(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True)

        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side='left', fill='both', expand=False)

        self.tab_control = ttk.Notebook(left_frame)
        self.tab_control.pack(fill='both', expand=True)

        self.basic_tab = ttk.Frame(self.tab_control)
        self.appearance_tab = ttk.Frame(self.tab_control)
        self.axis_tab = ttk.Frame(self.tab_control)
        self.colors_tab = ttk.Frame(self.tab_control)
        self.stats_settings_tab = ttk.Frame(self.tab_control)
        self.xy_fitting_tab = ttk.Frame(self.tab_control)

        self.tab_control.add(self.basic_tab, text="Basic")
        self.tab_control.add(self.appearance_tab, text="Appearance")
        self.tab_control.add(self.axis_tab, text="Axis")
        self.tab_control.add(self.stats_settings_tab, text="Statistics")
        self.tab_control.add(self.xy_fitting_tab, text="XY Fitting")
        self.tab_control.add(self.colors_tab, text="Colors")

        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side='right', fill='both', expand=True)
        
        # Create a top frame for action buttons with fixed height
        top_button_frame = ttk.Frame(right_frame, padding=10)
        top_button_frame.pack(fill='x', padx=10, pady=(5, 10))
        top_button_frame.columnconfigure(0, weight=1)
        top_button_frame.columnconfigure(1, weight=1)
        
        # Create a sub-frame to hold the buttons with fixed size
        button_container = ttk.Frame(top_button_frame)
        button_container.grid(row=0, column=0, sticky='w')
        
        # Plot and Stats buttons side by side with fixed positioning
        button_style = {'style': 'Accent.TButton', 'width': 14, 'padding': (5, 4)}
        
        # First button (always visible)
        plot_button = ttk.Button(
            button_container, 
            text="Generate Plot", 
            command=self.plot_graph,
            **button_style
        )
        plot_button.grid(row=0, column=0, padx=3, pady=2, sticky='nsew')
        
        # Second button (toggled by stats checkbox)
        self.stats_details_btn = ttk.Button(
            button_container, 
            text="Stats Details", 
            command=self.show_statistical_details,
            **button_style
        )
        self.stats_details_btn.grid(row=0, column=1, padx=3, pady=2, sticky='nsew')
        
        # Store reference to the container for toggling
        self.button_container = button_container
        
        # Scrollable canvas frame for the plot
        # Outer frame holds canvas + scrollbars
        self.preview_outer_frame = ttk.Frame(right_frame)
        self.preview_outer_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        
        # Create canvas for scrolling (this is the viewport)
        self.preview_scroll_canvas = tk.Canvas(self.preview_outer_frame, highlightthickness=0, bg='white')
        
        # Create scrollbars
        self.preview_v_scrollbar = ttk.Scrollbar(self.preview_outer_frame, orient='vertical', 
                                                  command=self.preview_scroll_canvas.yview)
        self.preview_h_scrollbar = ttk.Scrollbar(self.preview_outer_frame, orient='horizontal', 
                                                  command=self.preview_scroll_canvas.xview)
        
        # Configure canvas scrolling
        self.preview_scroll_canvas.configure(yscrollcommand=self.preview_v_scrollbar.set, 
                                             xscrollcommand=self.preview_h_scrollbar.set)
        
        # Grid layout
        self.preview_scroll_canvas.grid(row=0, column=0, sticky='nsew')
        self.preview_v_scrollbar.grid(row=0, column=1, sticky='ns')
        self.preview_h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights so canvas expands
        self.preview_outer_frame.grid_rowconfigure(0, weight=1)
        self.preview_outer_frame.grid_columnconfigure(0, weight=1)
        
        # Create inner frame for plot content - this frame will NOT be resized
        self.canvas_frame = ttk.Frame(self.preview_scroll_canvas)
        
        # Create window at top-left, anchor to northwest
        self.preview_window_id = self.preview_scroll_canvas.create_window(
            (0, 0), window=self.canvas_frame, anchor='nw'
        )
        
        # Update scroll region when content changes
        def update_scrollregion(event=None):
            self.preview_scroll_canvas.configure(scrollregion=self.preview_scroll_canvas.bbox('all'))
        self.canvas_frame.bind('<Configure>', update_scrollregion)
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            self.preview_scroll_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        def on_shift_mousewheel(event):
            self.preview_scroll_canvas.xview_scroll(int(-1 * (event.delta / 120)), 'units')
        self.preview_scroll_canvas.bind_all('<MouseWheel>', on_mousewheel)
        self.preview_scroll_canvas.bind_all('<Shift-MouseWheel>', on_shift_mousewheel)
        
        # Bottom frame (kept for potential future use)
        self.bottom_frame = ttk.Frame(self.root)
        self.bottom_frame.pack(side="bottom", fill="x")


        self.setup_basic_tab()
        self.setup_appearance_tab()
        self.setup_axis_tab()
        self.setup_colors_tab()
        self.setup_xy_fitting_tab()

    def setup_basic_tab(self):
        frame = self.basic_tab
        # File/Sheet group
        file_grp = ttk.LabelFrame(frame, text="File/Sheet", padding=6)
        file_grp.pack(fill='x', padx=6, pady=(8,4))
        ttk.Button(file_grp, text='Load Excel File', command=self.load_file).pack(fill='x', pady=2)
        ttk.Label(file_grp, text="Sheet:").pack(anchor="w")
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(file_grp, textvariable=self.sheet_var, width=18)
        self.sheet_dropdown.pack(fill='x', pady=2)
        self.sheet_dropdown.bind('<<ComboboxSelected>>', self.load_sheet)
        # Columns group
        col_grp = ttk.LabelFrame(frame, text="Columns", padding=6)
        col_grp.pack(fill='x', padx=6, pady=4)
        ttk.Label(col_grp, text="X-axis column:").grid(row=0, column=0, sticky="w", pady=2)
        self.xaxis_var = tk.StringVar()
        self.xaxis_var.trace_add('write', self.update_x_axis_label)
        self.xaxis_dropdown = ttk.Combobox(col_grp, textvariable=self.xaxis_var, width=18)
        self.xaxis_dropdown.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        ttk.Label(col_grp, text="Group column:").grid(row=1, column=0, sticky="w", pady=2)
        self.group_var = tk.StringVar()
        self.group_dropdown = ttk.Combobox(col_grp, textvariable=self.group_var, width=18)
        self.group_dropdown.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        ttk.Label(col_grp, text="Y-axis columns:").grid(row=2, column=0, sticky="nw", pady=(4,0))
        # Scrollable frame for value checkbuttons
        value_vars_scroll_frame = ttk.Frame(col_grp)
        value_vars_scroll_frame.grid(row=2, column=1, sticky="ew", pady=(2,0))
        value_vars_canvas = tk.Canvas(value_vars_scroll_frame, height=150)
        value_vars_scrollbar = ttk.Scrollbar(value_vars_scroll_frame, orient="vertical", command=value_vars_canvas.yview)
        self.value_vars_inner_frame = ttk.Frame(value_vars_canvas)
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
        opt_grp = ttk.LabelFrame(frame, text="Options", padding=6)
        opt_grp.pack(fill='x', padx=6, pady=4)
        
        # Statistics settings
        stats_frame = ttk.Frame(opt_grp)
        stats_frame.pack(anchor="w", pady=2)
        ttk.Label(stats_frame, text="Statistics:").pack(side="left")
        ttk.Checkbutton(stats_frame, text="Use statistics", variable=self.use_stats_var).pack(side="left")
        
        # Add checkbox for showing statistical annotations
        self.show_statistics_annotations_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_grp, text="Show statistical annotations on plot", variable=self.show_statistics_annotations_var).pack(anchor="w", pady=1)
        
        # --- Error bar type (SD/SEM) ---
        # We use the errorbar_type_var that was initialized in __init__
        # No need to create it again here
            
        errorbar_frame = ttk.Frame(opt_grp)
        errorbar_frame.pack(anchor="w", pady=2)
        ttk.Label(errorbar_frame, text="Error bars:").pack(side="left")
        ttk.Radiobutton(errorbar_frame, text="SD", variable=self.errorbar_type_var, value="SD").pack(side="left")
        ttk.Radiobutton(errorbar_frame, text="SEM", variable=self.errorbar_type_var, value="SEM").pack(side="left")
        # Black errorbars option
        self.errorbar_black_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt_grp, text="Black errorbars", variable=self.errorbar_black_var).pack(anchor="w", pady=1)
        btn_fr = ttk.Frame(opt_grp)
        btn_fr.pack(fill='x', pady=(2,0))
        ttk.Button(btn_fr, text="Modify X categories", command=self.modify_x_categories, width=20).pack(side="left", padx=1)
        
        # Plot Type group
        type_grp = ttk.LabelFrame(frame, text="Plot Type", padding=6)
        type_grp.pack(fill='x', padx=6, pady=4)
        # Frame for radio buttons (left)
        plot_type_radio_frame = ttk.Frame(type_grp)
        plot_type_radio_frame.grid(row=0, column=0, sticky="nw")
        ttk.Label(plot_type_radio_frame, text="Plot Type:").pack(anchor="w")
        bar_radio = ttk.Radiobutton(plot_type_radio_frame, text="Bar Graph", variable=self.plot_kind_var, value="bar")
        box_radio = ttk.Radiobutton(plot_type_radio_frame, text="Box Plot", variable=self.plot_kind_var, value="box")
        violin_radio = ttk.Radiobutton(plot_type_radio_frame, text="Violin Plot", variable=self.plot_kind_var, value="violin")
        xy_radio = ttk.Radiobutton(plot_type_radio_frame, text="XY Plot", variable=self.plot_kind_var, value="xy")
        bar_radio.pack(anchor="w")
        box_radio.pack(anchor="w")
        violin_radio.pack(anchor="w")
        xy_radio.pack(anchor="w")
        # XY options frame (right)
        self.xy_options_frame = ttk.Frame(type_grp)
        # XY options widgets in xy_options_frame
        self.xy_marker_symbol_label = ttk.Label(self.xy_options_frame, text="XY Marker Symbol:")
        self.xy_marker_symbol_dropdown = ttk.Combobox(self.xy_options_frame, textvariable=self.xy_marker_symbol_var, values=["o", "s", "^", "D", "v", "P", "X", "+", "x", "*", "."], width=10)
        self.xy_marker_size_label = ttk.Label(self.xy_options_frame, text="XY Marker Size:")
        self.xy_marker_size_entry = ttk.Entry(self.xy_options_frame, textvariable=self.xy_marker_size_var, width=6)
        self.xy_filled_check = ttk.Checkbutton(self.xy_options_frame, text="Filled symbols", variable=self.xy_filled_var)
        self.xy_line_style_label = ttk.Label(self.xy_options_frame, text="Line style:")
        self.xy_line_style_dropdown = ttk.Combobox(self.xy_options_frame, textvariable=self.xy_line_style_var, values=["solid", "dashed", "dotted", "dashdot"], width=10)
        self.xy_line_black_check = ttk.Checkbutton(self.xy_options_frame, text="Lines in black", variable=self.xy_line_black_var)
        self.xy_connect_check = ttk.Checkbutton(self.xy_options_frame, text="Connect mean with lines", variable=self.xy_connect_var)
        self.xy_show_mean_check = ttk.Checkbutton(self.xy_options_frame, text="Show mean values", variable=self.xy_show_mean_var, command=self.update_xy_mean_errorbar_state)
        self.xy_show_mean_errorbars_check = ttk.Checkbutton(self.xy_options_frame, text="With errorbars", variable=self.xy_show_mean_errorbars_var)
        self.xy_draw_band_check = ttk.Checkbutton(self.xy_options_frame, text="Draw bands (min-max or error)", variable=self.xy_draw_band_var)
        # Pack XY widgets in order
        self.xy_marker_symbol_label.grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.xy_marker_symbol_dropdown.grid(row=0, column=1, sticky="w", padx=2, pady=2)
        self.xy_marker_size_label.grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.xy_marker_size_entry.grid(row=1, column=1, sticky="w", padx=2, pady=2)
        self.xy_connect_check.grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_show_mean_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_show_mean_errorbars_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=24, pady=2)
        self.xy_draw_band_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=4, pady=2)
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
        size_grp = ttk.LabelFrame(frame, text="Figure Size", padding=6)
        size_grp.pack(fill='x', padx=6, pady=(8,4))
        ttk.Label(size_grp, text="Plot Width (inches):").grid(row=0, column=0, sticky="w", pady=2)
        self.width_entry = ttk.Entry(size_grp, textvariable=self.plot_width_var)
        self.width_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        ttk.Label(size_grp, text="Plot Height per plot (inches):").grid(row=1, column=0, sticky="w", pady=2)
        self.height_entry = ttk.Entry(size_grp, textvariable=self.plot_height_var)
        self.height_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        size_grp.columnconfigure(1, weight=1)
        # --- Font/Line group ---
        font_grp = ttk.LabelFrame(frame, text="Font & Line", padding=6)
        font_grp.pack(fill='x', padx=6, pady=4)
        ttk.Label(font_grp, text="Font Size:").grid(row=0, column=0, sticky="w", pady=2)
        self.fontsize_entry = ttk.Entry(font_grp)
        self.fontsize_entry.insert(0, "10")
        self.fontsize_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        ttk.Label(font_grp, text="Line Width:").grid(row=1, column=0, sticky="w", pady=2)
        self.linewidth_entry = ttk.Entry(font_grp, textvariable=self.linewidth)
        self.linewidth_entry.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        ttk.Label(font_grp, text="Error Bar Capsize:").grid(row=2, column=0, sticky="w", pady=2)
        self.capsize_dropdown = ttk.Combobox(font_grp, textvariable=self.errorbar_capsize_var, 
                                        values=["Default", "Narrow", "Wide", "Wider", "None"], width=10)
        self.capsize_dropdown.grid(row=2, column=1, sticky="ew", padx=2, pady=2)
        font_grp.columnconfigure(1, weight=1)
        # --- Bar Graph group ---
        bar_grp = ttk.LabelFrame(frame, text="Bar Graph", padding=6)
        bar_grp.pack(fill='x', padx=6, pady=4)
        ttk.Checkbutton(bar_grp, text="Draw bar outlines", variable=self.bar_outline_var).pack(anchor="w", pady=1)
        ttk.Checkbutton(bar_grp, text="Upward-only error bars", variable=self.upward_errorbar_var).pack(anchor="w", pady=1)
        
        # Bar gap multiplier control
        bar_gap_frame = ttk.Frame(bar_grp)
        bar_gap_frame.pack(anchor="w", pady=1, fill='x')
        ttk.Label(bar_gap_frame, text="Bar gap multiplier:").pack(side="left")
        tk.Spinbox(bar_gap_frame, from_=0.5, to=1.0, increment=0.05, 
                  textvariable=self.bar_gap_multiplier_var, width=5).pack(side="left", padx=4)
        
        # --- Violin Plot group ---
        violin_grp = ttk.LabelFrame(frame, text="Violin Plot", padding=6)
        violin_grp.pack(fill='x', padx=6, pady=4)
        ttk.Checkbutton(violin_grp, text="Show box inside violin", variable=self.violin_inner_box_var).pack(anchor="w", pady=1)
        
        # --- XY Plot group ---
        xy_grp = ttk.LabelFrame(frame, text="XY Plot", padding=6)
        xy_grp.pack(fill='x', padx=6, pady=4)
        ttk.Checkbutton(xy_grp, text="Filled symbols", variable=self.xy_filled_var).pack(anchor="w", pady=1)
        
        # Line style
        line_style_frame = ttk.Frame(xy_grp)
        line_style_frame.pack(anchor="w", pady=1, fill='x')
        ttk.Label(line_style_frame, text="Line style:").pack(side="left")
        ttk.Combobox(line_style_frame, textvariable=self.xy_line_style_var, 
                    values=["solid", "dashed", "dotted", "dashdot"], width=10).pack(side="left", padx=4)
        
        ttk.Checkbutton(xy_grp, text="Lines in black", variable=self.xy_line_black_var).pack(anchor="w", pady=1)
        
        # --- Stripplot group ---
        strip_grp = ttk.LabelFrame(frame, text="Stripplot", padding=6)
        strip_grp.pack(fill='x', padx=6, pady=4)
        self.strip_black_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(strip_grp, text="Show stripplot with black dots", variable=self.strip_black_var).pack(anchor="w", pady=1)
        self.show_stripplot_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(strip_grp, text="Show stripplot", variable=self.show_stripplot_var).pack(anchor="w", pady=1)
        # --- Display Options and Legend in same row ---
        display_legend_row = ttk.Frame(frame)
        display_legend_row.pack(fill='x', padx=6, pady=4)
        
        # Display options group (left side)
        disp_grp = ttk.LabelFrame(display_legend_row, text="Display Options", padding=6)
        disp_grp.pack(side='left', fill='both', expand=True, padx=(0,3))
        self.show_frame_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(disp_grp, text="Show graph frame", variable=self.show_frame_var).pack(anchor="w", pady=1)
        self.show_hgrid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(disp_grp, text="Show horizontal grid", variable=self.show_hgrid_var).pack(anchor="w", pady=1)
        self.show_vgrid_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(disp_grp, text="Show vertical grid", variable=self.show_vgrid_var).pack(anchor="w", pady=1)
        
        # Preview DPI setting
        dpi_frame = ttk.Frame(disp_grp)
        dpi_frame.pack(anchor="w", pady=1, fill='x')
        ttk.Label(dpi_frame, text="Preview DPI:").pack(side="left")
        # DPI controls pixel size of preview (higher = larger preview, same layout)
        tk.Spinbox(dpi_frame, from_=50, to=300, increment=25, 
                  textvariable=self.preview_dpi, width=5).pack(side="left", padx=4)
        ttk.Label(dpi_frame, text="dpi").pack(side="left")
        
        # Legend settings group (right side)
        legend_grp = ttk.LabelFrame(display_legend_row, text="Legend", padding=6)
        legend_grp.pack(side='left', fill='both', expand=True, padx=(3,0))
        
        # Show/hide legend
        ttk.Checkbutton(legend_grp, text="Show legend", variable=self.legend_visible_var).pack(anchor="w", pady=1)
        
        # Legend position
        pos_frame = ttk.Frame(legend_grp)
        pos_frame.pack(anchor="w", pady=1, fill='x')
        ttk.Label(pos_frame, text="Position:").pack(side="left")
        legend_positions = ["best", "upper right", "upper left", "lower left", "lower right", 
                           "right", "center left", "center right", "lower center", "upper center", 
                           "center", "outside right", "outside top"]
        ttk.Combobox(pos_frame, textvariable=self.legend_position_var, 
                    values=legend_positions, width=12, state="readonly").pack(side="left", padx=4)
        
        # Legend columns
        col_frame = ttk.Frame(legend_grp)
        col_frame.pack(anchor="w", pady=1, fill='x')
        ttk.Label(col_frame, text="Columns:").pack(side="left")
        tk.Spinbox(col_frame, from_=0, to=10, increment=1, 
                  textvariable=self.legend_ncol_var, width=5).pack(side="left", padx=4)
        ttk.Label(col_frame, text="(0 = auto)").pack(side="left")
        
        # Note: Swap axes setting moved to Axis tab

    def setup_axis_tab(self):
        frame = self.axis_tab
        
        # --- Combined X/Y Labels frame ---
        labels_frame = ttk.Frame(frame)
        labels_frame.pack(fill='x', padx=4, pady=1)
        
        # X-axis label row with grid layout for alignment
        xlabel_frame = ttk.Frame(labels_frame)
        xlabel_frame.pack(fill='x', padx=0, pady=1)
        xlabel_frame.columnconfigure(1, weight=1)
        ttk.Label(xlabel_frame, text="X-axis Label:", width=14, anchor="w").grid(row=0, column=0, padx=2)
        self.xlabel_entry = ttk.Entry(xlabel_frame)
        self.xlabel_entry.grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(xlabel_frame, text="Format", command=lambda: self.open_label_formatter('x')).grid(row=0, column=2, padx=2)
        
        # Y-axis label row (just below X)
        ylabel_frame = tk.Frame(labels_frame)
        ylabel_frame.pack(fill='x', padx=0, pady=1)
        ylabel_frame.columnconfigure(1, weight=1)
        ttk.Label(ylabel_frame, text="Y-axis Label:", width=14, anchor="w").grid(row=0, column=0, padx=2)
        self.ylabel_entry = ttk.Entry(ylabel_frame)
        self.ylabel_entry.grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Button(ylabel_frame, text="Format", command=lambda: self.open_label_formatter('y')).grid(row=0, column=2, padx=2)
        
        # --- Label orientation (horizontal layout) ---
        orient_frame = tk.Frame(frame)
        orient_frame.pack(fill='x', padx=4, pady=1)
        ttk.Label(orient_frame, text="X-axis Label Orientation:").pack(side="left")
        self.label_orientation = tk.StringVar(value="vertical")
        ttk.Radiobutton(orient_frame, text="Vertical", variable=self.label_orientation, 
                      value="vertical").pack(side="left", padx=5)
        ttk.Radiobutton(orient_frame, text="Horizontal", variable=self.label_orientation, 
                       value="horizontal").pack(side="left", padx=5)
        ttk.Radiobutton(orient_frame, text="Angled", variable=self.label_orientation, 
                      value="angled").pack(side="left", padx=5)
        
        # --- Axis swap option ---
        swap_frame = tk.Frame(frame)
        swap_frame.pack(fill='x', padx=4, pady=(10, 1))
        self.swap_axes_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(swap_frame, text="Swap X and Y axes", variable=self.swap_axes_var).pack(anchor="w", pady=2)
        
        # --- X-Axis settings (more compact with grid layout) ---
        xaxis_grp = ttk.LabelFrame(frame, text="X-Axis Settings", padding=4)
        xaxis_grp.pack(fill='x', padx=4, pady=1)
        
        # Use grid layout for better alignment
        xaxis_grid = tk.Frame(xaxis_grp)
        xaxis_grid.pack(fill='x', padx=2, pady=1)
        
        # Configure columns for alignment - fixed label widths
        xaxis_grid.columnconfigure(1, weight=0) # Min Entry
        xaxis_grid.columnconfigure(3, weight=0) # Max Entry
        
        # Row 1: Min/Max values
        tk.Label(xaxis_grid, text="Minimum:", width=12, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.xmin_entry = ttk.Entry(xaxis_grid, width=10)
        self.xmin_entry.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Maximum:", width=12, anchor="w").grid(row=0, column=2, sticky="w", padx=(10,0), pady=2)
        self.xmax_entry = ttk.Entry(xaxis_grid, width=10)
        self.xmax_entry.grid(row=0, column=3, sticky="w", pady=2)
        
        # Row 2: Tick settings
        tk.Label(xaxis_grid, text="Major Interval:", width=12, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.xinterval_entry = ttk.Entry(xaxis_grid, width=10)
        self.xinterval_entry.grid(row=1, column=1, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Minor/Major:", width=12, anchor="w").grid(row=1, column=2, sticky="w", padx=(10,0), pady=2)
        self.xminor_ticks_entry = ttk.Entry(xaxis_grid, width=10)
        self.xminor_ticks_entry.grid(row=1, column=3, sticky="w", pady=2)
        
        # Row 3: Log options
        self.xlogscale_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(xaxis_grid, text="Logarithmic X-axis", variable=self.xlogscale_var, 
                      command=self.update_xlog_options).grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Base:", width=6, anchor="e").grid(row=2, column=2, sticky="e", pady=2)
        self.xlog_base_var = tk.StringVar(value="10")
        self.xlog_base_dropdown = ttk.Combobox(xaxis_grid, textvariable=self.xlog_base_var, 
                                          values=["10", "2"], state="disabled", width=5)
        self.xlog_base_dropdown.grid(row=2, column=3, sticky="w", pady=2)
        
        # --- Y-Axis settings with grid layout ---
        yaxis_grp = ttk.LabelFrame(frame, text="Y-Axis Settings", padding=4)
        yaxis_grp.pack(fill='x', padx=4, pady=1)
        
        # Use grid layout for better alignment
        yaxis_grid = tk.Frame(yaxis_grp)
        yaxis_grid.pack(fill='x', padx=2, pady=1)
        
        # Configure columns for alignment - fixed label widths
        yaxis_grid.columnconfigure(1, weight=0) # Min Entry
        yaxis_grid.columnconfigure(3, weight=0) # Max Entry
        
        # Row 1: Min/Max values
        tk.Label(yaxis_grid, text="Minimum:", width=12, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.ymin_entry = ttk.Entry(yaxis_grid, width=10)
        self.ymin_entry.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Maximum:", width=12, anchor="w").grid(row=0, column=2, sticky="w", padx=(10,0), pady=2)
        self.ymax_entry = ttk.Entry(yaxis_grid, width=10)
        self.ymax_entry.grid(row=0, column=3, sticky="w", pady=2)
        
        # Row 2: Tick settings
        tk.Label(yaxis_grid, text="Major Interval:", width=12, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.yinterval_entry = ttk.Entry(yaxis_grid, width=10)
        self.yinterval_entry.grid(row=1, column=1, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Minor/Major:", width=12, anchor="w").grid(row=1, column=2, sticky="w", padx=(10,0), pady=2)
        self.minor_ticks_entry = ttk.Entry(yaxis_grid, width=10)
        self.minor_ticks_entry.grid(row=1, column=3, sticky="w", pady=2)
        
        # Row 3: Log options
        self.logscale_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(yaxis_grid, text="Logarithmic Y-axis", variable=self.logscale_var, 
                      command=self.update_ylog_options).grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Base:", width=6, anchor="e").grid(row=2, column=2, sticky="e", pady=2)
        self.ylog_base_var = tk.StringVar(value="10")
        self.ylog_base_dropdown = ttk.Combobox(yaxis_grid, textvariable=self.ylog_base_var, 
                                          values=["10", "2"], state="disabled", width=5)
        self.ylog_base_dropdown.grid(row=2, column=3, sticky="w", pady=2)

        if not hasattr(self, 'custom_x_marks_var'):
            self.custom_x_marks_var = tk.StringVar(value="")
        if not hasattr(self, 'custom_y_marks_var'):
            self.custom_y_marks_var = tk.StringVar(value="")

        custom_row = ttk.Frame(frame)
        custom_row.pack(fill='x', padx=4, pady=(6, 1))
        ttk.Label(custom_row, text="Custom marks:", width=14, anchor='w').pack(side='left')
        ttk.Button(custom_row, text="Edit…", command=self.open_custom_marks_editor, width=10).pack(side='left')
        
        # --- Y-Axis Break settings ---
        ybreak_grp = ttk.LabelFrame(frame, text="Y-Axis Break", padding=4)
        ybreak_grp.pack(fill='x', padx=4, pady=4)
        
        # Enable axis break checkbox
        self.yaxis_break_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ybreak_grp, text="Enable Y-axis break", variable=self.yaxis_break_var,
                       command=self.update_ybreak_options).pack(anchor="w", pady=2)
        
        # Break range settings with grid layout
        ybreak_grid = tk.Frame(ybreak_grp)
        ybreak_grid.pack(fill='x', padx=2, pady=1)
        
        tk.Label(ybreak_grid, text="Break from:", width=12, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.ybreak_min_entry = ttk.Entry(ybreak_grid, width=10, state="disabled")
        self.ybreak_min_entry.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(ybreak_grid, text="to:", width=4, anchor="w").grid(row=0, column=2, sticky="w", padx=(10,0), pady=2)
        self.ybreak_max_entry = ttk.Entry(ybreak_grid, width=10, state="disabled")
        self.ybreak_max_entry.grid(row=0, column=3, sticky="w", pady=2)
        
        # Height ratio for upper/lower sections
        tk.Label(ybreak_grid, text="Upper ratio:", width=12, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.ybreak_ratio_var = tk.DoubleVar(value=0.5)
        self.ybreak_ratio_entry = ttk.Entry(ybreak_grid, width=10, state="disabled")
        self.ybreak_ratio_entry.grid(row=1, column=1, sticky="w", pady=2)
        self.ybreak_ratio_entry.insert(0, "0.5")
        tk.Label(ybreak_grid, text="(0.1-0.9)", width=10, anchor="w").grid(row=1, column=2, columnspan=2, sticky="w", padx=(10,0), pady=2)

        tk.Label(ybreak_grid, text="Gap:", width=12, anchor="w").grid(row=2, column=0, sticky="w", pady=2)
        self.ybreak_gap_entry = ttk.Entry(ybreak_grid, width=10, state="disabled")
        self.ybreak_gap_entry.grid(row=2, column=1, sticky="w", pady=2)
        self.ybreak_gap_entry.insert(0, "0.07")
        tk.Label(ybreak_grid, text="(0-0.3)", width=10, anchor="w").grid(row=2, column=2, columnspan=2, sticky="w", padx=(10,0), pady=2)

        tk.Label(ybreak_grid, text="Marker style:", width=12, anchor="w").grid(row=3, column=0, sticky="w", pady=2)
        self.ybreak_marker_style_dropdown = ttk.Combobox(
            ybreak_grid,
            textvariable=self.ybreak_marker_style_var,
            values=["Diagonal", "Connected"],
            state="disabled",
            width=10
        )
        self.ybreak_marker_style_dropdown.grid(row=3, column=1, sticky="w", pady=2)

        try:
            def _on_ybreak_marker_style_selected(_event=None):
                if hasattr(self, 'ybreak_marker_style_user_set_var'):
                    self.ybreak_marker_style_user_set_var.set(True)
            self.ybreak_marker_style_dropdown.bind('<<ComboboxSelected>>', _on_ybreak_marker_style_selected)
        except Exception:
            pass

    def update_ybreak_options(self):
        """Enable or disable Y-axis break options based on checkbox state"""
        if self.yaxis_break_var.get():
            self.ybreak_min_entry.config(state="normal")
            self.ybreak_max_entry.config(state="normal")
            self.ybreak_ratio_entry.config(state="normal")
            if hasattr(self, 'ybreak_gap_entry'):
                self.ybreak_gap_entry.config(state="normal")
            if hasattr(self, 'ybreak_marker_style_dropdown'):
                self.ybreak_marker_style_dropdown.config(state="readonly")
        else:
            self.ybreak_min_entry.config(state="disabled")
            self.ybreak_max_entry.config(state="disabled")
            self.ybreak_ratio_entry.config(state="disabled")
            if hasattr(self, 'ybreak_gap_entry'):
                self.ybreak_gap_entry.config(state="disabled")
            if hasattr(self, 'ybreak_marker_style_dropdown'):
                self.ybreak_marker_style_dropdown.config(state="disabled")

    def update_xlog_options(self):
        """Enable or disable X-axis log options based on checkbox state"""
        if self.xlogscale_var.get():
            self.xlog_base_dropdown.config(state="readonly")
        else:
            self.xlog_base_dropdown.config(state="disabled")
    
    def update_ylog_options(self):
        """Enable or disable Y-axis log options based on checkbox state"""
        if self.logscale_var.get():
            self.ylog_base_dropdown.config(state="readonly")
        else:
            self.ylog_base_dropdown.config(state="disabled")
    
    def setup_xy_fitting_tab(self):
        frame = self.xy_fitting_tab
        
        # Enable fitting checkbox
        fit_enable_frame = tk.Frame(frame)
        fit_enable_frame.pack(fill='x', padx=6, pady=(8,4))
        
        # Set up a function to handle enabling/disabling fitting
        def toggle_fitting():
            if self.use_fitting_var.get():
                self.debug("Fitting enabled, setting plot type to XY")
                self.plot_kind_var.set("xy")
            else:
                self.debug("Fitting disabled")

        self.use_fitting_cb = ttk.Checkbutton(fit_enable_frame, text="Enable Model Fitting", 
                                             variable=self.use_fitting_var, 
                                             command=toggle_fitting)
        self.use_fitting_cb.pack(anchor="w", pady=2)
        
        # Model selection group
        model_grp = ttk.LabelFrame(frame, text="Model Selection", padding=6)
        model_grp.pack(fill='x', padx=6, pady=4)
        
        # Model dropdown
        tk.Label(model_grp, text="Fitting Model:").grid(row=0, column=0, sticky="w", pady=2)
        self.model_dropdown = ttk.Combobox(model_grp, textvariable=self.fitting_model_var, 
                                           values=sorted(list(self.fitting_models.keys())), width=25, state="readonly")
        self.model_dropdown.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        self.model_dropdown.bind('<<ComboboxSelected>>', self.update_model_parameters)
        
        # Confidence interval options
        tk.Label(model_grp, text="Confidence Interval:").grid(row=1, column=0, sticky="w", pady=2)
        self.ci_dropdown = ttk.Combobox(model_grp, textvariable=self.fitting_ci_var, 
                                        values=["None", "68% (1σ)", "95% (2σ)"], width=25, state="readonly")
        self.ci_dropdown.grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        
        # Color options for fit lines and bands
        tk.Label(model_grp, text="Appearance:").grid(row=2, column=0, sticky="w", pady=2)
        fit_color_frame = tk.Frame(model_grp)
        fit_color_frame.grid(row=2, column=1, sticky="ew", padx=2, pady=2)
        
        # Option to use black lines
        self.fit_black_lines_cb = ttk.Checkbutton(fit_color_frame, text="Black Lines", 
                                               variable=self.fitting_use_black_lines_var)
        self.fit_black_lines_cb.pack(side=tk.LEFT, padx=2)
        
        # Option to use black bands for confidence intervals
        self.fit_black_bands_cb = ttk.Checkbutton(fit_color_frame, text="Black Bands", 
                                               variable=self.fitting_use_black_bands_var)
        self.fit_black_bands_cb.pack(side=tk.LEFT, padx=2)
        
        # Option to match group colors
        self.fit_group_cb = ttk.Checkbutton(fit_color_frame, text="Match Groups", 
                                          variable=self.fitting_use_group_colors_var)
        self.fit_group_cb.pack(side=tk.LEFT, padx=2)
        
        # Button to manage models
        manage_models_btn = ttk.Button(model_grp, text="Manage Models", command=self.manage_fitting_models)
        manage_models_btn.grid(row=3, column=0, columnspan=2, sticky="ew", padx=2, pady=6)
        
        # Parameters group
        self.params_grp = ttk.LabelFrame(frame, text="Model Parameters", padding=6)
        self.params_grp.pack(fill='x', padx=6, pady=4)
        
        # Description display group
        description_grp = ttk.LabelFrame(frame, text="Model Description", padding=6)
        description_grp.pack(fill='x', padx=6, pady=4)
        
        description_frame = tk.Frame(description_grp)
        description_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        self.description_text = tk.Text(description_frame, height=4, width=40, wrap=tk.WORD)
        description_scrollbar = ttk.Scrollbar(description_frame, command=self.description_text.yview)
        self.description_text.config(yscrollcommand=description_scrollbar.set)
        
        self.description_text.pack(side=tk.LEFT, fill='both', expand=True)
        description_scrollbar.pack(side=tk.RIGHT, fill='y')
        
        # Formula display group
        formula_grp = ttk.LabelFrame(frame, text="Model Formula", padding=6)
        formula_grp.pack(fill='x', padx=6, pady=4, expand=True)
        
        # Formula display with scrollbar
        formula_frame = tk.Frame(formula_grp)
        formula_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        self.formula_text = tk.Text(formula_frame, height=5, width=40, wrap=tk.WORD)
        formula_scrollbar = ttk.Scrollbar(formula_frame, command=self.formula_text.yview)
        self.formula_text.config(yscrollcommand=formula_scrollbar.set)
        
        self.formula_text.pack(side=tk.LEFT, fill='both', expand=True)
        formula_scrollbar.pack(side=tk.RIGHT, fill='y')
        
        # Result display group
        result_grp = ttk.LabelFrame(frame, text="Fitting Results", padding=6)
        result_grp.pack(fill='x', padx=6, pady=4)
        
        # Results text with scrollbar
        result_frame = tk.Frame(result_grp)
        result_frame.pack(fill='both', expand=True, padx=2, pady=2)
        
        self.result_text = tk.Text(result_frame, height=8, width=40, wrap=tk.WORD)
        result_scrollbar = ttk.Scrollbar(result_frame, command=self.result_text.yview)
        self.result_text.config(yscrollcommand=result_scrollbar.set)
        
        self.result_text.pack(side=tk.LEFT, fill='both', expand=True)
        result_scrollbar.pack(side=tk.RIGHT, fill='y')
        
        # Initialize parameter fields
        self.param_entries = []
        self.update_model_parameters()
        
    def update_model_parameters(self, event=None):
        """Update parameter entry fields based on selected model"""
        # Clear existing parameter entries
        for widget in self.params_grp.winfo_children():
            widget.destroy()
        self.param_entries = []
        
        # Get selected model parameters
        model_name = self.fitting_model_var.get()
        model_info = self.fitting_models.get(model_name, {})
        parameters = model_info.get("parameters", [])
        formula = model_info.get("formula", "")
        
        # Update formula display
        self.formula_text.delete(1.0, tk.END)
        self.formula_text.insert(tk.END, formula)
        
        # Update description display
        description = model_info.get("description", "No description available.")
        self.description_text.delete(1.0, tk.END)
        self.description_text.insert(tk.END, description)
        
        # Create parameter entry fields
        for i, (param_name, default_value) in enumerate(parameters):
            frame = tk.Frame(self.params_grp)
            frame.pack(fill='x', pady=2)
            
            label = tk.Label(frame, text=f"{param_name} starting value:")
            label.pack(side=tk.LEFT, padx=2)
            
            var = tk.DoubleVar(value=default_value)
            entry = ttk.Entry(frame, textvariable=var, width=10)
            entry.pack(side=tk.RIGHT, padx=2)
            
            self.param_entries.append((param_name, var))
    
    def save_project(self):
        """Save the current plot settings to a file"""
        try:
            # Prompt for a filename to save the project to
            file_path = filedialog.asksaveasfilename(
                title="Save Project",
                defaultextension=".explt",
                filetypes=[("ExPlot Project Files", "*.explt"), ("All Files", "*.*")]
            )
            if not file_path:
                return  # User cancelled
                
            # Create a dictionary with all the current settings
            settings = self.get_current_settings()
            
            # Add the XY Fitting settings to the project
            settings['xy_fitting'] = {
                'use_fitting': self.use_fitting_var.get(),
                'fitting_model': self.fitting_model_var.get(),
                'fitting_ci': self.fitting_ci_var.get(),
                'fitting_models': self.fitting_models,
                'parameters': [(name, var.get()) for name, var in self.param_entries]
            }
            
            # Save to JSON file
            with open(file_path, 'w') as f:
                json.dump(settings, f, indent=2)
                
            messagebox.showinfo("Success", f"Project saved to {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error Saving Project", f"An error occurred: {str(e)}")
            
    def get_outline_color(self, color=None):
        """Get the outline color based on the outline_color_var setting.
        
        Parameters:
        color: Optional default color (used for as_set mode)
        
        Returns:
        The outline color to use
        """
        outline_color_setting = self.outline_color_var.get()
        
        if outline_color_setting == "as_set" and color:
            return color
        elif outline_color_setting == "black":
            return "black"
        elif outline_color_setting == "gray":
            return "gray"
        elif outline_color_setting == "white":
            return "white"
        else:
            # Default if no valid selection or no color provided for as_set
            return "black"
    
    def get_current_settings(self):
        """Gather current application settings for saving"""
        settings = {
            'version': self.version,
            'excel_file': self.excel_file,  # Keep as reference only
            'sheet': self.sheet_var.get() if hasattr(self, 'sheet_var') else '',
            'plot_kind': self.plot_kind_var.get(),
            'columns': {
                'x_axis': self.xaxis_var.get() if hasattr(self, 'xaxis_var') else '',
                'group': self.group_var.get() if hasattr(self, 'group_var') else '',
                'y_axis': [col for var, col in self.value_vars if var.get()] if hasattr(self, 'value_vars') else []
            },
            
            # Embed both raw and modified data and any customizations
            'embedded_data': {
                # Store the raw data
                'raw_dataframe': self.raw_df.to_dict() if hasattr(self, 'raw_df') and self.raw_df is not None else 
                               (self.df.to_dict() if hasattr(self, 'df') and self.df is not None else None),
                # Store the modified data
                'modified_dataframe': self.df.to_dict() if hasattr(self, 'df') and self.df is not None else None,
                'xaxis_renames': self.xaxis_renames if hasattr(self, 'xaxis_renames') else {},
                'xaxis_order': self.xaxis_order if hasattr(self, 'xaxis_order') else None,
                'excluded_x_values': list(self.excluded_x_values) if hasattr(self, 'excluded_x_values') else []
            },
            'appearance': {
                'plot_width': self.plot_width_var.get(),
                'plot_height': self.plot_height_var.get(),
                'font_size': self.fontsize_entry.get() if hasattr(self, 'fontsize_entry') else '10',
                'line_width': self.linewidth.get(),
                'swap_axes': self.swap_axes_var.get() if hasattr(self, 'swap_axes_var') else False,
                'show_frame': self.show_frame_var.get() if hasattr(self, 'show_frame_var') else False,
                'show_hgrid': self.show_hgrid_var.get() if hasattr(self, 'show_hgrid_var') else False,
                'show_vgrid': self.show_vgrid_var.get() if hasattr(self, 'show_vgrid_var') else False,
                'bar_gap_multiplier': self.bar_gap_multiplier_var.get() if hasattr(self, 'bar_gap_multiplier_var') else 0.75
            },
            'axis': {
                'x_label': self.xlabel_entry.get() if hasattr(self, 'xlabel_entry') else '',
                'y_label': self.ylabel_entry.get() if hasattr(self, 'ylabel_entry') else '',
                'x_orientation': self.label_orientation.get() if hasattr(self, 'label_orientation') else 'vertical',
                'x_min': self.xmin_entry.get() if hasattr(self, 'xmin_entry') else '',
                'x_max': self.xmax_entry.get() if hasattr(self, 'xmax_entry') else '',
                'y_min': self.ymin_entry.get() if hasattr(self, 'ymin_entry') else '',
                'y_max': self.ymax_entry.get() if hasattr(self, 'ymax_entry') else '',
                'x_log': self.xlogscale_var.get() if hasattr(self, 'xlogscale_var') else False,
                'y_log': self.logscale_var.get() if hasattr(self, 'logscale_var') else False,
                'x_log_base': self.xlog_base_var.get() if hasattr(self, 'xlog_base_var') else '10',
                'y_log_base': self.ylog_base_var.get() if hasattr(self, 'ylog_base_var') else '10',
                'custom_x_marks': self.custom_x_marks_var.get() if hasattr(self, 'custom_x_marks_var') else '',
                'custom_y_marks': self.custom_y_marks_var.get() if hasattr(self, 'custom_y_marks_var') else ''
            },
            'statistics': {
                'use_stats': self.use_stats_var.get() if hasattr(self, 'use_stats_var') else False,
                'ttest_type': self.ttest_type_var.get() if hasattr(self, 'ttest_type_var') else '',
                'ttest_alternative': self.ttest_alternative_var.get() if hasattr(self, 'ttest_alternative_var') else '',
                'anova_type': self.anova_type_var.get() if hasattr(self, 'anova_type_var') else '',
                'posthoc_type': self.posthoc_type_var.get() if hasattr(self, 'posthoc_type_var') else '',
                'alpha_level': self.alpha_level_var.get() if hasattr(self, 'alpha_level_var') else '0.05'
            },
            'xy_plot': {
                'marker_symbol': self.xy_marker_symbol_var.get() if hasattr(self, 'xy_marker_symbol_var') else 'o',
                'marker_size': self.xy_marker_size_var.get() if hasattr(self, 'xy_marker_size_var') else 5.0,
                'filled': self.xy_filled_var.get() if hasattr(self, 'xy_filled_var') else True,
                'line_style': self.xy_line_style_var.get() if hasattr(self, 'xy_line_style_var') else 'solid',
                'line_black': self.xy_line_black_var.get() if hasattr(self, 'xy_line_black_var') else False,
                'connect': self.xy_connect_var.get() if hasattr(self, 'xy_connect_var') else False,
                'show_mean': self.xy_show_mean_var.get() if hasattr(self, 'xy_show_mean_var') else True,
                'show_mean_errorbars': self.xy_show_mean_errorbars_var.get() if hasattr(self, 'xy_show_mean_errorbars_var') else True,
                'draw_band': self.xy_draw_band_var.get() if hasattr(self, 'xy_draw_band_var') else False
            },
            'colors': {
                'single_color': self.single_color_var.get() if hasattr(self, 'single_color_var') else list(self.custom_colors.keys())[0],
                'palette': self.palette_var.get() if hasattr(self, 'palette_var') else list(self.custom_palettes.keys())[0],
                'bar_outline': self.bar_outline_var.get() if hasattr(self, 'bar_outline_var') else False,
                'outline_color': self.outline_color_var.get() if hasattr(self, 'outline_color_var') else 'as_set',
                'strip_black': self.strip_black_var.get() if hasattr(self, 'strip_black_var') else True
            },
            'x_axis_renames': self.xaxis_renames,
            'x_axis_order': self.xaxis_order
        }
        
        # Add XY Fitting settings
        settings['xy_fitting'] = {
            'use_fitting': self.use_fitting_var.get(),
            'fitting_model': self.fitting_model_var.get(),
            'fitting_ci': self.fitting_ci_var.get(),
            'fitting_models': self.fitting_models,
            'parameters': [(name, var.get()) for name, var in self.param_entries]
        }
        
        return settings
        
    def apply_settings(self, settings):
        """Apply loaded settings to the current state"""
        try:
            # Check if the file is from an older version
            file_version = settings.get('version', '0.0.0')
            current_version = self.version
            
            # Handle embedded data if present
            if 'embedded_data' in settings:
                # Initialize raw and modified dataframes
                raw_df = None
                modified_df = None
                
                # Load raw dataframe if available
                if settings['embedded_data'].get('raw_dataframe'):
                    raw_df = pd.DataFrame.from_dict(settings['embedded_data']['raw_dataframe'])
                
                # Load modified dataframe if available
                if settings['embedded_data'].get('modified_dataframe'):
                    modified_df = pd.DataFrame.from_dict(settings['embedded_data']['modified_dataframe'])
                # If only one is available, use it for both
                elif settings['embedded_data'].get('dataframe'):  # For backward compatibility
                    modified_df = pd.DataFrame.from_dict(settings['embedded_data']['dataframe'])
                    raw_df = modified_df.copy()
                
                # If we have valid dataframes
                if raw_df is not None or modified_df is not None:
                    # Set up the dataframes
                    self.raw_df = raw_df if raw_df is not None else modified_df.copy()
                    self.df = modified_df if modified_df is not None else raw_df.copy()
                    
                    # Load customizations
                    self.xaxis_renames = settings['embedded_data'].get('xaxis_renames', {})
                    self.xaxis_order = settings['embedded_data'].get('xaxis_order', [])
                    self.excluded_x_values = set(settings['embedded_data'].get('excluded_x_values', []))
                    
                    # Update sheet dropdown to show both embedded data options
                    self.sheet_options = ['Raw Embedded Data', 'Modified Embedded Data']
                    if hasattr(self, 'sheet_dropdown'):
                        self.sheet_dropdown['values'] = self.sheet_options
                    if hasattr(self, 'sheet_var'):
                        self.sheet_var.set('Modified Embedded Data')
                        
                    # We initially work with the modified data
                    self.df = self.modified_df if hasattr(self, 'modified_df') else self.df
                
                # Update dropdowns with available columns
                self.update_columns()
            # Otherwise try loading from Excel file
            elif 'excel_file' in settings and os.path.exists(settings['excel_file']):
                self.excel_file = settings['excel_file']
                self.load_file(settings['excel_file'])
                
                # Select sheet
                if 'sheet' in settings and settings['sheet'] in self.sheet_dropdown['values']:
                    self.sheet_var.set(settings['sheet'])
                    self.load_sheet()
            
            # Columns selection
            if 'columns' in settings:
                cols = settings['columns']
                if 'x_axis' in cols and cols['x_axis'] in self.xaxis_dropdown['values']:
                    self.xaxis_var.set(cols['x_axis'])
                
                if 'group' in cols and cols['group'] in self.group_dropdown['values']:
                    self.group_var.set(cols['group'])
                
                if 'y_axis' in cols and hasattr(self, 'value_vars'):
                    for var, col in self.value_vars:
                        var.set(col in cols['y_axis'])
            
            # Plot kind
            if 'plot_kind' in settings:
                self.plot_kind_var.set(settings['plot_kind'])
            
            # Appearance
            if 'appearance' in settings:
                app = settings['appearance']
                if 'plot_width' in app:
                    self.plot_width_var.set(float(app['plot_width']))
                if 'plot_height' in app:
                    self.plot_height_var.set(float(app['plot_height']))
                if 'font_size' in app and hasattr(self, 'fontsize_entry'):
                    self.fontsize_entry.delete(0, tk.END)
                    self.fontsize_entry.insert(0, app['font_size'])
                if 'line_width' in app:
                    self.linewidth.set(float(app['line_width']))
                if 'swap_axes' in app and hasattr(self, 'swap_axes_var'):
                    self.swap_axes_var.set(app['swap_axes'])
                if 'show_frame' in app and hasattr(self, 'show_frame_var'):
                    self.show_frame_var.set(app['show_frame'])
                if 'show_hgrid' in app and hasattr(self, 'show_hgrid_var'):
                    self.show_hgrid_var.set(app['show_hgrid'])
                if 'show_vgrid' in app and hasattr(self, 'show_vgrid_var'):
                    self.show_vgrid_var.set(app['show_vgrid'])
                if 'bar_gap_multiplier' in app and hasattr(self, 'bar_gap_multiplier_var'):
                    self.bar_gap_multiplier_var.set(float(app['bar_gap_multiplier']))
            
            # Axis settings
            if 'axis' in settings:
                axis = settings['axis']
                if 'x_label' in axis and hasattr(self, 'xlabel_entry'):
                    self.xlabel_entry.delete(0, tk.END)
                    self.xlabel_entry.insert(0, axis['x_label'])
                if 'y_label' in axis and hasattr(self, 'ylabel_entry'):
                    self.ylabel_entry.delete(0, tk.END)
                    self.ylabel_entry.insert(0, axis['y_label'])
                if 'x_orientation' in axis and hasattr(self, 'label_orientation'):
                    self.label_orientation.set(axis['x_orientation'])
                if 'x_min' in axis and hasattr(self, 'xmin_entry'):
                    self.xmin_entry.delete(0, tk.END)
                    self.xmin_entry.insert(0, axis['x_min'])
                if 'x_max' in axis and hasattr(self, 'xmax_entry'):
                    self.xmax_entry.delete(0, tk.END)
                    self.xmax_entry.insert(0, axis['x_max'])
                if 'y_min' in axis and hasattr(self, 'ymin_entry'):
                    self.ymin_entry.delete(0, tk.END)
                    self.ymin_entry.insert(0, axis['y_min'])
                if 'y_max' in axis and hasattr(self, 'ymax_entry'):
                    self.ymax_entry.delete(0, tk.END)
                    self.ymax_entry.insert(0, axis['y_max'])
                if 'x_log' in axis and hasattr(self, 'xlogscale_var'):
                    self.xlogscale_var.set(axis['x_log'])
                    self.update_xlog_options()
                if 'y_log' in axis and hasattr(self, 'logscale_var'):
                    self.logscale_var.set(axis['y_log'])
                    self.update_ylog_options()
                if 'x_log_base' in axis and hasattr(self, 'xlog_base_var'):
                    self.xlog_base_var.set(axis['x_log_base'])
                if 'y_log_base' in axis and hasattr(self, 'ylog_base_var'):
                    self.ylog_base_var.set(axis['y_log_base'])
                if 'custom_x_marks' in axis and hasattr(self, 'custom_x_marks_var'):
                    self.custom_x_marks_var.set(axis['custom_x_marks'])
                if 'custom_y_marks' in axis and hasattr(self, 'custom_y_marks_var'):
                    self.custom_y_marks_var.set(axis['custom_y_marks'])
            
            # Statistics settings
            if 'statistics' in settings:
                stats = settings['statistics']
                if 'use_stats' in stats and hasattr(self, 'use_stats_var'):
                    self.use_stats_var.set(stats['use_stats'])
                if 'ttest_type' in stats and hasattr(self, 'ttest_type_var'):
                    self.ttest_type_var.set(stats['ttest_type'])
                if 'ttest_alternative' in stats and hasattr(self, 'ttest_alternative_var'):
                    self.ttest_alternative_var.set(stats['ttest_alternative'])
                if 'anova_type' in stats and hasattr(self, 'anova_type_var'):
                    self.anova_type_var.set(stats['anova_type'])
                if 'posthoc_type' in stats and hasattr(self, 'posthoc_type_var'):
                    self.posthoc_type_var.set(stats['posthoc_type'])
                if 'alpha_level' in stats and hasattr(self, 'alpha_level_var'):
                    self.alpha_level_var.set(stats['alpha_level'])
            
            # XY plot settings
            if 'xy_plot' in settings:
                xy = settings['xy_plot']
                if 'marker_symbol' in xy and hasattr(self, 'xy_marker_symbol_var'):
                    self.xy_marker_symbol_var.set(xy['marker_symbol'])
                if 'marker_size' in xy and hasattr(self, 'xy_marker_size_var'):
                    self.xy_marker_size_var.set(float(xy['marker_size']))
                if 'filled' in xy and hasattr(self, 'xy_filled_var'):
                    self.xy_filled_var.set(xy['filled'])
                if 'line_style' in xy and hasattr(self, 'xy_line_style_var'):
                    self.xy_line_style_var.set(xy['line_style'])
                if 'line_black' in xy and hasattr(self, 'xy_line_black_var'):
                    self.xy_line_black_var.set(xy['line_black'])
                if 'connect' in xy and hasattr(self, 'xy_connect_var'):
                    self.xy_connect_var.set(xy['connect'])
                if 'show_mean' in xy and hasattr(self, 'xy_show_mean_var'):
                    self.xy_show_mean_var.set(xy['show_mean'])
                    self.update_xy_mean_errorbar_state()
                if 'show_mean_errorbars' in xy and hasattr(self, 'xy_show_mean_errorbars_var'):
                    self.xy_show_mean_errorbars_var.set(xy['show_mean_errorbars'])
                if 'draw_band' in xy and hasattr(self, 'xy_draw_band_var'):
                    self.xy_draw_band_var.set(xy['draw_band'])
            
            # Color settings
            if 'colors' in settings:
                colors = settings['colors']
                if 'single_color' in colors and hasattr(self, 'single_color_var'):
                    if colors['single_color'] in self.single_color_dropdown['values']:
                        self.single_color_var.set(colors['single_color'])
                if 'palette' in colors and hasattr(self, 'palette_var'):
                    if colors['palette'] in self.palette_dropdown['values']:
                        self.palette_var.set(colors['palette'])
                if 'bar_outline' in colors and hasattr(self, 'bar_outline_var'):
                    self.bar_outline_var.set(colors['bar_outline'])
                if 'outline_color' in colors and hasattr(self, 'outline_color_var'):
                    self.outline_color_var.set(colors['outline_color'])
                if 'strip_black' in colors and hasattr(self, 'strip_black_var'):
                    self.strip_black_var.set(colors['strip_black'])
            
            # X-axis customizations
            if 'x_axis_renames' in settings:
                self.xaxis_renames = settings['x_axis_renames']
            if 'x_axis_order' in settings:
                self.xaxis_order = settings['x_axis_order']
            
            # Apply XY Fitting settings if present
            if 'xy_fitting' in settings and hasattr(self, 'use_fitting_var'):
                xy_settings = settings['xy_fitting']
                
                # Load fitting models if present
                if 'fitting_models' in xy_settings:
                    # Get the saved models
                    saved_models = xy_settings['fitting_models']
                    
                    # Merge saved models with current models (prioritize saved models for existing names)
                    # This ensures newly added models remain available
                    for model_name, model_info in saved_models.items():
                        self.fitting_models[model_name] = model_info
                    
                    # Update the dropdown with all available models
                    if hasattr(self, 'model_dropdown'):
                        self.model_dropdown['values'] = sorted(list(self.fitting_models.keys()))
                    
                # Set the fitting enabled state
                if 'use_fitting' in xy_settings:
                    self.use_fitting_var.set(xy_settings['use_fitting'])
                    
                # Set the model and CI options
                if 'fitting_model' in xy_settings and xy_settings['fitting_model'] in self.fitting_models:
                    self.fitting_model_var.set(xy_settings['fitting_model'])
                    self.update_model_parameters()
                    
                if 'fitting_ci' in xy_settings:
                    self.fitting_ci_var.set(xy_settings['fitting_ci'])
                    
                # Set parameter values if available
                if 'parameters' in xy_settings and hasattr(self, 'param_entries'):
                    params = xy_settings['parameters']
                    for i, (param_name, param_val) in enumerate(params):
                        if i < len(self.param_entries):
                            self.param_entries[i][1].set(param_val)
            
        except Exception as e:
            print(f"Error applying settings: {e}")
            raise
            
    def load_project(self):
        """Load plot settings from a project file"""
        try:
            # Prompt for a project file to load
            file_path = filedialog.askopenfilename(
                title="Open Project",
                filetypes=[("ExPlot Project Files", "*.explt"), ("All Files", "*.*")]
            )
            if not file_path:
                return  # User cancelled
                
            # Load the settings from the file
            with open(file_path, 'r') as f:
                settings = json.load(f)
                
            # Apply the settings to the current state
            self.apply_settings(settings)
            
            messagebox.showinfo("Success", f"Project loaded from {file_path}")
            
        except Exception as e:
            messagebox.showerror("Error Loading Project", f"An error occurred: {str(e)}")
    
    def manage_fitting_models(self):
        """Open a dialog to manage fitting models"""
        # Create a top-level window
        dialog = tk.Toplevel(self.root)
        dialog.title("Manage Fitting Models")
        dialog.geometry("700x600")
        dialog.transient(self.root)  # Set to be on top of the main window
        dialog.grab_set()  # Modal
        
        # Left side - model list
        left_frame = tk.Frame(dialog, borderwidth=1, relief=tk.SUNKEN)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        tk.Label(left_frame, text="Models:").pack(anchor=tk.W, padx=5, pady=5)
        
        # Scrollable list of models
        model_listbox_frame = tk.Frame(left_frame)
        model_listbox_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview for model list with a single column
        model_tree = ttk.Treeview(model_listbox_frame, selectmode='browse', show='tree', height=10)
        model_tree.column('#0', width=200, stretch=tk.YES)
        model_tree.heading('#0', text='Models', anchor='w')
        model_scrollbar = ttk.Scrollbar(model_listbox_frame, orient='vertical', command=model_tree.yview)
        model_tree.configure(yscrollcommand=model_scrollbar.set)
        
        model_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        model_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add models to treeview
        for model in sorted(self.fitting_models.keys()):
            model_tree.insert('', 'end', text=model, values=(model,))
        
        # Buttons frame for model management
        btn_frame = tk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        add_btn = ttk.Button(btn_frame, text="Add New Model", 
                          command=lambda: self.add_new_model(model_tree, dialog))
        add_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        remove_btn = ttk.Button(btn_frame, text="Remove Model", 
                             command=lambda: self.remove_model(model_tree, dialog))
        remove_btn.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=2)
        
        # Additional buttons frame
        extra_btn_frame = tk.Frame(left_frame)
        extra_btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        restore_btn = ttk.Button(extra_btn_frame, text="Restore Default Models", 
                             command=lambda: self.restore_default_models(model_tree))
        restore_btn.pack(fill=tk.X, expand=True)
        
        # Right side - model details
        right_frame = tk.Frame(dialog)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Model name
        name_frame = tk.Frame(right_frame)
        name_frame.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(name_frame, text="Model Name:").pack(side=tk.LEFT, padx=2)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(name_frame, textvariable=name_var, width=30)
        name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        
        # Description
        desc_frame = ttk.LabelFrame(right_frame, text="Description")
        desc_frame.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        desc_text = tk.Text(desc_frame, height=9, width=40)
        desc_scrollbar = ttk.Scrollbar(desc_frame, command=desc_text.yview)
        desc_text.config(yscrollcommand=desc_scrollbar.set)
        
        desc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        desc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Parameters
        param_frame = ttk.LabelFrame(right_frame, text="Parameters (name, default value)")
        param_frame.pack(fill=tk.BOTH, expand=False, padx=5, pady=5)
        
        param_text = tk.Text(param_frame, height=5, width=40)
        param_scrollbar = ttk.Scrollbar(param_frame, command=param_text.yview)
        param_text.config(yscrollcommand=param_scrollbar.set)
        
        param_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        param_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Formula
        formula_frame = ttk.LabelFrame(right_frame, text="Formula")
        formula_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        formula_text = tk.Text(formula_frame, height=4, width=40)
        formula_scrollbar = ttk.Scrollbar(formula_frame, command=formula_text.yview)
        formula_text.config(yscrollcommand=formula_scrollbar.set)
        
        formula_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        formula_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Save button
        save_btn = ttk.Button(right_frame, text="Save Model", 
                           command=lambda: self.save_model(name_var, desc_text, param_text, formula_text, model_listbox, dialog))
        save_btn.pack(fill=tk.X, padx=5, pady=5)
        
        # Function to update right panel when model is selected
        def on_model_select(event):
            selection = model_tree.selection()
            if selection:
                item = model_tree.item(selection[0])
                model_name = item['text']
                model_info = self.fitting_models.get(model_name, {})
                
                name_var.set(model_name)
                
                # Clear and fill description text
                desc_text.delete(1.0, tk.END)
                desc_text.insert(tk.END, model_info.get("description", ""))
                
                # Clear and fill parameter text
                param_text.delete(1.0, tk.END)
                parameters = model_info.get("parameters", [])
                param_str = "\n".join([f"{name}, {value}" for name, value in parameters])
                param_text.insert(tk.END, param_str)
                
                # Clear and fill formula text
                formula_text.delete(1.0, tk.END)
                formula_text.insert(tk.END, model_info.get("formula", ""))
        
        model_tree.bind('<<TreeviewSelect>>', on_model_select)
        
    def add_new_model(self, tree, dialog):
        """Add a new empty model to the list"""
        count = 1
        new_name = f"New Model {count}"
        while new_name in self.fitting_models:
            count += 1
            new_name = f"New Model {count}"
            
        # Add to models dictionary and treeview
        self.fitting_models[new_name] = {
            "parameters": [("A", 1.0), ("B", 1.0)],
            "formula": "# formula here\ny = A * x + B",
            "description": "New model - add description here"
        }
        
        # Save to file
        self.save_fitting_models()
        
        # Add to treeview and select it
        tree.insert('', 'end', text=new_name, values=(new_name,))
        self.model_dropdown['values'] = sorted(list(self.fitting_models.keys()))
        
        # Select the new model
        for item in tree.get_children():
            if tree.item(item, 'text') == new_name:
                tree.selection_set(item)
                tree.focus(item)
                tree.see(item)
                break
    
    def remove_model(self, tree, dialog):
        """Remove the selected model"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a model to remove.")
            return
            
        model_name = tree.item(selection[0], 'text')
        
        # Prevent removing all models
        if len(self.fitting_models) <= 1:
            messagebox.showwarning("Cannot Remove", "You must keep at least one model.")
            return
        
        # Confirm removal
        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove the model '{model_name}'?"):
            # User confirmed removal
            self.fitting_models.pop(model_name, None)
            tree.delete(selection[0])
            
            # Save to file
            self.save_fitting_models()
            
            # If current model was removed, select a different one
            if self.fitting_model_var.get() == model_name:
                self.fitting_model_var.set(list(self.fitting_models.keys())[0])
                self.update_model_parameters()
    
    def save_fitting_models(self):
        """Save the fitting models to a JSON file"""
        try:
            with open(self.models_file, 'w') as f:
                json.dump(self.fitting_models, f, indent=2)
        except Exception as e:
            print(f"Error saving fitting models: {str(e)}")
    
    def load_fitting_models(self):
        """Load fitting models from JSON file or return defaults if file doesn't exist"""
        try:
            if os.path.exists(self.models_file):
                with open(self.models_file, 'r') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Error loading fitting models: {str(e)}")
            return {}
    
    def restore_default_models(self, tree=None):
        """Restore the default models while preserving custom models"""
        if messagebox.askyesno("Restore Default Models", 
                             "This will restore all default models while preserving your custom models. Continue?"):
            # Get current custom models (any model not in default_fitting_models)
            custom_models = {name: info for name, info in self.fitting_models.items() 
                           if name not in self.default_fitting_models}
            
            # Start with default models
            self.fitting_models = self.default_fitting_models.copy()
            
            # Add back custom models
            self.fitting_models.update(custom_models)
            
            # Save to file
            self.save_fitting_models()
            
            # Update UI if tree is provided
            if tree is not None:
                for item in tree.get_children():
                    tree.delete(item)
                for model_name in sorted(self.fitting_models.keys()):
                    tree.insert('', 'end', text=model_name, values=(model_name,))
            
            # Update any open model UI
            if hasattr(self, 'model_dropdown'):
                self.model_dropdown['values'] = sorted(list(self.fitting_models.keys()))
                if self.fitting_model_var.get() not in self.fitting_models:
                    self.fitting_model_var.set(list(self.fitting_models.keys())[0])
                self.update_model_parameters()
            
            messagebox.showinfo("Success", "Default models have been restored while preserving custom models.")
    
    def save_model(self, name_var, desc_text, param_text, formula_text, listbox, dialog):
        """Save the current model details"""
        model_name = name_var.get().strip()
        if not model_name:
            messagebox.showwarning("Invalid Name", "Please enter a model name.")
            return
            
        # Parse parameters
        param_lines = param_text.get(1.0, tk.END).strip().split('\n')
        parameters = []
        for line in param_lines:
            if line.strip():
                parts = line.split(',', 1)
                if len(parts) != 2:
                    messagebox.showwarning("Invalid Parameter", 
                                         f"Parameter must be in format 'name, value': {line}")
                    return
                param_name = parts[0].strip()
                try:
                    param_value = float(parts[1].strip())
                    parameters.append((param_name, param_value))
                except ValueError:
                    messagebox.showwarning("Invalid Parameter Value", 
                                         f"Parameter value must be a number: {parts[1]}")
                    return
        
        formula = formula_text.get(1.0, tk.END).strip()
        if not formula:
            messagebox.showwarning("Invalid Formula", "Please enter a model formula.")
            return
            
        # Check if this is a renamed model
        old_name = None
        selection = listbox.curselection()
        if selection:
            old_name = listbox.get(selection[0])
            
        if old_name and old_name != model_name:
            # Handle renaming - remove old name
            self.fitting_models.pop(old_name, None)
            listbox.delete(selection[0])
        
        # Get description
        description = desc_text.get(1.0, tk.END).strip()
        
        # Save the model
        self.fitting_models[model_name] = {
            "parameters": parameters,
            "formula": formula,
            "description": description
        }
        
        # Save to file
        self.save_fitting_models()
        
        # Update the list and dropdown
        if old_name and old_name != model_name:
            listbox.insert(tk.END, model_name)
        elif not old_name:  # New model
            listbox.insert(tk.END, model_name)
            
        self.model_dropdown['values'] = sorted(list(self.fitting_models.keys()))
        
        # If currently selected model was renamed, update the variable
        if old_name and old_name == self.fitting_model_var.get():
            self.fitting_model_var.set(model_name)
            self.update_model_parameters()
            
        messagebox.showinfo("Model Saved", f"Model '{model_name}' has been saved.")
    
    def generate_model_function(self, model_name):
        """Generate a callable function based on the model formula"""
        model_info = self.fitting_models.get(model_name, {})
        formula = model_info.get("formula", "")
        
        # Extract the actual formula line (assuming it starts with 'y =')
        formula_lines = formula.split('\n')
        formula_line = ""
        for line in formula_lines:
            if line.strip().startswith('y ='):
                formula_line = line.strip()[4:].strip()  # Remove 'y = ' prefix
                break
        
        if not formula_line:
            return None
        
        # Get parameter names
        parameters = [p[0] for p in model_info.get("parameters", [])]
        param_str = ", ".join(parameters)
        
        # Create a function that will be used for curve_fit
        try:
            # Create the function definition
            func_code = f"def model_func(x, {param_str}):\n    return {formula_line}"
            
            # Create a local namespace
            namespace = {}
            
            # Add math functions to the namespace
            import math
            for name in dir(math):
                if name.startswith('__'):
                    continue
                namespace[name] = getattr(math, name)
                
            # Add numpy functions
            import numpy as np
            for name in dir(np):
                if name.startswith('__'):
                    continue
                namespace[name] = getattr(np, name)
            
            # Execute the function definition in the namespace
            exec(func_code, namespace)
            
            # Return the function
            return namespace['model_func']
        except Exception as e:
            print(f"Error generating model function: {e}")
            return None
    
    def setup_colors_tab(self):
        frame = self.colors_tab
        # --- Color management group ---
        color_mgmt_grp = ttk.LabelFrame(frame, text="Color Management", padding=6)
        color_mgmt_grp.pack(fill='x', padx=6, pady=(8,4))
        ttk.Button(color_mgmt_grp, text="Manage Colors & Palettes", command=self.manage_colors_palettes).pack(fill='x', pady=2)
        # --- Single color group ---
        single_grp = ttk.LabelFrame(frame, text="Single Data Color", padding=6)
        single_grp.pack(fill='x', padx=6, pady=4)
        # Don't reinitialize single_color_var which already exists and has been set with preferences
        tk.Label(single_grp, text="Single Data Color:").pack(anchor="w")
        self.single_color_dropdown = ttk.Combobox(single_grp, textvariable=self.single_color_var, values=list(self.custom_colors.keys()))
        self.single_color_dropdown.pack(fill='x', pady=2)
        # Add preview canvas for single color
        self.single_color_preview = tk.Canvas(single_grp, width=60, height=20, highlightthickness=0, bg='white')
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
        palette_grp = ttk.LabelFrame(frame, text="Group Palette", padding=6)
        palette_grp.pack(fill='x', padx=6, pady=4)
        # No need to re-initialize self.palette_var as it's already set in __init__
        tk.Label(palette_grp, text="Group Palette:").pack(anchor="w")
        self.palette_dropdown = ttk.Combobox(palette_grp, textvariable=self.palette_var, values=list(self.custom_palettes.keys()))
        self.palette_dropdown.pack(fill='x', pady=2)
        # Add preview canvas for palette
        self.palette_preview = tk.Canvas(palette_grp, height=20, highlightthickness=0, bg='white')
        self.palette_preview.pack(fill='x', pady=(0, 8))
        def update_palette_preview(event=None):
            self.palette_preview.delete('all')
            name = self.palette_var.get()
            colors = self.custom_palettes.get(name, [])
            for i, hexcode in enumerate(colors):
                x0 = 5 + i*14
                x1 = x0 + 12
                self.palette_preview.create_rectangle(x0, 2, x1, 18, fill=hexcode, outline='black')
        self.palette_dropdown.bind('<<ComboboxSelected>>', update_palette_preview)
        update_palette_preview()
        
        # --- Outline color group ---
        outline_grp = ttk.LabelFrame(frame, text="Outline Color", padding=6)
        outline_grp.pack(fill='x', padx=6, pady=4)
        tk.Label(outline_grp, text="Set the outline color for bar, box, and violin plots:").pack(anchor="w")
        
        # Radio buttons for outline color options
        outline_colors_frame = tk.Frame(outline_grp)
        outline_colors_frame.pack(fill='x', pady=4)
        
        ttk.Radiobutton(outline_colors_frame, text="Black", variable=self.outline_color_var, value="black").pack(anchor="w")
        ttk.Radiobutton(outline_colors_frame, text="Gray", variable=self.outline_color_var, value="gray").pack(anchor="w")
        ttk.Radiobutton(outline_colors_frame, text="As set", variable=self.outline_color_var, value="as_set").pack(anchor="w")
        ttk.Radiobutton(outline_colors_frame, text="White", variable=self.outline_color_var, value="white").pack(anchor="w")

    def update_color_palette_dropdowns(self):
        self.single_color_dropdown['values'] = list(self.custom_colors.keys())
        if self.single_color_var.get() not in self.custom_colors and self.custom_colors:
            self.single_color_var.set(list(self.custom_colors.keys())[0])
        self.palette_dropdown['values'] = list(self.custom_palettes.keys())
        if self.palette_var.get() not in self.custom_palettes and self.custom_palettes:
            self.palette_var.set(list(self.custom_palettes.keys())[0])
        
        # Also update settings dropdowns if they exist
        if hasattr(self, 'settings_single_color_dropdown'):
            self.settings_single_color_dropdown['values'] = list(self.custom_colors.keys())
        if hasattr(self, 'settings_palette_dropdown'):
            self.settings_palette_dropdown['values'] = list(self.custom_palettes.keys())

    def create_dropdown(self, parent, label_text, attr_name):
        tk.Label(parent, text=label_text).pack()
        var = tk.StringVar()
        dropdown = ttk.Combobox(parent, textvariable=var)
        dropdown.pack()
        setattr(self, f"{attr_name}_var", var)
        setattr(self, f"{attr_name}_dropdown", dropdown)

    def add_labeled_entry(self, parent, label):
        tk.Label(parent, text=label).pack()
        entry = ttk.Entry(parent)
        entry.pack()
        return entry

    def load_file(self, file_path=None):
        """Load an Excel file either from a provided path or by prompting the user to select one"""
        if file_path is None:
            file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
            
        if file_path:
            self.excel_file = file_path
            xls = pd.ExcelFile(self.excel_file)
            # Filter out sheets starting with underscore as they're meant for internal use
            visible_sheets = [sheet for sheet in xls.sheet_names if not str(sheet).startswith('_')]
            self.sheet_dropdown['values'] = visible_sheets

            # Select 'export' sheet if it's in the visible sheets, otherwise select first visible sheet
            if visible_sheets and "export" in visible_sheets:
                self.sheet_var.set("export")
            elif visible_sheets:
                self.sheet_var.set(visible_sheets[0])
            else:
                # If there are no visible sheets (unlikely, but possible), show a message
                messagebox.showinfo("No Visible Sheets", "This Excel file only contains hidden sheets (starting with '_').")

            self.load_sheet()

    def load_sheet(self, event=None):
        try:
            # Store previous X and Y axis labels to check if they were from previous sheet columns
            old_xlabel = self.xlabel_entry.get() if hasattr(self, 'xlabel_entry') else ''
            old_ylabel = self.ylabel_entry.get() if hasattr(self, 'ylabel_entry') else ''
            old_columns = list(self.df.columns) if self.df is not None else []
            
            selected_sheet = self.sheet_var.get()
            
            if selected_sheet == 'Raw Embedded Data':
                # Load the raw data
                if hasattr(self, 'raw_df') and self.raw_df is not None:
                    self.df = self.raw_df.copy()
                    
                    # Store the current modifications (to restore later)
                    self._stored_renames = self.xaxis_renames.copy() if hasattr(self, 'xaxis_renames') else {}
                    self._stored_excluded = self.excluded_x_values.copy() if hasattr(self, 'excluded_x_values') else set()
                    self._stored_order = self.xaxis_order.copy() if hasattr(self, 'xaxis_order') else []
                    
                    # When switching to raw data, temporarily clear all modifications
                    # so the plot and dialogs show the original data
                    self.xaxis_renames = {}
                    self.excluded_x_values = set()
                    self.xaxis_order = []
                    self.current_sheet_type = 'raw'
            elif selected_sheet == 'Modified Embedded Data':
                # Load the modified data
                if hasattr(self, 'modified_df') and self.modified_df is not None:
                    self.df = self.modified_df.copy()
                elif hasattr(self, 'df') and self.df is not None and hasattr(self, 'raw_df'):
                    # If we don't have a modified dataframe yet, create one from the current df
                    self.modified_df = self.df.copy()
                    # We're already using the current df, so no need to reassign
                
                # If we're switching back from raw data, restore the saved modifications
                if hasattr(self, '_stored_renames') and self.current_sheet_type == 'raw':
                    self.xaxis_renames = self._stored_renames.copy()
                    self.excluded_x_values = self._stored_excluded.copy()
                    self.xaxis_order = self._stored_order.copy()
                
                self.current_sheet_type = 'modified'
            elif selected_sheet not in ['Raw Embedded Data', 'Modified Embedded Data']:
                # Load from Excel file for normal sheets
                self.df = pd.read_excel(self.excel_file, sheet_name=selected_sheet, dtype=object)
                # When loading a new sheet, create a raw copy and initialize modifications
                self.raw_df = self.df.copy()
                self.modified_df = self.df.copy()
                self.xaxis_renames = {}
                self.excluded_x_values = set()
                self.xaxis_order = []
                self.current_sheet_type = 'external'
            
            # Flag to indicate if labels should be reset due to sheet change
            reset_labels = (old_xlabel in old_columns or not old_xlabel) and \
                           (old_ylabel in old_columns or not old_ylabel)
            
            self.update_columns(reset_labels=reset_labels)
            if not hasattr(self, 'xaxis_order'):
                self.xaxis_order = []
            
            # Initialize excluded values if not already done
            if not hasattr(self, 'excluded_x_values'):
                self.excluded_x_values = set()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet: {e}")

    def update_columns(self, reset_labels=False):
        columns = list(self.df.columns)
        
        # Update dropdown values
        self.xaxis_dropdown['values'] = columns
        self.group_dropdown['values'] = ['None'] + columns
        
        # Reset selections when switching sheets
        self.xaxis_var.set('')  # Clear X-axis selection
        self.group_var.set('None')  # Reset group selection to None
        
        # Reset axis labels if coming from a different sheet
        if reset_labels:
            if hasattr(self, 'xlabel_entry'):
                self.xlabel_entry.delete(0, tk.END)
            if hasattr(self, 'ylabel_entry'):
                self.ylabel_entry.delete(0, tk.END)
        
        # Clear Y-axis checkboxes
        for cb in self.value_checkbuttons:
            cb.destroy()
        self.value_vars.clear()
        
        # Recreate value column checkboxes
        for col in columns:
            var = tk.BooleanVar()
            var.trace_add('write', lambda *args, c=col: self.update_y_axis_label(c))
            cb = ttk.Checkbutton(self.value_vars_inner_frame, text=col, variable=var)
            cb.pack(anchor='w')
            self.value_vars.append((var, col))
            self.value_checkbuttons.append(cb)

    def update_x_axis_label(self, *args):
        """Update the X-axis label entry with the selected column name."""
        if hasattr(self, 'xlabel_entry') and self.xaxis_var.get():
            # Only update if the field is empty or matches a previous column name
            current_label = self.xlabel_entry.get()
            if not current_label or current_label in self.df.columns:
                self.xlabel_entry.delete(0, tk.END)
                self.xlabel_entry.insert(0, self.xaxis_var.get())
    
    def update_y_axis_label(self, column):
        """Update the Y-axis label entry with the selected Y column name."""
        if hasattr(self, 'ylabel_entry'):
            # Get the first selected Y column
            selected_y_cols = [col for var, col in self.value_vars if var.get()]
            if selected_y_cols:
                # Only update if the field is empty or matches a previous column name
                current_label = self.ylabel_entry.get()
                if not current_label or current_label in self.df.columns:
                    self.ylabel_entry.delete(0, tk.END)
                    self.ylabel_entry.insert(0, selected_y_cols[0])
    
    def modify_x_categories(self):
        """Opens a dialog to modify X-axis categories (rename and reorder)."""
        if self.df is None or not self.xaxis_var.get():
            messagebox.showerror("Error", "Load a file and select an X-axis first.")
            return
            
        # Check if we're working with raw data and give options to the user
        if hasattr(self, 'current_sheet_type') and self.current_sheet_type == 'raw':
            result = messagebox.askyesnocancel("Raw Data Selected", 
                "You are viewing raw data. What would you like to do?\n\n" +
                "Yes: Start over with fresh modifications based on raw data\n" +
                "No: Continue with existing modifications\n" +
                "Cancel: Don't modify data")
            
            if result is None:  # Cancel
                return
                
            if result:  # Yes - start over with raw data
                # Clear any existing modifications but keep working with raw data for now
                # We'll switch to modified view after making changes
                self.xaxis_renames = {}
                self.excluded_x_values = set()
                self.xaxis_order = []
                # We'll update modified_df after changes are made
            else:  # No - continue with existing modifications
                # Switch to modified data and restore existing modifications
                if hasattr(self, 'sheet_var') and 'Modified Embedded Data' in self.sheet_dropdown['values']:
                    self.sheet_var.set('Modified Embedded Data')
                    self.load_sheet()
        
        x_col = self.xaxis_var.get()
        
        # Get original values from dataframe
        all_df_values = list(pd.unique(self.df[x_col].dropna()))
        
        # If we don't have existing excluded values, create an empty set
        if not hasattr(self, 'excluded_x_values'):
            self.excluded_x_values = set()
        
        # Create the dialog window
        modify_window = tk.Toplevel(self.root)
        modify_window.title("Modify X Categories")
        modify_window.geometry("650x500")
        modify_window.resizable(True, True)
        
        # Add main frame with padding
        main_frame = tk.Frame(modify_window, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add explanation text
        explanation = "Modify X-axis categories: rename labels, reorder items, or exclude them from the plot."
        tk.Label(main_frame, text=explanation, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        
        # Create frame for the listbox and edit panel
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create a listbox with scrollbar (left side)
        listbox_frame = tk.Frame(list_frame)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create the listbox
        listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE, height=15, width=40)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Connect scrollbar to listbox
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        # Create edit panel (right side)
        edit_frame = tk.Frame(list_frame, padx=10)
        edit_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # Create button frame (for reordering buttons)
        button_frame = tk.Frame(edit_frame)
        button_frame.pack(side=tk.TOP, fill=tk.Y, pady=(0, 15))
        
        # Organize values: first display values in existing order, then any remaining values
        display_order = []
        display_values = {}
        original_values = []
        
        # Map displayed values (for renamed items)
        for val in all_df_values:
            display_val = self.xaxis_renames.get(val, val)
            display_values[val] = display_val
        
        # Start with values that are already ordered (if any)
        if hasattr(self, 'xaxis_order') and self.xaxis_order:
            # Include values that are in both the dataframe and our existing order
            for val in self.xaxis_order:
                # Map renamed values back to original values
                orig_val = None
                for k, v in self.xaxis_renames.items():
                    if v == val and k in all_df_values:
                        orig_val = k
                        break
                # If it's not a renamed value, check if it's directly in the dataframe
                if orig_val is None and val in all_df_values:
                    orig_val = val
                    
                if orig_val is not None and orig_val in all_df_values:
                    display_order.append(display_values.get(orig_val, str(orig_val)))
                    original_values.append(orig_val)
        
        # Add any remaining values from the dataframe that aren't already in display_order
        for val in all_df_values:
            display_val = display_values.get(val, str(val))
            if val not in [ov for ov in original_values]:
                display_order.append(display_val)
                original_values.append(val)
        
        # Populate the listbox
        for i, item in enumerate(display_order):
            # Mark excluded items with an asterisk
            if original_values[i] in self.excluded_x_values:
                listbox.insert(tk.END, f"* {item}")
            else:
                listbox.insert(tk.END, item)
        
        # Select the first item if available
        if listbox.size() > 0:
            listbox.selection_set(0)
        
        # Variables for the edit panel
        selected_index = tk.IntVar(value=-1)
        original_value_var = tk.StringVar()
        display_value_var = tk.StringVar()
        include_var = tk.BooleanVar(value=True)
        
        # Labels for the edit panel
        tk.Label(edit_frame, text="Original Value:").pack(anchor="w", pady=(0, 2))
        original_label = tk.Label(edit_frame, textvariable=original_value_var, wraplength=150)
        original_label.pack(anchor="w", pady=(0, 10))
        
        tk.Label(edit_frame, text="Display Value:").pack(anchor="w", pady=(0, 2))
        display_entry = tk.Entry(edit_frame, textvariable=display_value_var, width=20)
        display_entry.pack(anchor="w", pady=(0, 10))
        
        # Callback function to update when checkbox is changed
        def on_include_change():
            update_listbox_item()
            
        include_check = ttk.Checkbutton(edit_frame, text="Include in plot", variable=include_var, command=on_include_change)
        include_check.pack(anchor="w", pady=(0, 10))
        
        # Function to update the edit panel when a listbox item is selected
        def on_select(event):
            selection = listbox.curselection()
            if not selection:
                return
                
            idx = selection[0]
            selected_index.set(idx)
            
            # Get the original value and current display value
            orig_val = original_values[idx]
            original_value_var.set(str(orig_val))
            
            # Set the display value (either renamed or original)
            display_val = self.xaxis_renames.get(orig_val, str(orig_val))
            display_value_var.set(str(display_val))
            
            # Set the include checkbox
            include_var.set(orig_val not in self.excluded_x_values)
        
        # Function to update the listbox item when display value or inclusion changes
        def update_listbox_item():
            idx = selected_index.get()
            if idx < 0 or idx >= len(original_values):
                return
                
            # Get the original value
            orig_val = original_values[idx]
            
            # Update the display in the listbox
            display_text = display_value_var.get()
            if not include_var.get():
                listbox.delete(idx)
                listbox.insert(idx, f"* {display_text}")
            else:
                listbox.delete(idx)
                listbox.insert(idx, display_text)
                
            # Reselect the item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(idx)
            listbox.see(idx)
        
        # Bind the listbox selection event
        listbox.bind('<<ListboxSelect>>', on_select)
        
        # Apply button for the current edit
        def apply_current_edit():
            idx = selected_index.get()
            if idx < 0 or idx >= len(original_values):
                return
                
            update_listbox_item()
        
        ttk.Button(edit_frame, text="Apply Edit", command=apply_current_edit).pack(pady=(5, 15))
        
        # Add reordering buttons
        ttk.Label(button_frame, text="Reorder:").pack(anchor="w")
        ttk.Button(button_frame, text="Move Up", command=lambda: move_up()).pack(fill="x", pady=2)
        ttk.Button(button_frame, text="Move Down", command=lambda: move_down()).pack(fill="x", pady=2)
        
        # Function to move an item up in the list
        def move_up():
            selected_idx = listbox.curselection()
            if not selected_idx or selected_idx[0] == 0:
                return
            
            idx = selected_idx[0]
            item = listbox.get(idx)
            orig_val = original_values[idx]
            
            # Remove from current position
            listbox.delete(idx)
            original_values.pop(idx)
            
            # Insert at new position
            new_idx = idx - 1
            listbox.insert(new_idx, item)
            original_values.insert(new_idx, orig_val)
            
            # Select the moved item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(new_idx)
            listbox.see(new_idx)
        
        # Function to move an item down in the list
        def move_down():
            selected_idx = listbox.curselection()
            if not selected_idx or selected_idx[0] == listbox.size() - 1:
                return
            
            idx = selected_idx[0]
            item = listbox.get(idx)
            orig_val = original_values[idx]
            
            # Remove from current position
            listbox.delete(idx)
            original_values.pop(idx)
            
            # Insert at new position
            new_idx = idx + 1
            listbox.insert(new_idx, item)
            original_values.insert(new_idx, orig_val)
            
            # Select the moved item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(new_idx)
            listbox.see(new_idx)
        
        # Add control buttons at the bottom
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        
        def save_changes():
            # Clear existing settings
            self.xaxis_renames = {}
            self.excluded_x_values = set()
            self.xaxis_order = []
            
            # Process each item in the listbox
            for i in range(listbox.size()):
                # Get the display text (remove asterisk if present)
                display_text = str(listbox.get(i))
                is_excluded = display_text.startswith("* ")
                if is_excluded:
                    display_text = display_text[2:]  # Remove the "* " prefix
                
                # Get the original value
                orig_val = original_values[i]
                
                # Handle excluded values
                if is_excluded:
                    self.excluded_x_values.add(orig_val)
                    continue
                
                # Add to the order list
                self.xaxis_order.append(display_text)
                
                # Handle renamed values
                if str(display_text) != str(orig_val):
                    # Try to preserve numerical types
                    try:
                        if isinstance(orig_val, (int, float)) and display_text.replace('.', '', 1).isdigit():
                            if '.' in display_text:
                                display_text = float(display_text)
                            else:
                                display_text = int(display_text)
                    except (ValueError, AttributeError):
                        pass  # Keep as string if conversion fails
                    
                    self.xaxis_renames[orig_val] = display_text
            
            # Make sure we have both raw and modified data frames
            if not hasattr(self, 'raw_df') or self.raw_df is None:
                self.raw_df = self.df.copy()
                
            # Apply the changes to create the modified DataFrame
            x_col = self.xaxis_var.get()
            if x_col:
                # Determine which dataframe to start with based on current view
                if hasattr(self, 'current_sheet_type') and self.current_sheet_type == 'raw':
                    # If we're in raw data view and modifying, start with the raw data
                    base_df = self.raw_df.copy()
                else:
                    # Otherwise use the current df
                    base_df = self.df.copy()
                
                # Apply renames to this column
                if x_col in base_df.columns:
                    for orig_val, new_val in self.xaxis_renames.items():
                        # Replace the values in the dataframe
                        base_df.loc[base_df[x_col] == orig_val, x_col] = new_val
                    
                    # Store as the modified dataframe
                    self.modified_df = base_df
                    self.df = base_df.copy()
                    
            # Always switch to modified data view after saving changes
            if hasattr(self, 'sheet_var') and 'Modified Embedded Data' in self.sheet_dropdown['values']:
                self.sheet_var.set('Modified Embedded Data')
                self.current_sheet_type = 'modified'
            
            # Update the sheet options to include both data types if not already there
            if hasattr(self, 'sheet_dropdown'):
                if 'Raw Embedded Data' not in self.sheet_dropdown['values'] or \
                   'Modified Embedded Data' not in self.sheet_dropdown['values']:
                    # Add the embedded data options
                    current_options = list(self.sheet_dropdown['values'])
                    if 'Embedded Data' in current_options:
                        current_options.remove('Embedded Data')
                    if 'Raw Embedded Data' not in current_options:
                        current_options.append('Raw Embedded Data')
                    if 'Modified Embedded Data' not in current_options:
                        current_options.append('Modified Embedded Data')
                    self.sheet_dropdown['values'] = current_options
                    # Set to Modified Embedded Data
                    self.sheet_var.set('Modified Embedded Data')
            
            # Record that we're currently working with modified data
            self.current_sheet_type = 'modified'
            
            modify_window.destroy()
        
        def cancel():
            modify_window.destroy()
        
        # Add save and cancel buttons
        ttk.Button(control_frame, text="Cancel", command=cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="Apply Changes", command=save_changes).pack(side=tk.RIGHT, padx=5)

    def rename_x_labels(self):
        """Opens a dialog to rename X-axis labels using a listbox interface."""
        if self.df is None or not self.xaxis_var.get():
            messagebox.showerror("Error", "Load a file and select an X-axis first.")
            return
        
        x_col = self.xaxis_var.get()
        
        # Get original values from dataframe
        all_df_values = list(pd.unique(self.df[x_col].dropna()))
        
        # If we don't have existing excluded values, create an empty set
        if not hasattr(self, 'excluded_x_values'):
            self.excluded_x_values = set()
        
        # Create the dialog window
        rename_window = tk.Toplevel(self.root)
        rename_window.title("Rename X Labels")
        rename_window.geometry("500x450")
        rename_window.resizable(True, True)
        
        # Add main frame with padding
        main_frame = ttk.Frame(rename_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add explanation text
        explanation = "Select an item to edit its display name or to exclude it from the plot."
        ttk.Label(main_frame, text=explanation, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        
        # Create frame for the listbox and edit panel
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create a listbox with scrollbar (left side)
        listbox_frame = tk.Frame(list_frame)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create the listbox
        listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE, height=15, width=40)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Connect scrollbar to listbox
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        # Create edit panel (right side)
        edit_frame = tk.Frame(list_frame, padx=10)
        edit_frame.pack(side=tk.LEFT, fill=tk.Y)
        
        # Organize values: first display values in existing order, then any remaining values
        display_order = []
        display_values = {}
        original_values = []
        
        # Map displayed values (for renamed items)
        for val in all_df_values:
            display_val = self.xaxis_renames.get(val, val)
            display_values[val] = display_val
        
        # Start with values that are already ordered (if any)
        if hasattr(self, 'xaxis_order') and self.xaxis_order:
            # Include values that are in both the dataframe and our existing order
            for val in self.xaxis_order:
                # Map renamed values back to original values
                orig_val = None
                for k, v in self.xaxis_renames.items():
                    if v == val and k in all_df_values:
                        orig_val = k
                        break
                # If it's not a renamed value, check if it's directly in the dataframe
                if orig_val is None and val in all_df_values:
                    orig_val = val
                    
                if orig_val is not None and orig_val in all_df_values:
                    display_order.append(display_values.get(orig_val, str(orig_val)))
                    original_values.append(orig_val)
        
        # Add any remaining values from the dataframe that aren't already in display_order
        for val in all_df_values:
            display_val = display_values.get(val, str(val))
            if val not in [ov for ov in original_values]:
                display_order.append(display_val)
                original_values.append(val)
        
        # Populate the listbox
        for i, item in enumerate(display_order):
            # Mark excluded items with an asterisk
            if original_values[i] in self.excluded_x_values:
                listbox.insert(tk.END, f"* {item}")
            else:
                listbox.insert(tk.END, item)
        
        # Select the first item if available
        if listbox.size() > 0:
            listbox.selection_set(0)
        
        # Variables for the edit panel
        selected_index = tk.IntVar(value=-1)
        original_value_var = tk.StringVar()
        display_value_var = tk.StringVar()
        include_var = tk.BooleanVar(value=True)
        
        # Labels for the edit panel
        tk.Label(edit_frame, text="Original Value:").pack(anchor="w", pady=(0, 2))
        original_label = tk.Label(edit_frame, textvariable=original_value_var, wraplength=150)
        original_label.pack(anchor="w", pady=(0, 10))
        
        tk.Label(edit_frame, text="Display Value:").pack(anchor="w", pady=(0, 2))
        display_entry = tk.Entry(edit_frame, textvariable=display_value_var, width=20)
        display_entry.pack(anchor="w", pady=(0, 10))
        
        include_check = ttk.Checkbutton(edit_frame, text="Include in plot", variable=include_var, command=update_listbox_item)
        include_check.pack(anchor="w", pady=(0, 10))
        
        # Function to update the edit panel when a listbox item is selected
        def on_select(event):
            selection = listbox.curselection()
            if not selection:
                return
                
            idx = selection[0]
            selected_index.set(idx)
            
            # Get the original value and current display value
            orig_val = original_values[idx]
            original_value_var.set(str(orig_val))
            
            # Set the display value (either renamed or original)
            display_val = self.xaxis_renames.get(orig_val, str(orig_val))
            display_value_var.set(str(display_val))
            
            # Set the include checkbox
            include_var.set(orig_val not in self.excluded_x_values)
        
        # Function to update the listbox item when display value or inclusion changes
        def update_listbox_item():
            idx = selected_index.get()
            if idx < 0 or idx >= len(original_values):
                return
                
            # Get the original value
            orig_val = original_values[idx]
            
            # Update the display in the listbox
            display_text = display_value_var.get()
            if not include_var.get():
                listbox.delete(idx)
                listbox.insert(idx, f"* {display_text}")
            else:
                listbox.delete(idx)
                listbox.insert(idx, display_text)
                
            # Reselect the item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(idx)
            listbox.see(idx)
        
        # Bind the listbox selection event
        listbox.bind('<<ListboxSelect>>', on_select)
        
        # Apply button for the current edit
        def apply_current_edit():
            idx = selected_index.get()
            if idx < 0 or idx >= len(original_values):
                return
                
            update_listbox_item()
        
        ttk.Button(edit_frame, text="Apply Edit", command=apply_current_edit).pack(pady=(10, 0))
        
        # Add control buttons at the bottom
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        
        def save_renames():
            # Clear existing settings
            self.xaxis_renames = {}
            self.excluded_x_values = set()
            self.xaxis_order = []
            
            # Process each item in the listbox
            for i in range(listbox.size()):
                # Get the display text (remove asterisk if present)
                display_text = str(listbox.get(i))
                is_excluded = display_text.startswith("* ")
                if is_excluded:
                    display_text = display_text[2:]  # Remove the "* " prefix
                
                # Get the original value
                orig_val = original_values[i]
                
                # Handle excluded values
                if is_excluded:
                    self.excluded_x_values.add(orig_val)
                    continue
                
                # Add to the order list
                self.xaxis_order.append(display_text)
                
                # Handle renamed values
                if str(display_text) != str(orig_val):
                    # Try to preserve numerical types
                    try:
                        if isinstance(orig_val, (int, float)) and display_text.replace('.', '', 1).isdigit():
                            if '.' in display_text:
                                display_text = float(display_text)
                            else:
                                display_text = int(display_text)
                    except (ValueError, AttributeError):
                        pass  # Keep as string if conversion fails
                    
                    self.xaxis_renames[orig_val] = display_text
            
            rename_window.destroy()
        
        def cancel():
            rename_window.destroy()
        
        # Add save and cancel buttons
        tk.Button(control_frame, text="Cancel", command=cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="Apply Changes", command=save_renames).pack(side=tk.RIGHT, padx=5)
    
    def reorder_x_categories(self):
        """Opens a dialog to reorder X-axis categories using a listbox for more reliable reordering."""
        if self.df is None or not self.xaxis_var.get():
            messagebox.showerror("Error", "Load a file and select an X-axis first.")
            return
        
        x_col = self.xaxis_var.get()
        
        # Get original values from dataframe
        all_df_values = list(pd.unique(self.df[x_col].dropna()))
        
        # Create the dialog window
        reorder_window = tk.Toplevel(self.root)
        reorder_window.title("Reorder X Categories")
        reorder_window.geometry("400x450")
        reorder_window.resizable(True, True)
        
        # Add main frame with padding
        main_frame = tk.Frame(reorder_window, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add explanation text
        explanation = "Select an item and use the buttons to move it up or down in the list."
        tk.Label(main_frame, text=explanation, justify=tk.LEFT).pack(anchor="w", pady=(0, 10))
        
        # Create frame for the listbox and buttons
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create a listbox with scrollbar
        listbox_frame = tk.Frame(list_frame)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create the listbox
        listbox = tk.Listbox(listbox_frame, selectmode=tk.SINGLE, height=15, width=40)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Connect scrollbar to listbox
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        # Create button frame
        button_frame = tk.Frame(list_frame)
        button_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        # Organize values: first display values in existing order, then any remaining values
        display_order = []
        display_values = {}
        original_values = []
        
        # If we don't have existing excluded values, create an empty set
        if not hasattr(self, 'excluded_x_values'):
            self.excluded_x_values = set()
            
        # Map displayed values (for renamed items)
        for val in all_df_values:
            if val in self.excluded_x_values:
                continue
            display_val = self.xaxis_renames.get(val, val)
            display_values[val] = display_val
        
        # Start with values that are already ordered (if any)
        if hasattr(self, 'xaxis_order') and self.xaxis_order:
            # Include values that are in both the dataframe and our existing order
            for val in self.xaxis_order:
                # Map renamed values back to original values
                orig_val = None
                for k, v in self.xaxis_renames.items():
                    if v == val and k in all_df_values:
                        orig_val = k
                        break
                # If it's not a renamed value, check if it's directly in the dataframe
                if orig_val is None and val in all_df_values:
                    orig_val = val
                    
                if orig_val is not None and orig_val in all_df_values and orig_val not in self.excluded_x_values:
                    display_order.append(display_values.get(orig_val, str(orig_val)))
                    original_values.append(orig_val)
        
        # Add any remaining values from the dataframe that aren't already in display_order
        for val in all_df_values:
            display_val = display_values.get(val, str(val))
            if display_val not in display_order and val not in self.excluded_x_values:
                display_order.append(display_val)
                original_values.append(val)
        
        # Populate the listbox
        for item in display_order:
            listbox.insert(tk.END, item)
        
        # Select the first item if available
        if listbox.size() > 0:
            listbox.selection_set(0)
        
        # Function to move an item up in the list
        def move_up():
            selected_idx = listbox.curselection()
            if not selected_idx or selected_idx[0] == 0:
                return
            
            idx = selected_idx[0]
            item = listbox.get(idx)
            orig_val = original_values[idx]
            
            # Remove from current position
            listbox.delete(idx)
            original_values.pop(idx)
            
            # Insert at new position
            new_idx = idx - 1
            listbox.insert(new_idx, item)
            original_values.insert(new_idx, orig_val)
            
            # Select the moved item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(new_idx)
            listbox.see(new_idx)
        
        # Function to move an item down in the list
        def move_down():
            selected_idx = listbox.curselection()
            if not selected_idx or selected_idx[0] == listbox.size() - 1:
                return
            
            idx = selected_idx[0]
            item = listbox.get(idx)
            orig_val = original_values[idx]
            
            # Remove from current position
            listbox.delete(idx)
            original_values.pop(idx)
            
            # Insert at new position
            new_idx = idx + 1
            listbox.insert(new_idx, item)
            original_values.insert(new_idx, orig_val)
            
            # Select the moved item
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(new_idx)
            listbox.see(new_idx)
        
        # Function to delete an item from the list
        def delete_item():
            selected_idx = listbox.curselection()
            if not selected_idx:
                return
                
            idx = selected_idx[0]
            item = listbox.get(idx)
            
            # Ask for confirmation before deleting
            confirm = messagebox.askyesno(
                "Confirm Delete", 
                f"Are you sure you want to delete '{item}' from the list?"
            )
            
            if not confirm:
                return
                
            # Remove from listbox and original_values list
            listbox.delete(idx)
            original_values.pop(idx)
            
            # Select the next item if available, or the last item
            if listbox.size() > 0:
                new_idx = min(idx, listbox.size() - 1)
                listbox.selection_set(new_idx)
                listbox.see(new_idx)
        
        # Add buttons to move and delete items
        ttk.Button(button_frame, text="Move Up", command=move_up, width=10).pack(pady=5)
        ttk.Button(button_frame, text="Move Down", command=move_down, width=10).pack(pady=5)
        ttk.Button(button_frame, text="Delete", command=delete_item, width=10).pack(pady=5)
        
        # Add control buttons at the bottom
        control_frame = tk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(10, 0))
        
        def save_order():
            # Clear existing order
            self.xaxis_order = []
            
            # Create new order based on current listbox order
            for i in range(listbox.size()):
                display_val = listbox.get(i)
                self.xaxis_order.append(display_val)
            
            reorder_window.destroy()
        
        def cancel():
            reorder_window.destroy()
        
        # Add save and cancel buttons
        tk.Button(control_frame, text="Cancel", command=cancel).pack(side=tk.RIGHT, padx=5)
        ttk.Button(control_frame, text="Apply Order", command=save_order).pack(side=tk.RIGHT, padx=5)
    
    # The modify_dataframe method has been replaced by rename_x_labels and reorder_x_categories methods

    def plot_graph(self):
        """Generate a plot using the current data and settings.
        This method now integrates statistics calculation directly with plotting.
        If 'Use statistics' is enabled, statistics will be calculated and annotated on the plot
        based on the 'Show statistical annotations on plot' setting.
        """
        try:
            if hasattr(self, 'fig') and self.fig is not None:
                plt.close(self.fig)
                self.fig = None
        except Exception:
            pass

        # Robustly initialize show_errorbars at the very top
        show_errorbars = getattr(self, 'show_errorbars_var', None)
        if show_errorbars is not None:
            show_errorbars = show_errorbars.get()
        else:
            show_errorbars = True
            
        # If model fitting is enabled, ensure we're using XY plot type
        if hasattr(self, 'use_fitting_var') and self.use_fitting_var.get():
            current_plot_kind = self.plot_kind_var.get()
            if current_plot_kind != "xy":
                self.plot_kind_var.set("xy")

        # Clear any existing statistics when generating a new plot
        self.latest_pvals = {}
        
        # Prepare the working dataframe
        df_work, x_col, value_cols, group_col = self.prepare_working_dataframe()
        
        if df_work is None:
            return
            
        # Store current plot information for statistics generation
        self.current_plot_info = {
            'x_col': x_col,
            'value_cols': value_cols,
            'group_col': group_col
        }

        try:
            linewidth = float(self.linewidth.get())
        except Exception:
            linewidth = 1.0
            
        plot_mode = 'single' if len(value_cols) == 1 else 'overlay'
        plot_width = self.plot_width_var.get()
        plot_height = self.plot_height_var.get()
        fontsize = int(self.fontsize_entry.get())
        n_rows = 1  # No longer supporting split Y-axis

        # Get plot type early for margin calculations
        plot_kind = self.plot_kind_var.get()  # "bar", "box", "violin", or "xy"

        swap_axes_enabled = self.swap_axes_var.get() if plot_kind in ["bar", "box", "violin"] else False
        
        # Store the original X column for reference
        original_x_col = x_col
        if '_renamed_x' in df_work.columns:
            original_x_col = self.xaxis_var.get()
            
        # Handle categorical vs numeric X values
        # For non-swapped categorical plots, create a numeric mapping for consistent spacing.
        # For swapped plots, keep categories as strings (seaborn expects categorical y; numeric y breaks badly).
        if plot_kind in ["bar", "box", "violin"] and not swap_axes_enabled:
            unique_vals = df_work[x_col].unique()
            unique_vals = [val for val in unique_vals if not pd.isna(val)]

            if self.xaxis_order:
                x_values = [val for val in self.xaxis_order if val in unique_vals]
                x_values.extend([val for val in unique_vals if val not in self.xaxis_order])
            else:
                x_values = sorted(unique_vals, key=lambda x: str(x))

            self.x_categorical_map = {val: i for i, val in enumerate(x_values)}
            self.x_categorical_reverse_map = {i: val for val, i in self.x_categorical_map.items()}
            df_work['_x_plot'] = df_work[x_col].map(self.x_categorical_map)
            x_col = '_x_plot'
            self.df = df_work
        else:
            if hasattr(self, 'x_categorical_map'):
                try:
                    delattr(self, 'x_categorical_map')
                except Exception:
                    pass
            if hasattr(self, 'x_categorical_reverse_map'):
                try:
                    delattr(self, 'x_categorical_reverse_map')
                except Exception:
                    pass
            if '_x_plot' in df_work.columns:
                try:
                    df_work = df_work.drop('_x_plot', axis=1)
                except Exception:
                    pass
            if plot_kind in ["bar", "box", "violin"] and swap_axes_enabled and self.xaxis_order:
                try:
                    df_work[x_col] = pd.Categorical(df_work[x_col], categories=self.xaxis_order, ordered=True)
                except Exception:
                    pass
            self.df = df_work
        
        # Scale margins based on plot size - smaller plots need relatively larger margins
        plot_height_val = self.plot_height_var.get()  # User-specified plot height
        
        # Base margins - keep these modest, tight_layout handles the rest for export
        left_margin = 1.0
        right_margin = 0.5
        top_margin = 0.8  # Base top margin
        bottom_margin = 1.5  # Bottom margin for rotated x-tick labels and axis title
        
        # Add modest extra margin for statistics annotations if using statistics
        if group_col and self.use_stats_var.get():
            top_margin += 0.4

        fig_width = plot_width + left_margin + right_margin
        fig_height = plot_height * n_rows + top_margin + bottom_margin
        
        # Configure legend placement strategy based on plot size
        if plot_height_val <= 2.0 and plot_kind == "xy" and self.use_fitting_var.get():
            # For small XY plots with fitting, we'll place legends outside the plot area
            self.legend_outside = True
            # Add more space on the right for external legend
            fig_width += 2.0
        else:
            # Default placement within the plot
            self.legend_outside = False
            
        # Get legend settings from UI
        legend_visible = self.legend_visible_var.get()
        legend_position = self.legend_position_var.get()
        legend_ncol = self.legend_ncol_var.get()
        
        # Add extra figure space for outside legend positions
        if legend_visible:
            if legend_position == "outside right":
                fig_width += 2.0  # Extra width for legend on the right
            elif legend_position == "outside top":
                fig_height += 1.0  # Extra height for legend on top
        
        # Define a utility function for consistent legend placement with user settings
        def place_legend(ax, handles, labels, draggable=True):
            """Place legend based on user settings and make it draggable."""
            # Check if legend should be shown
            if not legend_visible:
                return None
            
            # Determine number of columns (0 = auto)
            if legend_ncol > 0:
                ncol = legend_ncol
            elif legend_position == "outside right":
                # For outside right, default to 1 column (vertical layout) for better fit
                ncol = 1
            else:
                ncol = self.optimize_legend_layout(ax, handles, labels, fontsize=fontsize)
            
            # Map position string to matplotlib loc and bbox_to_anchor
            if legend_position == "outside right":
                legend = ax.legend(handles, labels, bbox_to_anchor=(1.02, 1), loc='upper left', 
                                  borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)
            elif legend_position == "outside top":
                legend = ax.legend(handles, labels, bbox_to_anchor=(0.5, 1.15), loc='upper center', 
                                  borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)
            elif self.legend_outside:
                # Override for small XY plots with fitting
                legend = ax.legend(handles, labels, bbox_to_anchor=(1.05, 1), loc='upper left', 
                                  borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)
            else:
                # Standard matplotlib positions
                legend = ax.legend(handles, labels, loc=legend_position, frameon=False, 
                                  fontsize=fontsize, ncol=ncol)
            
            # Make legend draggable so user can fine-tune position
            if legend and draggable:
                legend.set_draggable(True)
            
            return legend
                
        # Add the function as an attribute of the class instance
        self.place_legend = place_legend

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
        # Store original figure size for PDF export
        self.original_fig_width = fig_width
        self.original_fig_height = fig_height
        
        # Preview scale controls DPI (pixel output) not figure size (layout)
        # 100% = 100 DPI, 175% = 175 DPI, etc.
        preview_dpi = self.preview_dpi.get()
        
        # Check if Y-axis break is enabled
        use_yaxis_break = self.yaxis_break_var.get() if hasattr(self, 'yaxis_break_var') else False
        ybreak_min = None
        ybreak_max = None
        ybreak_ratio = None
        ybreak_gap = None

        if swap_axes_enabled and use_yaxis_break:
            use_yaxis_break = False
            try:
                messagebox.showwarning("Y-axis break disabled", "Y-axis break is not supported when 'Swap X and Y axes' is enabled. Disable swap to use Y-axis break.")
            except Exception:
                pass
        
        self.debug(f"Y-axis break checkbox: {use_yaxis_break}")
        
        if use_yaxis_break:
            try:
                ybreak_min_str = self.ybreak_min_entry.get().strip()
                ybreak_max_str = self.ybreak_max_entry.get().strip()
                ybreak_ratio_str = self.ybreak_ratio_entry.get().strip()
                ybreak_gap_str = self.ybreak_gap_entry.get().strip() if hasattr(self, 'ybreak_gap_entry') else ''
                self.debug(f"Break values from UI: min='{ybreak_min_str}', max='{ybreak_max_str}', ratio='{ybreak_ratio_str}', gap='{ybreak_gap_str}'")
                
                if not ybreak_min_str or not ybreak_max_str:
                    raise ValueError("Break from/to values are required")
                
                ybreak_min = float(ybreak_min_str)
                ybreak_max = float(ybreak_max_str)
                
                # Handle ratio
                if ybreak_ratio_str:
                    # Handle case where user enters a range like "0.2-0.8" - take first number
                    if '-' in ybreak_ratio_str and not ybreak_ratio_str.startswith('-'):
                        ybreak_ratio_str = ybreak_ratio_str.split('-')[0]
                    ybreak_ratio = float(ybreak_ratio_str)
                    ybreak_ratio = max(0.1, min(0.9, ybreak_ratio))  # Clamp to valid range

                if ybreak_gap_str:
                    if '-' in ybreak_gap_str and not ybreak_gap_str.startswith('-'):
                        ybreak_gap_str = ybreak_gap_str.split('-')[0]
                    ybreak_gap = float(ybreak_gap_str)
                    ybreak_gap = max(0.0, min(0.3, ybreak_gap))
                
                
                if ybreak_min >= ybreak_max:
                    use_yaxis_break = False
                    print("Warning: Break min must be less than break max. Axis break disabled.")
            except (ValueError, AttributeError) as e:
                use_yaxis_break = False
                print(f"Warning: Invalid axis break values (min='{ybreak_min_str}', max='{ybreak_max_str}', ratio='{ybreak_ratio_str}'): {e}. Axis break disabled.")
        
        # Create figure - with or without axis break
        if use_yaxis_break and n_rows == 1:
            # Calculate height ratios from DATA RANGES before creating figure
            # This ensures consistent figure sizing
            try:
                # Get user-specified y-limits if any
                user_ymin = float(self.ymin_entry.get()) if self.ymin_entry.get().strip() else 0
                user_ymax = float(self.ymax_entry.get()) if self.ymax_entry.get().strip() else None
            except ValueError:
                user_ymin = 0
                user_ymax = None
            
            # Calculate data max from dataframe using value_cols (already defined above)
            data_max = max(self.df[col].max() for col in value_cols) * 1.05
            upper_ymax = user_ymax if user_ymax is not None else data_max
            
            # Calculate height ratios based on actual data ranges
            lower_range = ybreak_min - user_ymin  # e.g., 50 - 0 = 50
            upper_range = upper_ymax - ybreak_max  # e.g., 200 - 70 = 130
            total_range = lower_range + upper_range
            
            if total_range > 0:
                calc_lower_ratio = lower_range / total_range
                calc_upper_ratio = upper_range / total_range
            else:
                calc_lower_ratio = 0.5
                calc_upper_ratio = 0.5
            
            
            # First create a reference 1-subplot figure to get its default margins
            ref_fig, ref_ax = plt.subplots(1, 1, figsize=(fig_width, fig_height), dpi=preview_dpi)
            ref_params = ref_fig.subplotpars
            ref_left, ref_right = ref_params.left, ref_params.right
            ref_bottom, ref_top = ref_params.bottom, ref_params.top

            # Store the exact axes position of the reference plot.
            # We'll reuse this bbox for the broken-axis layout so the graph area matches exactly.
            self._ybreak_ref_ax_pos = ref_ax.get_position().frozen()
            plt.close(ref_fig)
            
            # Create 2-subplot figure with SAME margins as 1-subplot figure
            self.ybreak_ratio = ybreak_ratio
            if ybreak_ratio is not None:
                height_ratios = [ybreak_ratio, 1.0 - ybreak_ratio]
            else:
                height_ratios = [calc_upper_ratio, calc_lower_ratio]
            self.fig, (ax_upper, ax_lower) = plt.subplots(
                2, 1, figsize=(fig_width, fig_height), dpi=preview_dpi,
                gridspec_kw={'height_ratios': height_ratios, 'hspace': 0,
                            'left': ref_left, 'right': ref_right, 'bottom': ref_bottom, 'top': ref_top},
                sharex=True
            )
            axes = [ax_lower]  # Use lower axis as the "main" axis for compatibility
            self.ax_upper = ax_upper
            self.ax_lower = ax_lower
            self.ybreak_enabled = True
            self.ybreak_min = ybreak_min
            self.ybreak_max = ybreak_max
            self.ybreak_gap = ybreak_gap
        else:
            # Standard figure creation
            self.fig, axes = plt.subplots(n_rows, 1, figsize=(fig_width, fig_height), dpi=preview_dpi, squeeze=False)
            axes = axes.flatten()
            self.ybreak_enabled = False
            self.ax_upper = None
            self.ax_lower = None
            self.ybreak_gap = None

        show_frame = self.show_frame_var.get()
        show_hgrid = self.show_hgrid_var.get()
        show_vgrid = self.show_vgrid_var.get()
        # Only apply swap_axes to supported plot types (bar, box, violin)
        plot_kind = self.plot_kind_var.get()  # "bar", "box", "violin", or "xy"
        swap_axes = swap_axes_enabled
        strip_black = self.strip_black_var.get()
        show_stripplot = self.show_stripplot_var.get()
        errorbar_black = self.errorbar_black_var.get()

        disable_plot_stat_annotations = False

        # Default initialization of show_mean
        show_mean = False
        if plot_kind == 'xy':
            show_mean = self.xy_show_mean_var.get()

        for idx, value_col in enumerate(value_cols):
            ax = axes[idx] if n_rows > 1 else axes[0]
            # Use df_work which has the _renamed_x column if it exists
            df_plot = df_work.copy()

            if plot_mode == 'overlay' and len(value_cols) > 1:
                df_plot = pd.melt(df_plot, id_vars=[x_col] + ([group_col] if group_col else []),
                                   value_vars=value_cols, var_name='Measurement', value_name='_plotted_value')
                value_col = '_plotted_value'  # Use this new column for plotting
                hue_col = 'Measurement'
            else:
                # Value column is already set to the selected column
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
                # Always set the appropriate CI value based on error bar type
                if errorbar_type == "SD":
                    ci_val = 'sd'
                    estimator = np.mean
                else: # SEM
                    ci_val = 'sd'  # Will handle SEM manually
                    estimator = np.mean
                    sem_mode = True
                
                # Handle palette consistently regardless of error bar type
                if hue_col and hue_col in df_plot.columns:
                    # For grouped data, use the palette
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    hue_groups = df_plot[hue_col].dropna().unique()
                    if len(palette_full) < len(hue_groups):
                        # Repeat the palette to ensure we have enough colors
                        palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                    palette = palette_full[:len(hue_groups)]
                else:
                    # For ungrouped data, use the single color
                    single_color_name = self.single_color_var.get()
                    single_color = self.custom_colors.get(single_color_name, 'black')
                    palette = [single_color]
                # --- Always use full palette for grouped XY means ---
                if show_mean:
                    groupers = [x_col]
                    if hue_col:
                        groupers.append(hue_col)
                    grouped = df_plot.groupby(groupers)[value_col]
                    means = grouped.mean().reset_index()
                    if self.errorbar_type_var.get() == "SEM":
                        try:
                            errors = grouped.apply(lambda x: np.std(x.dropna().astype(float), ddof=1) / np.sqrt(len(x.dropna())) if len(x.dropna()) > 1 else 0).reset_index(name='err')
                        except Exception as e:
                            print(f"Error calculating SEM: {e}")
                            errors = grouped.std(ddof=1).reset_index(name='err')

            # --- Swap axes logic ---
            if swap_axes:
                plot_args = dict(
                    data=df_plot, y=x_col, x=value_col, hue=hue_col, ax=ax,
                )
                if plot_kind == "bar":
                    plot_dict = dict(ci=ci_val, capsize=0.2, palette=palette, errcolor='black', errwidth=linewidth, estimator=estimator)
                    # Add edge color if bar outline is enabled
                    if self.bar_outline_var.get():
                        # Get the appropriate outline color based on setting
                        if hue_col and hue_col in df_plot.columns:  # Use palette colors for grouped data
                            outline_color = self.get_outline_color(None)  # We'll handle colors per group
                        else:  # Use single color for ungrouped data
                            single_color_name = self.single_color_var.get()
                            single_color = self.custom_colors.get(single_color_name, 'black')
                            outline_color = self.get_outline_color(single_color)
                        
                        # Ensure linewidth is at least 0.5 for visibility
                        effective_linewidth = max(linewidth, 0.5)
                        plot_dict.update(dict(edgecolor=outline_color, linewidth=effective_linewidth))
                    plot_args.update(plot_dict)
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
                        grouped = df_plot.groupby(groupers)[value_col]
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
                                y = group[value_col]
                                yerr = group['err']
                                ecolor = 'black' if errorbar_black else c
                                mfc = c if filled else 'none'
                                mec = c
                                if show_mean_errorbars:
                                    ax.errorbar(x, y, yerr=yerr, fmt=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, capsize=3, ecolor=ecolor, label=str(name))
                                ax.plot(x, y, marker=marker_symbol, color=c, markerfacecolor=mfc, markeredgecolor=mec, markersize=marker_size, linewidth=linewidth, linestyle='None', label=None if show_mean_errorbars else str(name))
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
                            y = means[value_col]
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
                            handles, labels = ax.get_legend_handles_labels()
                        if handles and len(handles) > 0:
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
                                # Always show marker edge in group color if not filled
                                edge = c
                                face = c if filled else 'none'
                                scatter = ax.scatter(group[x_col], group[value_col], marker=marker_symbol, s=marker_size**2, color=c, label=str(name), edgecolors=edge, facecolors=face, linewidth=linewidth)
                                if draw_band:
                                    x_sorted = np.sort(group[x_col].unique())
                                    min_vals = [group[group[x_col] == x][value_col].min() for x in x_sorted]
                                    max_vals = [group[group[x_col] == x][value_col].max() for x in x_sorted]
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
                                            means = [group[group[x_col] == x][value_col].mean() for x in x_sorted]
                                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                            means_numeric = pd.to_numeric(means, errors='coerce')
                                            ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                                    else:
                                        c = palette[0]
                                        x_sorted = np.sort(df_plot[x_col].unique())
                                        means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                        means_numeric = pd.to_numeric(means, errors='coerce')
                                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                            handles, labels = ax.get_legend_handles_labels()
                        if handles and len(handles) > 0:
                            ax.legend()
                        else:
                            c = palette[0]
                            # Always show marker edge in group color if not filled
                            edge = c
                            face = c if filled else 'none'
                            ax.scatter(df_plot[x_col], df_plot[value_col], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=edge, facecolors=face, linewidth=linewidth)
                            if draw_band:
                                x_sorted = np.sort(df_plot[x_col].unique())
                                min_vals = [df_plot[df_plot[x_col] == x][value_col].min() for x in x_sorted]
                                max_vals = [df_plot[df_plot[x_col] == x][value_col].max() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                                max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                                ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                            if connect:
                                # Connect means of raw data at each x value
                                c = palette[0]
                                x_sorted = np.sort(df_plot[x_col].unique())
                                means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                                x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                means_numeric = pd.to_numeric(means, errors='coerce')
                                ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                stripplot_args = dict(
                    data=df_plot, y=x_col, x=value_col, hue=hue_col, dodge=True,
                    jitter=True, marker='o', alpha=0.55,
                    ax=ax
                )
            else:
                plot_args = dict(
                    data=df_plot, x=x_col, y=value_col, hue=hue_col, ax=ax,
                )
                # Set default estimator
                estimator = np.mean  # Default to mean if not specified
                

                show_stripplot = self.show_stripplot_var.get()  # Use the actual UI setting
                
                # --- Bar plot: always horizontal bars if swap_axes, for both categorical and numerical x ---
                if plot_kind == "bar":
                    # Build base arguments common to both orientations
                    base_args = {
                        'errorbar': 'sd', 
                        'capsize': 0.2, 
                        'palette': palette, 
                        'err_kws': {'color': 'black', 'linewidth': linewidth}, 
                        'estimator': estimator
                    }
                    
                    # Add edge color if bar outline is enabled
                    if self.bar_outline_var.get():
                        # Get the appropriate outline color based on setting
                        if hue_col and hue_col in df_plot.columns:  # Use palette colors for grouped data
                            outline_color = self.get_outline_color(None)  # We'll handle colors per group
                        else:  # Use single color for ungrouped data
                            single_color_name = self.single_color_var.get()
                            single_color = self.custom_colors.get(single_color_name, 'black')
                            outline_color = self.get_outline_color(single_color)
                        
                        # Ensure linewidth is at least 0.5 for visibility
                        effective_linewidth = max(linewidth, 0.5)
                        base_args.update({'edgecolor': outline_color, 'linewidth': effective_linewidth})
                    
                    if swap_axes:
                        plot_args = dict(
                            data=df_plot, y=x_col, x=value_col, hue=hue_col, ax=ax,
                            **base_args
                        )
                    else:
                        plot_args = dict(
                            data=df_plot, x=x_col, y=value_col, hue=hue_col, ax=ax,
                            **base_args
                        )
                elif plot_kind == "box":
                    plot_args.update(dict(palette=palette, linewidth=linewidth, showcaps=True, boxprops=dict(linewidth=linewidth), medianprops=dict(linewidth=linewidth), dodge=True, width=0.7))
                elif plot_kind == "violin":
                    # Create a fresh set of arguments specifically for violin plots
                    # Don't modify plot_args which might be used by other plot types
                    violin_specific_args = dict(
                        inner="box",  # Show box plot inside violin
                        scale="width",  # Scale violins to have the same width
                    )
                    
                    # Set up violin plot specific arguments without affecting other plot types
                    plot_args.update(dict(
                        palette=palette, 
                        linewidth=linewidth, 
                        dodge=True, 
                        width=0.8,
                        **violin_specific_args  # Add violin-specific arguments separately
                    ))
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
                        grouped = df_plot.groupby(groupers)[value_col]
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
                                y = group[value_col]
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
                        # When show_mean is False, plot raw data points
                        c = palette[0]
                        # Always show marker edge in group color if not filled
                        edge = c
                        face = c if filled else 'none'
                        ax.scatter(df_plot[x_col], df_plot[value_col], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=edge, facecolors=face, linewidth=linewidth)
                        if draw_band:
                            x_sorted = np.sort(df_plot[x_col].unique())
                            min_vals = [df_plot[df_plot[x_col] == x][value_col].min() for x in x_sorted]
                            max_vals = [df_plot[df_plot[x_col] == x][value_col].max() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                            max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                            ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                        if connect:
                            # Connect means of raw data at each x value
                            c = palette[0]
                            x_sorted = np.sort(df_plot[x_col].unique())
                            y_means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            y_means_numeric = pd.to_numeric(y_means, errors='coerce')
                            ax.plot(x_sorted_numeric, y_means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        
                    if hue_col:
                        handles, labels = ax.get_legend_handles_labels()
                        
                        if handles and len(handles) > 0:
                            # Use our utility function for consistent legend placement
                            self.place_legend(ax, handles, labels)
                
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
                            # Always show marker edge in group color if not filled
                            edge = c
                            face = c if filled else 'none'
                            scatter = ax.scatter(group[x_col], group['Value'], marker=marker_symbol, s=marker_size**2, color=c, label=str(name), edgecolors=edge, facecolors=face, linewidth=linewidth)
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
                                        means = [group[group[x_col] == x][value_col].mean() for x in x_sorted]
                                        x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                        means_numeric = pd.to_numeric(means, errors='coerce')
                                        ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                                else:
                                    c = palette[0]
                                    x_sorted = np.sort(df_plot[x_col].unique())
                                    means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                                    x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                                    means_numeric = pd.to_numeric(means, errors='coerce')
                                    ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                        handles, labels = ax.get_legend_handles_labels()
                        if handles and len(handles) > 0:
                            ax.legend()
                    else:
                        c = palette[0]
                        # Always show marker edge in group color if not filled
                        edge = c
                        face = c if filled else 'none'
                        ax.scatter(df_plot[x_col], df_plot[value_col], marker=marker_symbol, s=marker_size**2, color=c, edgecolors=edge, facecolors=face, linewidth=linewidth)
                        if draw_band:
                            x_sorted = np.sort(df_plot[x_col].unique())
                            min_vals = [df_plot[df_plot[x_col] == x][value_col].min() for x in x_sorted]
                            max_vals = [df_plot[df_plot[x_col] == x][value_col].max() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            min_vals_numeric = pd.to_numeric(min_vals, errors='coerce')
                            max_vals_numeric = pd.to_numeric(max_vals, errors='coerce')
                            ax.fill_between(x_sorted_numeric, min_vals_numeric, max_vals_numeric, color=c, alpha=0.18, zorder=1)
                        if connect:
                            # Connect means of raw data at each x value
                            c = palette[0]
                            x_sorted = np.sort(df_plot[x_col].unique())
                            means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                            x_sorted_numeric = pd.to_numeric(x_sorted, errors='coerce')
                            means_numeric = pd.to_numeric(means, errors='coerce')
                            ax.plot(x_sorted_numeric, means_numeric, color='black' if line_black else c, linewidth=linewidth, alpha=0.7, linestyle=line_style)
                stripplot_args = dict(
                    data=df_plot, x=x_col, y=value_col, hue=hue_col, dodge=True,
                    jitter=True, marker='o', alpha=0.55,
                    ax=ax
                )

            # --- Plotting ---
            # First prepare the palette (common to all plot types)
            if hue_col and hue_col in df_plot.columns:
                hue_groups = df_plot[hue_col].dropna().unique()
                palette_name = self.palette_var.get()
                palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                if len(palette_full) < len(hue_groups):
                    palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                palette = palette_full[:len(hue_groups)]
            else:
                # Use single color for non-grouped data
                single_color_name = self.single_color_var.get()
                palette = [self.custom_colors.get(single_color_name, 'black')]
            
            # Create a separate section for each plot type with its own parameters
            # This prevents parameter leakage between plot types
            if plot_kind == "bar":
                # Start with fresh parameters for bar plots
                bar_args = {}
                
                # Data parameters
                bar_args['data'] = df_plot
                bar_args['x'] = x_col if not swap_axes else value_col
                bar_args['y'] = value_col if not swap_axes else x_col
                bar_args['hue'] = hue_col
                bar_args['ax'] = ax
                bar_args['palette'] = palette
                
                # Bar plot specific settings
                errorbar_type = self.errorbar_type_var.get()
                capsize_option = self.errorbar_capsize_var.get()
                
                # Map text options to numeric capsize values (proportional to typical bar width)
                capsize_values = {
                    "Default": 0.4,  # Approximately half of typical bar width
                    "Narrow": 0.2,  # Narrower than default
                    "Wide": 0.8,    # Wider than default
                    "Wider": 1.2,   # Much wider
                    "None": 0.0     # No capsize
                }
                
                # Get numeric capsize value from the selected option
                capsize = capsize_values.get(capsize_option, 0.4)  # Default to 0.4 if option not found
                
                if errorbar_type == 'SD':
                    bar_args['errorbar'] = 'sd'  # or ci='sd' for older Seaborn
                else:
                    bar_args['errorbar'] = 'se'  # Using native standard error parameter
                
                # Set errorbar styling - capsize needs to be passed directly to barplot
                # and line width needs to be in err_kws
                linewidth = self.linewidth.get()
                
                # Modern way to handle errorbar styling in Seaborn
                # Handle error bars in a way that's compatible with newer seaborn versions
                # Prepare error bar styling
                err_kws = {}
                
                # Set up error bar styling specifically for bar plots
                err_kws = {}
                
                # We're using isolated parameters, so no need to clean up old values

                # Get the linewidth setting
                linewidth = self.linewidth.get()

                # Determine error bar type for Seaborn
                try:
                    if hasattr(self, 'errorbar_type_var'):
                        errorbar_type = self.errorbar_type_var.get().lower()
                    elif hasattr(self, 'settings_errorbar_type_var'):
                        errorbar_type = self.settings_errorbar_type_var.get().lower()
                    else:
                        errorbar_type = 'sd'
                except Exception:
                    errorbar_type = 'sd'

                # Map to Seaborn errorbar parameter
                if errorbar_type == 'sem':
                    # Standard error of the mean
                    bar_args['errorbar'] = 'se'
                else:
                    # Standard deviation
                    bar_args['errorbar'] = 'sd'

                # Get error bar styling preferences
                try:
                    # Check appropriate variable for black errorbars
                    if hasattr(self, 'errorbar_black_var'):
                        black_errorbars = bool(self.errorbar_black_var.get())
                    elif hasattr(self, 'settings_errorbar_black_var'):
                        black_errorbars = bool(self.settings_errorbar_black_var.get())
                    else:
                        black_errorbars = bool(self.black_errorbars_var.get())
                except Exception:
                    black_errorbars = False

                # Apply error bar color based on settings
                if black_errorbars:
                    # Use black when the option is checked
                    err_kws['color'] = 'black'
                else:
                    # Use plot's original color
                    if palette:
                        # Use first color from palette
                        err_kws['color'] = palette[0]
                    elif 'color' in bar_args:
                        # Explicit color specified in bar args
                        err_kws['color'] = bar_args['color']
                    else:
                        # Use single color from custom colors
                        color = self.custom_colors.get(self.single_color_var.get(), 'black')
                        err_kws['color'] = color

                # Always set linewidth
                err_kws['linewidth'] = linewidth

                # Since we're using isolated parameters, we don't need to clean up old values
                
                # Handle capsize based on errorbar_capsize_var
                capsize_val = 0  # Default to no caps
                if hasattr(self, 'errorbar_capsize_var'):
                    capsize_setting = self.errorbar_capsize_var.get()
                    
                    # Determine bar width (default to 0.8 if not specified)
                    bar_width = bar_args.get('width', 0.8)
                    
                    # Calculate capsize proportional to bar width
                    if capsize_setting == 'Default':
                        capsize_val = bar_width * 0.5  # Moderate caps
                    elif capsize_setting == 'Narrow':
                        capsize_val = bar_width * 0.2  # Narrow caps
                    elif capsize_setting == 'Wide':
                        capsize_val = bar_width * 0.7  # Wide caps
                    elif capsize_setting == 'Wider':
                        capsize_val = bar_width  # Very wide caps
                    elif capsize_setting == 'None':
                        capsize_val = 0  # No caps
                
                # Add capsize to bar_args
                bar_args['capsize'] = capsize_val
                
                # Add error bar styling if we have color settings
                if err_kws:
                    bar_args['err_kws'] = err_kws
                
                # Add yerr to bar_args if provided
                if 'yerr' in bar_args:
                    # Ensure only upward error bars
                    yerr_data = bar_args['yerr']
                    bar_args['yerr'] = [[0] * len(yerr_data), yerr_data]
                
                # Remove parameters that cause issues in newer seaborn versions
                # Keep 'errorbar' since we need it for SD/SEM, but remove others
                for param in ['errwidth', 'elinewidth', 'capthick']:
                    if param in bar_args:
                        bar_args.pop(param)
                
                # Get desired z-order for error bars based on upward-only setting
                # z-order concept: error bars (behind=5, front=15) | bars (10) | stripplot (15) | axes (20)
                
                # Check if upward-only error bars are enabled
                upward_only = self.upward_errorbar_var.get() if hasattr(self, 'upward_errorbar_var') else False
                
                # Modify err_kws to include appropriate z-order
                if 'err_kws' not in bar_args:
                    bar_args['err_kws'] = {}
                    
                if upward_only:
                    # For upward-only: error bars should be BEHIND the bars
                    bar_args['err_kws']['zorder'] = 5  # Lower z-order than bars
                else:
                    # For bidirectional: error bars should be IN FRONT of bars
                    bar_args['err_kws']['zorder'] = 15  # Higher z-order than bars
                
                # Set z-order for bars (always the same)
                bar_args['zorder'] = 10
                
                # Create barplot with completely isolated parameters
                ax = sns.barplot(**bar_args)
                
                # If axis break is enabled, also plot on upper axis with same parameters
                if self.ybreak_enabled and self.ax_upper is not None:
                    bar_args_upper = bar_args.copy()
                    bar_args_upper['ax'] = self.ax_upper
                    sns.barplot(**bar_args_upper)
                    # Remove legend from upper (will be handled in apply_yaxis_break)
                    if self.ax_upper.get_legend():
                        self.ax_upper.get_legend().remove()
                
                # Remove duplicate legend entries for all bar graphs
                if plot_kind == "bar":
                    # Get the current handles and labels
                    handles, labels = ax.get_legend_handles_labels()
                    
                    # Only proceed if we have handles
                    if handles and len(handles) > 0:
                        # Create a dictionary to store unique labels and their handles
                        unique_labels = {}
                        for i, label in enumerate(labels):
                            if label not in unique_labels:
                                unique_labels[label] = handles[i]
                        
                        # Remove the existing legend
                        if ax.get_legend():
                            ax.get_legend().remove()
                        
                        # Create a new legend with unique entries only if we have labels
                        if unique_labels:
                            unique_handles = list(unique_labels.values())
                            unique_label_texts = list(unique_labels.keys())
                            ax.legend(unique_handles, unique_label_texts)
                
                # Adjust bar widths and positions to create gaps between bars within groups
                if plot_kind == "bar" and hue_col and hue_col in df_plot.columns:
                    # Get the number of hue groups
                    n_groups = len(df_plot[hue_col].unique())
                    # Get the bar gap multiplier setting
                    bar_gap_multiplier = self.bar_gap_multiplier_var.get()
                    # Calculate the width for each bar (smaller for more groups)
                    bar_width = 0.8 / n_groups * bar_gap_multiplier  # Adjust bar width based on the multiplier
                    
                    # Adjust each bar's width to create gaps
                    for bar in ax.patches:
                        if swap_axes:
                            current_height = bar.get_height()
                            current_y = bar.get_y()
                            center = current_y + current_height / 2
                            bar.set_height(bar_width)
                            bar.set_y(center - bar_width / 2)
                        else:
                            current_width = bar.get_width()
                            current_x = bar.get_x()
                            center = current_x + current_width / 2
                            bar.set_width(bar_width)
                            bar.set_x(center - bar_width / 2)
                    
                    # Apply same bar width adjustment to upper axis when axis break is enabled
                    if self.ybreak_enabled and self.ax_upper is not None:
                        for bar in self.ax_upper.patches:
                            if swap_axes:
                                current_height = bar.get_height()
                                current_y = bar.get_y()
                                center = current_y + current_height / 2
                                bar.set_height(bar_width)
                                bar.set_y(center - bar_width / 2)
                            else:
                                current_width = bar.get_width()
                                current_x = bar.get_x()
                                center = current_x + current_width / 2
                                bar.set_width(bar_width)
                                bar.set_x(center - bar_width / 2)
                
                # Fix axis element visibility by bringing them to the front
                for spine in ax.spines.values():
                    spine.set_zorder(20)  # Highest z-order for axis elements
                ax.xaxis.set_zorder(20)
                ax.yaxis.set_zorder(20)
                
                # Handle outline colors after bar adjustments
                if hasattr(self, 'bar_outline_var'):
                    if not self.bar_outline_var.get():
                        # When outlines are disabled, set edgecolor to match facecolor
                        for bar in ax.patches:
                            bar.set_edgecolor(bar.get_facecolor())
                    else:
                        # When outlines are enabled, apply the outline color setting
                        # The width adjustments above might have reset the edgecolor, so we need to set it again
                        if hue_col and hue_col in df_plot.columns:
                            # For grouped data, each group needs its own color or the setting's color
                            outline_color = self.get_outline_color(None)
                            if self.outline_color_var.get() == "as_set":
                                # For "As set", make each bar's outline match its own color
                                for bar in ax.patches:
                                    bar.set_edgecolor(bar.get_facecolor())
                                    bar.set_linewidth(max(linewidth, 0.5))
                            else:
                                # Apply the same outline color to all bars
                                for bar in ax.patches:
                                    bar.set_edgecolor(outline_color)
                                    # Ensure the linewidth is visible
                                    bar.set_linewidth(max(linewidth, 0.5))
                        else:
                            # For ungrouped data, use the single color setting
                            single_color_name = self.single_color_var.get()
                            single_color = self.custom_colors.get(single_color_name, 'black')
                            outline_color = self.get_outline_color(single_color)
                            for bar in ax.patches:
                                if self.outline_color_var.get() == "as_set":
                                    # Explicitly set the edgecolor to match the facecolor
                                    bar.set_edgecolor(bar.get_facecolor())
                                    bar.set_linewidth(max(linewidth, 0.5))
                                else:
                                    # Use the explicit outline color
                                    bar.set_edgecolor(outline_color)
                                    bar.set_linewidth(max(linewidth, 0.5))
                
                # Apply same outline settings to upper axis bars when axis break is enabled
                if self.ybreak_enabled and hasattr(self, 'ax_upper') and self.ax_upper is not None:
                    if not self.bar_outline_var.get():
                        for bar in self.ax_upper.patches:
                            bar.set_edgecolor(bar.get_facecolor())
                    else:
                        if hue_col and hue_col in df_plot.columns:
                            outline_color = self.get_outline_color(None)
                            if self.outline_color_var.get() == "as_set":
                                for bar in self.ax_upper.patches:
                                    bar.set_edgecolor(bar.get_facecolor())
                                    bar.set_linewidth(max(linewidth, 0.5))
                            else:
                                for bar in self.ax_upper.patches:
                                    bar.set_edgecolor(outline_color)
                                    bar.set_linewidth(max(linewidth, 0.5))
                        else:
                            single_color_name = self.single_color_var.get()
                            single_color = self.custom_colors.get(single_color_name, 'black')
                            outline_color = self.get_outline_color(single_color)
                            for bar in self.ax_upper.patches:
                                if self.outline_color_var.get() == "as_set":
                                    bar.set_edgecolor(bar.get_facecolor())
                                    bar.set_linewidth(max(linewidth, 0.5))
                                else:
                                    bar.set_edgecolor(outline_color)
                                    bar.set_linewidth(max(linewidth, 0.5))
            elif plot_kind == "box":
                # For ungrouped data, we need to adjust the boxplot parameters
                # to ensure proper centering over X-values
                if 'hue' in plot_args and plot_args['hue'] is None:
                    # Make a copy of plot_args to avoid modifying the original
                    box_args = plot_args.copy()
                    
                    # Remove dodge parameter for ungrouped data
                    if 'dodge' in box_args:
                        box_args.pop('dodge')
                    
                    # Adjust width to ensure proper centering
                    box_args['width'] = 0.5
                    
                    # Get outline color based on setting
                    single_color_name = self.single_color_var.get()
                    single_color = self.custom_colors.get(single_color_name, 'black')
                    outline_color = self.get_outline_color(single_color)
                    
                    # Set outline properties with the determined color
                    box_args['boxprops'] = {'edgecolor': outline_color}
                    box_args['whiskerprops'] = {'color': outline_color}
                    box_args['medianprops'] = {'color': outline_color}
                    box_args['capprops'] = {'color': outline_color}
                    
                    sns.boxplot(**box_args)

                    # If axis break is enabled, also plot on upper axis with same parameters
                    if self.ybreak_enabled and self.ax_upper is not None:
                        box_args_upper = box_args.copy()
                        box_args_upper['ax'] = self.ax_upper
                        sns.boxplot(**box_args_upper)
                        # Remove legend from upper (will be handled in apply_yaxis_break)
                        if self.ax_upper.get_legend():
                            self.ax_upper.get_legend().remove()
                else:
                    # Adjust parameters for grouped data to add spacing
                    if 'hue' in plot_args and plot_args['hue'] is not None:
                        # For grouped box plots
                        plot_args['width'] = 0.6  # Make boxes narrower than default
                        plot_args['dodge'] = True  # Enable dodging for groups
                        plot_args['gap'] = 0.2    # Add gap between boxes in the same group
                    
                    # For grouped data, determine outline color
                    if self.outline_color_var.get() == "as_set":
                        # Will use palette colors for each box
                        # Don't set explicit colors here as Seaborn will handle them
                        plot_args['boxprops'] = {'linewidth': max(linewidth, 0.5)}
                        plot_args['whiskerprops'] = {'linewidth': max(linewidth, 0.5)}
                        plot_args['medianprops'] = {'linewidth': max(linewidth, 0.5)}
                        plot_args['capprops'] = {'linewidth': max(linewidth, 0.5)}
                    else:
                        # Use explicit color from setting
                        outline_color = self.get_outline_color(None)  # No default color provided
                        plot_args['boxprops'] = {'edgecolor': outline_color, 'linewidth': max(linewidth, 0.5)}
                        plot_args['whiskerprops'] = {'color': outline_color, 'linewidth': max(linewidth, 0.5)}
                        plot_args['medianprops'] = {'color': outline_color, 'linewidth': max(linewidth, 0.5)}
                        plot_args['capprops'] = {'color': outline_color, 'linewidth': max(linewidth, 0.5)}
                    
                    sns.boxplot(**plot_args)

                    # If axis break is enabled, also plot on upper axis with same parameters
                    if self.ybreak_enabled and self.ax_upper is not None:
                        plot_args_upper = plot_args.copy()
                        plot_args_upper['ax'] = self.ax_upper
                        sns.boxplot(**plot_args_upper)
                        # Remove legend from upper (will be handled in apply_yaxis_break)
                        if self.ax_upper.get_legend():
                            self.ax_upper.get_legend().remove()
                    
                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)
                
            elif plot_kind == "violin":
                # For violin plots, create a completely separate args dictionary
                # with no shared parameters with other plot types
                violin_args = {}
                
                # Basic data parameters
                violin_args['data'] = df_plot
                violin_args['x'] = x_col if not swap_axes else value_col
                violin_args['y'] = value_col if not swap_axes else x_col
                violin_args['hue'] = hue_col
                violin_args['ax'] = ax
                
                # Styling parameters
                violin_args['palette'] = palette
                violin_args['linewidth'] = max(linewidth, 0.5)  # Ensure minimum visible linewidth
                
                # Determine outline color based on settings
                if self.outline_color_var.get() == "as_set":
                    if hue_col and hue_col in df_plot.columns:
                        # For grouped violin plots with 'as_set', each violin will use its own color
                        # Don't set explicit edgecolor here as Seaborn will use the palette colors
                        pass
                    else:
                        # For single color violin plot with 'as_set'
                        single_color_name = self.single_color_var.get()
                        single_color = self.custom_colors.get(single_color_name, 'black')
                        violin_args['edgecolor'] = single_color
                else:
                    # Use the explicit color from settings
                    outline_color = self.get_outline_color(None)
                    violin_args['edgecolor'] = outline_color
                
                # For ungrouped data, we need different width settings
                if hue_col is None or hue_col not in df_plot.columns:
                    # Adjust width for ungrouped data
                    violin_args['width'] = 0.7
                else:
                    # For grouped violin plots
                    violin_args['width'] = 0.8
                    violin_args['dodge'] = True
                    violin_args['gap'] = 0.15    # Add gap between violins in the same group
                
                # Add violin-specific parameters based on user preference
                if self.violin_inner_box_var.get():
                    violin_args['inner'] = "box"  # Shows a mini boxplot inside each violin
                else:
                    violin_args['inner'] = "stick"  # Just show the quartiles
                
                violin_args['scale'] = "width"  # Makes all violins have the same width
                
                # Create the violin plot with completely isolated parameters
                sns.violinplot(**violin_args)

                # If axis break is enabled, also plot on upper axis with same parameters
                if self.ybreak_enabled and self.ax_upper is not None:
                    violin_args_upper = violin_args.copy()
                    violin_args_upper['ax'] = self.ax_upper
                    sns.violinplot(**violin_args_upper)
                    # Remove legend from upper (will be handled in apply_yaxis_break)
                    if self.ax_upper.get_legend():
                        self.ax_upper.get_legend().remove()
                
                # Ensure axis ticks are properly formatted
                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)
                
                # Stripplots will be handled by the global show_stripplot functionality
            elif plot_kind == "xy":
                # For XY plots, always use original X values (not categorical)
                if hasattr(self, 'x_categorical_map'):
                    delattr(self, 'x_categorical_map')
                    if '_x_plot' in self.df.columns:
                        self.df = self.df.drop('_x_plot', axis=1)

                if len(value_cols) == 1:
                    color = self.custom_colors.get(self.single_color_var.get(), 'black')
                    palette = [color]
                else:
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    if hue_col and hue_col in df_plot.columns:
                        hue_groups = df_plot[hue_col].dropna().unique()
                        if len(palette_full) < len(hue_groups):
                            palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                        palette = palette_full[:len(hue_groups)]
                    else:
                        palette = palette_full[:len(value_cols)]

                self._plot_xy_base(ax, df_plot, x_col, value_col, hue_col, value_cols, errorbar_black, linewidth, allow_legend=True)
                self._plot_xy_fitting(ax, df_plot, x_col, value_col, hue_col, palette, linewidth, update_results=True, allow_legend=True)

                if self.ybreak_enabled and self.ax_upper is not None:
                    self._plot_xy_base(self.ax_upper, df_plot, x_col, value_col, hue_col, value_cols, errorbar_black, linewidth, allow_legend=False)
                    self._plot_xy_fitting(self.ax_upper, df_plot, x_col, value_col, hue_col, palette, linewidth, update_results=False, allow_legend=False)
                    if self.ax_upper.get_legend():
                        self.ax_upper.get_legend().remove()

                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)

            # --- Stripplot (if enabled and not XY plot) ---
            if show_stripplot and plot_kind != "xy":
                # Rebuild stripplot args here so swap_axes is always applied consistently
                if swap_axes:
                    stripplot_args = dict(
                        data=df_plot, y=x_col, x=value_col, hue=hue_col, dodge=True,
                        jitter=True, marker='o', alpha=0.55,
                        ax=ax
                    )
                else:
                    stripplot_args = dict(
                        data=df_plot, x=x_col, y=value_col, hue=hue_col, dodge=True,
                        jitter=True, marker='o', alpha=0.55,
                        ax=ax
                    )

                # First, prepare all the parameters for a single stripplot call
                
                # Check if "Show stripplot with black dots" option is selected
                strip_black = self.strip_black_var.get()
                
                if strip_black:
                    # If black dots option is selected, use black color
                    stripplot_args["color"] = "black"
                    # Remove any palette parameter if it exists
                    if "palette" in stripplot_args:
                        del stripplot_args["palette"]
                else:
                    # Use palette colors
                    if hue_col:
                        # Make sure stripplot uses the same palette as the barplot
                        hue_groups = df_plot[hue_col].dropna().unique()
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        # Ensure we have enough colors for all hue groups
                        if len(palette_full) < len(hue_groups):
                            palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                        stripplot_args["palette"] = palette_full[:len(hue_groups)]
                    else:
                        # For no hue column, use the first color from palette
                        stripplot_args["palette"] = palette
                
                # Set specific parameters for bar plots without hue
                if plot_kind == 'bar' and not hue_col:
                    # Reduce jitter for more precise positioning
                    stripplot_args['jitter'] = 0.2
                    stripplot_args['dodge'] = False  # Prevent automatic dodging
                
                # Set z-order higher than bars (10) to ensure stripplot points are visible
                stripplot_args['zorder'] = 15
                
                # Suppress legend entries for stripplot (don't show duplicate labels)
                stripplot_args['legend'] = False
                
                # Make a single call to stripplot with all parameters set properly
                sns.stripplot(**stripplot_args)
                
                # If axis break is enabled, also plot stripplot on upper axis
                if self.ybreak_enabled and self.ax_upper is not None:
                    stripplot_args_upper = stripplot_args.copy()
                    stripplot_args_upper['ax'] = self.ax_upper
                    sns.stripplot(**stripplot_args_upper)

            # --- Always rebuild legend after all plotting ---
            if hue_col and (plot_kind == "box" or plot_kind == "violin"):
                import matplotlib.patches as mpatches
                hue_levels = list(df_plot[hue_col].dropna().unique())
                palette_name = self.palette_var.get()
                palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                if len(palette_full) < len(hue_levels):
                    palette_full = (palette_full * ((len(hue_levels) // len(palette_full)) + 1))[:len(hue_levels)]
                handles = [mpatches.Patch(facecolor=palette_full[i], edgecolor='black', label=str(hue_levels[i])) for i in range(len(hue_levels))]
                ax.legend(
                    handles,
                    [str(l) for l in hue_levels],
                    loc="upper center", bbox_to_anchor=(0.5, 1.18), borderaxespad=0,
                    frameon=False, fontsize=fontsize,
                    ncol=self.optimize_legend_layout(ax, handles, [str(l) for l in hue_levels], fontsize=fontsize)
                )
            elif hue_col and plot_kind == "bar":
                # Get all handles and labels
                handles, labels = ax.get_legend_handles_labels()
                
                # For multiple y-axis columns, we need to handle the legend differently
                if len(value_cols) > 1 and hue_col == 'Measurement':
                    # Create a dictionary to store unique labels and their handles
                    unique_labels = {}
                    unique_handles = []
                    
                    # Get the first occurrence of each label
                    for h, label in zip(handles, labels):
                        if label not in unique_labels:
                            unique_labels[label] = h
                    
                    # Create the legend with unique entries
                    if unique_labels:
                        ax.legend(
                            unique_labels.values(),
                            unique_labels.keys(),
                            loc="upper center", 
                            bbox_to_anchor=(0.5, 1.18), 
                            borderaxespad=0,
                            frameon=False, 
                            fontsize=fontsize,
                            ncol=self.optimize_legend_layout(ax, list(unique_labels.values()), list(unique_labels.keys()), fontsize=fontsize)
                        )
                else:
                    # Original behavior for single y-axis column
                    from matplotlib.patches import Rectangle
                    bar_handles = [h for h in handles if isinstance(h, Rectangle) and h.get_height() != 0]
                    bar_labels = [l for h, l in zip(handles, labels) if isinstance(h, Rectangle) and h.get_height() != 0]
                    if not bar_handles:  # fallback: use all handles
                        bar_handles, bar_labels = handles, labels
                    ax.legend(
                        bar_handles,
                        bar_labels,
                        loc="upper center", 
                        bbox_to_anchor=(0.5, 1.18), 
                        borderaxespad=0,
                        frameon=False, 
                        fontsize=fontsize,
                        ncol=self.optimize_legend_layout(ax, bar_handles, bar_labels, fontsize=fontsize)
                    )
            elif hue_col and plot_kind == "xy":
                # Build legend with marker symbols using palette colors directly
                from matplotlib.lines import Line2D
                
                # Get unique group names and their colors from custom palette (same as XY plot uses)
                group_names = list(df_plot[hue_col].dropna().unique())
                palette_name = self.palette_var.get()
                palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                if len(palette_full) < len(group_names):
                    palette_full = (palette_full * ((len(group_names) // len(palette_full)) + 1))[:len(group_names)]
                
                marker_symbol = self.xy_marker_symbol_var.get()
                marker_size = self.xy_marker_size_var.get()
                filled = self.xy_filled_var.get()
                connect = self.xy_connect_var.get()
                line_style = self.xy_line_style_var.get()
                line_black = self.xy_line_black_var.get()
                
                # Create legend handles with markers (and lines if connecting)
                legend_handles = []
                legend_labels = []
                for i, name in enumerate(group_names):
                    color = palette_full[i]
                    line_color = 'black' if line_black else color
                    
                    if connect:
                        # Show both marker and line
                        marker_handle = Line2D([0], [0], marker=marker_symbol,
                                              color=line_color,
                                              markerfacecolor=color if filled else 'none',
                                              markeredgecolor=color,
                                              markersize=marker_size * 0.8,
                                              markeredgewidth=linewidth,
                                              linestyle=line_style,
                                              linewidth=linewidth)
                    else:
                        # Show only marker
                        marker_handle = Line2D([0], [0], marker=marker_symbol, color='none',
                                              markerfacecolor=color if filled else 'none',
                                              markeredgecolor=color,
                                              markersize=marker_size * 0.8,
                                              markeredgewidth=linewidth,
                                              linestyle='none')
                    legend_handles.append(marker_handle)
                    legend_labels.append(str(name))
                
                if legend_handles:
                    ax.legend(
                        legend_handles, legend_labels,
                        loc="upper center", bbox_to_anchor=(0.5, 1.18), borderaxespad=0,
                        frameon=False, fontsize=fontsize,
                        ncol=self.optimize_legend_layout(ax, legend_handles, legend_labels, fontsize=fontsize)
                    )

            # Set categorical tick labels for bar, box, and violin plots using categorical mapping
            if plot_kind in ["bar", "box", "violin"] and hasattr(self, 'x_categorical_reverse_map') and not swap_axes:
                # Extract labels from our categorical mapping
                x_tick_locs = sorted(self.x_categorical_reverse_map.keys())
                x_tick_labels = [self.x_categorical_reverse_map[i] for i in x_tick_locs]
                if swap_axes:
                    ax.set_yticks(x_tick_locs)
                    rotation_angle = 0
                    ax.set_yticklabels(x_tick_labels, rotation=rotation_angle, fontsize=fontsize)
                else:
                    ax.set_xticks(x_tick_locs)
                    # Set appropriate rotation based on number of labels and existing settings
                    rotation_angle = 45 if len(x_tick_labels) > 3 else 0
                    ax.set_xticklabels(x_tick_labels, rotation=rotation_angle, fontsize=fontsize)

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
                ax.set_ylabel(self.xlabel_entry.get() or original_x_col, fontsize=fontsize)
                ax.set_xlabel(value_col if n_rows > 1 else (self.ylabel_entry.get() or value_col), fontsize=fontsize)
            else:
                # Enable math text rendering
                plt.rcParams['mathtext.default'] = 'regular'
                ax.set_xlabel(self.xlabel_entry.get() or x_col, fontsize=fontsize)
                # Enable math text rendering
                plt.rcParams['mathtext.default'] = 'regular'
                ax.set_ylabel(value_col if n_rows > 1 else (self.ylabel_entry.get() or value_col), fontsize=fontsize)

            rotation = 90 if self.label_orientation.get() == 'vertical' and not swap_axes else 0
            ax.tick_params(axis='x', labelsize=fontsize, rotation=rotation, direction='in', length=4, width=linewidth)
            ax.tick_params(axis='y', labelsize=fontsize, direction='in', length=4, width=linewidth, color='black', left=True)

            if plot_kind == "bar":
                if swap_axes:
                    ax.tick_params(axis='y', which='both', length=0)
                    ax.tick_params(axis='x', which='both', length=4, width=linewidth)
                else:
                    ax.tick_params(axis='x', which='both', length=0)
                    ax.tick_params(axis='y', which='both', length=4, width=linewidth)

            # --- The rest of your axis/tick/statistics/annotation code ---
            try:
                def _parse_float(val):
                    s = str(val).strip()
                    if not s:
                        return None
                    try:
                        return float(s)
                    except Exception:
                        import re
                        m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
                        return float(m.group(0)) if m else None

                def _parse_int(val):
                    s = str(val).strip()
                    if not s:
                        return None
                    try:
                        return int(s)
                    except Exception:
                        import re
                        m = re.search(r"[-+]?\d+", s)
                        return int(m.group(0)) if m else None

                ymin = _parse_float(self.ymin_entry.get())
                ymax = _parse_float(self.ymax_entry.get())
                yinterval = _parse_float(self.yinterval_entry.get())
                minor_ticks = _parse_int(self.minor_ticks_entry.get())

                xmin = _parse_float(self.xmin_entry.get())
                xmax = _parse_float(self.xmax_entry.get())
                xinterval = _parse_float(self.xinterval_entry.get())
                xminor_ticks = _parse_int(self.xminor_ticks_entry.get())

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

            # Apply logarithmic scales first
            # Y-axis logarithmic scale
            use_log_y = self.logscale_var.get()
            if use_log_y:
                # Get Y-axis log base
                y_log_base = int(self.ylog_base_var.get())
                if not swap_axes:
                    ax.set_yscale('log', base=y_log_base)
                else:
                    ax.set_xscale('log', base=y_log_base)
            
            # X-axis logarithmic scale
            use_log_x = self.xlogscale_var.get()
            if use_log_x:
                # Get X-axis log base
                x_log_base = int(self.xlog_base_var.get())
                if not swap_axes:
                    ax.set_xscale('log', base=x_log_base)
                else:
                    ax.set_yscale('log', base=x_log_base)

            # Apply appropriate tickers/locators for Y-axis
            minor_ticks_str = self.minor_ticks_entry.get()
            if minor_ticks_str:
                try:
                    minor_ticks = int(minor_ticks_str)
                    if use_log_y:
                        # Get Y-axis log base for ticks
                        y_log_base = int(self.ylog_base_var.get())
                        
                        # Set up minor ticks appropriate for the log base
                        if y_log_base == 10:
                            # For log10, use standard logarithmic ticks (1-9)
                            if minor_ticks >= 9:
                                # Use all digits as ticks if user wants many ticks
                                subs = np.arange(2, 10)
                            else:
                                # Distribute the ticks evenly across the decade
                                step = max(1, int(8 / minor_ticks))
                                subs = np.arange(2, 10, step)[:minor_ticks]
                        else:  # log2
                            # For log2, create appropriate subdivisions between each power of 2
                            if minor_ticks == 1:
                                subs = [1.5]  # Just one division at 1.5
                            elif minor_ticks == 2:
                                subs = [1.33, 1.66]  # Two evenly spaced points
                            elif minor_ticks == 3:
                                subs = [1.25, 1.5, 1.75]  # Three evenly spaced points
                            else:
                                # Multiple minor ticks between powers of 2
                                # Create evenly distributed points between 1 and 2 on a linear scale
                                # We don't include 1 or 2 as those are the major ticks
                                subs = np.linspace(1, 2, minor_ticks+2)[1:-1]
                                
                        if not swap_axes:
                            # Apply locators to the appropriate axis
                            ax.yaxis.set_major_locator(LogLocator(base=y_log_base, numticks=10))
                            # Fix minor ticks placement with appropriate number of ticks
                            if y_log_base == 2:
                                ax.yaxis.set_minor_locator(LogLocator(base=y_log_base, subs=subs, numticks=20))
                            else:
                                ax.yaxis.set_minor_locator(LogLocator(base=y_log_base, subs=subs, numticks=10))
                            # Make sure minor ticks are properly sized and colored
                            ax.tick_params(axis='y', which='minor', direction='in', length=2, width=linewidth/2, color='black')
                        else:
                            ax.xaxis.set_major_locator(LogLocator(base=y_log_base, numticks=10))
                            # Fix minor ticks placement with appropriate number of ticks
                            if y_log_base == 2:
                                ax.xaxis.set_minor_locator(LogLocator(base=y_log_base, subs=subs, numticks=20))
                            else:
                                ax.xaxis.set_minor_locator(LogLocator(base=y_log_base, subs=subs, numticks=10))
                            # Make sure minor ticks are properly sized and colored and don't show labels
                            ax.tick_params(axis='x', which='minor', direction='in', length=2, width=linewidth/2, 
                                          color='black', labelsize=0, labelbottom=False, labeltop=False, 
                                          labelleft=False, labelright=False)
                    else:
                        # Linear scale - use regular tick spacing
                        if not swap_axes:
                            ax.yaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                        else:
                            ax.xaxis.set_minor_locator(AutoMinorLocator(minor_ticks + 1))
                            
                    # Style the minor ticks
                    if not swap_axes:
                        ax.tick_params(axis='y', which='minor', direction='in', 
                                      length=2, width=linewidth, color='black', left=True)
                    else:
                        ax.tick_params(axis='x', which='minor', direction='in',
                                      length=2, width=linewidth, color='black', bottom=True)
                except Exception as e:
                    print(f"Y-axis minor ticks setting error: {e}")
            else:
                # No minor ticks requested
                if not swap_axes:
                    ax.yaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='y', which='minor', length=0)
                else:
                    ax.xaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='x', which='minor', length=0)
                    
            # Apply appropriate tickers/locators for X-axis
            xminor_ticks_str = self.xminor_ticks_entry.get()
            if xminor_ticks_str:
                try:
                    xminor_ticks = int(xminor_ticks_str)
                    if use_log_x:
                        # Get X-axis log base for ticks
                        x_log_base = int(self.xlog_base_var.get())
                        
                        # Set up minor ticks appropriate for the log base
                        if x_log_base == 10:
                            # For log10, use standard logarithmic ticks (1-9)
                            if xminor_ticks >= 9:
                                # Use all digits as ticks if user wants many ticks
                                subs = np.arange(2, 10)
                            else:
                                # Distribute the ticks evenly across the decade
                                step = max(1, int(8 / xminor_ticks))
                                subs = np.arange(2, 10, step)[:xminor_ticks]
                        else:  # log2
                            # For log2, create appropriate subdivisions between each power of 2
                            if xminor_ticks == 1:
                                subs = [1.5]  # Just one division at 1.5
                            elif xminor_ticks == 2:
                                subs = [1.33, 1.66]  # Two evenly spaced points
                            elif xminor_ticks == 3:
                                subs = [1.25, 1.5, 1.75]  # Three evenly spaced points
                            else:
                                # Multiple minor ticks between powers of 2
                                # Create evenly distributed points between 1 and 2 on a linear scale
                                # We don't include 1 or 2 as those are the major ticks
                                subs = np.linspace(1, 2, xminor_ticks+2)[1:-1]
                                
                        if not swap_axes:
                            # Apply locators to the appropriate axis
                            ax.xaxis.set_major_locator(LogLocator(base=x_log_base, numticks=10))
                            # Fix minor ticks placement with appropriate number of ticks
                            if x_log_base == 2:
                                ax.xaxis.set_minor_locator(LogLocator(base=x_log_base, subs=subs, numticks=20))
                            else:
                                ax.xaxis.set_minor_locator(LogLocator(base=x_log_base, subs=subs, numticks=10))
                            # Make sure minor ticks are properly sized and colored and don't show labels
                            ax.tick_params(axis='x', which='minor', direction='in', length=2, width=linewidth/2, 
                                          color='black', labelsize=0, labelbottom=False, labeltop=False, 
                                          labelleft=False, labelright=False)
                        else:
                            ax.yaxis.set_major_locator(LogLocator(base=x_log_base, numticks=10))
                            # Fix minor ticks placement with appropriate number of ticks
                            if x_log_base == 2:
                                ax.yaxis.set_minor_locator(LogLocator(base=x_log_base, subs=subs, numticks=20))
                            else:
                                ax.yaxis.set_minor_locator(LogLocator(base=x_log_base, subs=subs, numticks=10))
                            # Make sure minor ticks are properly sized and colored
                            ax.tick_params(axis='y', which='minor', direction='in', length=2, width=linewidth/2, color='black')
                    else:
                        # Linear scale - use regular tick spacing
                        if not swap_axes:
                            ax.xaxis.set_minor_locator(AutoMinorLocator(xminor_ticks + 1))
                        else:
                            ax.yaxis.set_minor_locator(AutoMinorLocator(xminor_ticks + 1))
                            
                    # Style the minor ticks
                    if not swap_axes:
                        ax.tick_params(axis='x', which='minor', direction='in', 
                                      length=2, width=linewidth, color='black', bottom=True)
                    else:
                        ax.tick_params(axis='y', which='minor', direction='in',
                                      length=2, width=linewidth, color='black', left=True)
                except Exception as e:
                    print(f"X-axis minor ticks setting error: {e}")
            else:
                # No minor ticks requested
                if not swap_axes:
                    ax.xaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='x', which='minor', length=0)
                else:
                    ax.yaxis.set_minor_locator(NullLocator())
                    ax.tick_params(axis='y', which='minor', length=0)
                    
            # Set general tick parameters
            ax.tick_params(axis='both', which='both', direction='in', width=linewidth)
            ax.tick_params(axis='both', which='major', length=4)
            # Ensure minor ticks never have labels
            ax.tick_params(axis='both', which='minor', labelsize=0, labelbottom=False, labeltop=False, 
                           labelleft=False, labelright=False)

            if plot_kind == "bar":
                if swap_axes:
                    ax.tick_params(axis='y', which='both', length=0)
                    ax.tick_params(axis='x', which='major', bottom=True, top=False, length=4, width=linewidth, color='black')
                else:
                    ax.tick_params(axis='x', which='both', length=0)
                    ax.tick_params(axis='y', which='major', left=True, right=False, length=4, width=linewidth, color='black')
            
            # Apply tick label orientation based on user selection
            orientation = self.label_orientation.get()
            if not swap_axes:
                # X-axis orientation when axes are not swapped
                if orientation == "horizontal":
                    ax.tick_params(axis='x', which='major', labelrotation=0)
                elif orientation == "vertical":
                    ax.tick_params(axis='x', which='major', labelrotation=90)
                elif orientation == "angled":
                    # For angled labels, we need to set both rotation and alignment to center properly
                    ax.tick_params(axis='x', which='major', labelrotation=45)
                    # Set proper alignment for angled labels to center under their ticks
                    for label in ax.get_xticklabels():
                        label.set_horizontalalignment('right')
                        label.set_rotation_mode('anchor')
            else:
                # Y-axis orientation when axes are swapped (becomes the category axis)
                if orientation == "horizontal":
                    ax.tick_params(axis='y', which='major', labelrotation=0)
                elif orientation == "vertical":
                    ax.tick_params(axis='y', which='major', labelrotation=90)
                elif orientation == "angled":
                    # For angled labels, we need to set both rotation and alignment to center properly
                    ax.tick_params(axis='y', which='major', labelrotation=45)
                    # Set proper alignment for angled labels to center under their ticks
                    for label in ax.get_yticklabels():
                        label.set_verticalalignment('top')
                        label.set_rotation_mode('anchor')

            ax_position = [left_margin / fig_width,
                           bottom_margin / fig_height,
                           plot_width / fig_width,
                           plot_height / fig_height]
            ax.set_position(ax_position)

            # --- Statistics/Annotations ---
            if self.use_stats_var.get():
                try:
                    print(f"[DEBUG] Starting statistics calculation: x_col={x_col}, value_col={value_col}, hue_col={hue_col}")
                    # Calculate p-values for all pairs
                    self.calculate_and_store_pvals(df_plot, x_col, value_col, hue_col)
                    print(f"[DEBUG] After calculate_and_store_pvals, latest_pvals = {self.latest_pvals}")
                    
                    # Only add annotations if the checkbox is selected
                    show_annotations = getattr(self, 'show_statistics_annotations_var', None)
                    if show_annotations is None or not show_annotations.get():
                        print(f"[DEBUG] Statistical annotations disabled by user preference")
                        continue
                    if disable_plot_stat_annotations:
                        continue
                    
                    # Import required modules for annotations
                    from statannotations.Annotator import Annotator
                    import itertools
                    

                    # Define p-value format configuration for statannotations using current alpha level
                    try:
                        alpha = float(self.alpha_level_var.get())
                    except (ValueError, AttributeError):
                        alpha = 0.05  # Default if not set or invalid
                    
                    # Use the same threshold logic as in our pval_to_annotation function
                    pvalue_format = {
                        'text_format': 'star',
                        'pvalue_thresholds': [
                            (alpha/5000, '****'),  # 4 stars threshold (e.g., 0.00001 at alpha=0.05)
                            (alpha/50, '***'),    # 3 stars threshold (e.g., 0.001 at alpha=0.05)
                            (alpha/5, '**'),      # 2 stars threshold (e.g., 0.01 at alpha=0.05)
                            (alpha, '*'),         # 1 star threshold (alpha itself)
                            (1, 'ns')             # Not significant
                        ]
                    }
                    
                    print(f"[DEBUG] Using pvalue thresholds: {[t[0] for t in pvalue_format['pvalue_thresholds']]}")

                    annot_x_col = x_col
                    annot_y_col = value_col
                    annot_orient = 'v'
                    if swap_axes and plot_kind in ("bar", "box", "violin"):
                        annot_x_col = value_col
                        annot_y_col = x_col
                        annot_orient = 'h'
                    
                    
                    # Determine what kind of annotations we need based on the data structure
                    if hue_col and hue_col in df_plot.columns:
                        # For data with groups (hue column), create pairwise comparisons within each x-category
                        hue_groups = list(df_plot[hue_col].dropna().unique())
                        base_groups = list(df_plot[x_col].dropna().unique())
                        print(f"[DEBUG] Hue annotation case: x_values={base_groups}, hue_groups={hue_groups}")
                        
                        # For each x value, create pairs to compare the hue groups
                        for x_val in base_groups:
                            # Get all pairwise combinations of hue groups
                            hue_pairs = list(itertools.combinations(hue_groups, 2))
                            
                            # Create list of comparisons
                            pairs_to_compare = []
                            pair_pvalues = []
                            
                            for h1, h2 in hue_pairs:
                                # Try different key formats to find p-values
                                key = self.stat_key(x_val, h1, h2)
                                pval = self.latest_pvals.get(key)
                                
                                if pval is not None and not pd.isna(pval):
                                    # Add to our comparison list
                                    pairs_to_compare.append([(x_val, h1), (x_val, h2)])
                                    pair_pvalues.append(pval)
                            
                            if pairs_to_compare:
                                try:
                                    # Create annotator
                                    annotator = Annotator(ax, pairs_to_compare, data=df_plot, 
                                                       x=annot_x_col, y=annot_y_col, hue=hue_col, orient=annot_orient)
                                    
                                    # Add annotations
                                    annotator.configure(test=None, text_format='star', 
                                                     pvalue_format=pvalue_format,
                                                     loc='inside',
                                                     line_width=self.linewidth.get()
                                                     )
                                    annotator.set_pvalues(pair_pvalues)
                                    annotator.annotate()
                                except Exception as e:
                                    print(f"[DEBUG] Error adding annotations for {x_val}: {e}")
                    else:
                        # For data without groups (no hue column), compare between x-categories directly
                        x_categories = list(df_plot[x_col].dropna().unique())
                        if len(x_categories) > 1:
                            # Get all pairwise combinations of x categories
                            x_pairs = list(itertools.combinations(x_categories, 2))
                            
                            # Create list of comparisons
                            pairs_to_compare = []
                            pair_pvalues = []
                            
                            for x1, x2 in x_pairs:
                                # Try different key formats to find p-values
                                key = self.stat_key(x1, x2)
                                pval = self.latest_pvals.get(key)
                                
                                if pval is not None and not pd.isna(pval):
                                    # Add to our comparison list
                                    pairs_to_compare.append([x1, x2])
                                    pair_pvalues.append(pval)
                        
                            if pairs_to_compare:
                                try:
                                    # Create annotator for x-category comparisons
                                    annotator = Annotator(ax, pairs_to_compare, data=df_plot, 
                                                       x=annot_x_col, y=annot_y_col, orient=annot_orient)
                                    
                                    # Add annotations
                                    annotator.configure(test=None, text_format='star', 
                                                     pvalue_format=pvalue_format,
                                                     loc='inside',
                                                     line_width=self.linewidth.get()
                                                     )
                                    annotator.set_pvalues(pair_pvalues)
                                    annotator.annotate()
                                except Exception:
                                    pass
                    
                    # Save the figure for later reference
                    fig = ax.figure
                except Exception as e:
                    ax.text(0.01, 0.98, f"Stats error: {e}", transform=ax.transAxes, fontsize=fontsize*0.8, va='top', ha='left',
                            bbox=dict(boxstyle="round,pad=0.3", fc="white", alpha=0.8))
            # Reset all linewidths (defensive: after annotations/statistics)
            for spine in ax.spines.values():
                spine.set_linewidth(linewidth)
            ax.tick_params(axis='both', which='both', width=linewidth)
            # (grid lines already set above)

        # Apply Y-axis break if enabled
        if self.ybreak_enabled and self.ax_upper is not None and self.ax_lower is not None:
            self.apply_yaxis_break(self.ax_lower, self.ax_upper, self.ybreak_min, self.ybreak_max,
                                  show_frame, linewidth)

        try:
            self._sync_custom_marks_vars_from_ui()
            x_marks = self._parse_custom_marks(self.custom_x_marks_var.get() if hasattr(self, 'custom_x_marks_var') else '', axis='x')
            y_marks = self._parse_custom_marks(self.custom_y_marks_var.get() if hasattr(self, 'custom_y_marks_var') else '', axis='y')
            for _ax in self.fig.axes:
                draw_x_labels = True
                if getattr(self, 'ybreak_enabled', False) and getattr(self, 'ax_upper', None) is not None and _ax is self.ax_upper:
                    draw_x_labels = False
                self._apply_custom_marks_to_axis(_ax, 'x', x_marks, linewidth=linewidth, fontsize=fontsize, draw_labels=draw_x_labels)
                self._apply_custom_marks_to_axis(_ax, 'y', y_marks, linewidth=linewidth, fontsize=fontsize, draw_labels=True)
        except Exception:
            pass
        
        # Apply legend settings from UI to all axes
        legend_visible = self.legend_visible_var.get()
        legend_position = self.legend_position_var.get()
        legend_ncol = self.legend_ncol_var.get()
        
        for ax in self.fig.axes:
            old_legend = ax.get_legend()
            if old_legend:
                # Get handles and labels from existing legend
                handles = old_legend.legend_handles
                labels = [t.get_text() for t in old_legend.get_texts()]
                
                # Remove old legend
                old_legend.remove()
                
                if not legend_visible or not handles:
                    continue
                
                # Determine number of columns
                if legend_ncol > 0:
                    ncol = legend_ncol
                else:
                    ncol = self.optimize_legend_layout(ax, handles, labels, fontsize=fontsize)
                
                # Create new legend with user settings
                if legend_position == "outside right":
                    new_legend = ax.legend(handles, labels, bbox_to_anchor=(1.02, 1), loc='upper left',
                                          borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)
                elif legend_position == "outside top":
                    # For axis break, the "main" axes bbox is the pre-split bbox stored during apply_yaxis_break.
                    # Anchor the outside-top legend in FIGURE coordinates so its distance to the plot area
                    # matches the non-broken plot.
                    if getattr(self, 'ybreak_enabled', False) and hasattr(self, '_ybreak_base_pos') and self._ybreak_base_pos is not None:
                        base_pos = self._ybreak_base_pos
                        x_center = base_pos.x0 + base_pos.width / 2
                        y_top = base_pos.y0 + base_pos.height
                        new_legend = ax.legend(
                            handles, labels,
                            bbox_to_anchor=(x_center, y_top + base_pos.height * 0.15),
                            bbox_transform=self.fig.transFigure,
                            loc='upper center',
                            borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol
                        )
                    else:
                        new_legend = ax.legend(handles, labels, bbox_to_anchor=(0.5, 1.15), loc='upper center',
                                              borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)
                else:
                    # Standard matplotlib positions
                    new_legend = ax.legend(handles, labels, loc=legend_position, frameon=False,
                                          fontsize=fontsize, ncol=ncol)
                
                # Make legend draggable
                if new_legend:
                    new_legend.set_draggable(True)
        
        # Display interactive preview (before saving, to preserve figure state)
        self.display_preview()

    def apply_yaxis_break(self, ax_lower, ax_upper, break_min, break_max, show_frame, linewidth):
        """
        Apply Y-axis break styling to the dual axes.
        Copies plot elements from lower to upper, sets y-limits, hides spines, adds break markers.
        """
        # Get current y-limits from the lower axis (which has all the data)
        y_data_min, y_data_max = ax_lower.get_ylim()
        
        # Get user-specified y-axis limits if provided
        try:
            user_ymin = float(self.ymin_entry.get()) if self.ymin_entry.get().strip() else None
        except ValueError:
            user_ymin = None
        try:
            user_ymax = float(self.ymax_entry.get()) if self.ymax_entry.get().strip() else None
        except ValueError:
            user_ymax = None
        
        # Set y-limits for lower axis (from 0 or user min to break_min)
        lower_ymin = user_ymin if user_ymin is not None else 0
        ax_lower.set_ylim(lower_ymin, break_min)
        
        # Set y-limits for upper axis (from break_max to data max or user max)
        upper_ymax = user_ymax if user_ymax is not None else y_data_max * 1.05
        ax_upper.set_ylim(break_max, upper_ymax)
        
        # Reposition axes while PRESERVING the original y-scale.
        # Previous behavior renormalized (lower+upper) to fill the entire plot height,
        # making the visible bars/points appear larger than the non-broken plot.
        #
        # New behavior: allocate a *proportional* gap height that corresponds to the removed
        # y-range (break_max - break_min). This keeps pixels-per-y-unit consistent with the
        # non-broken plot.
        lower_range = break_min - lower_ymin
        upper_range = upper_ymax - break_max
        total_visible_range = lower_range + upper_range
        
        if total_visible_range > 0:
            # IMPORTANT: the app later sets the "main" axis position explicitly via ax.set_position(...)
            # (see ax_position assignment in plot_graph). When axis break is enabled, that adjustment
            # is applied to ax_lower (the "main" axis) but not necessarily to the combined two-subplot
            # bbox from plt.subplots. If we base our layout on the combined bbox, the broken-axis plot
            # area can become larger than the non-broken plot.
            #
            # Therefore, always split WITHIN the final current bbox of ax_lower.
            base_pos = ax_lower.get_position().frozen()
            self._ybreak_base_pos = base_pos
            left = base_pos.x0
            width = base_pos.width
            bottom = base_pos.y0
            total_height = base_pos.height
            top = bottom + total_height

            gap_frac = getattr(self, 'ybreak_gap', None)
            if gap_frac is None:
                gap_frac = 0.07
            gap_frac = max(0.0, min(0.3, float(gap_frac)))
            gap = total_height * gap_frac
            available_height = max(total_height - gap, 0)
            
            # Allocate heights either from user ratio (upper_ratio) or automatically from visible y ranges
            upper_ratio = getattr(self, 'ybreak_ratio', None)
            if upper_ratio is not None:
                upper_ratio = max(0.1, min(0.9, float(upper_ratio)))
                upper_height = available_height * upper_ratio
                lower_height = available_height - upper_height
            else:
                lower_height = available_height * (lower_range / total_visible_range)
                upper_height = available_height * (upper_range / total_visible_range)
            
            # Position lower axis at bottom, upper axis above it with proportional gap
            ax_lower.set_position([left, bottom, width, lower_height])
            ax_upper.set_position([left, bottom + lower_height + gap, width, upper_height])

            self.debug(
                f"Repositioned (preserve y-scale): lower_h={lower_height:.3f}, gap={gap:.3f}, "
                f"upper_h={upper_height:.3f}, total={lower_height+gap+upper_height:.3f} vs orig={top-bottom:.3f}"
            )
        
        # Note: Plot elements are now created by seaborn on both axes during plotting
        # No need to copy elements - just set y-limits and styling
        
        # Get fontsize from the app settings
        try:
            fontsize = self.fontsize.get()
        except:
            fontsize = 10
        
        # Copy EXACT styling from lower axis to upper axis
        # This ensures identical appearance (font size, tick params, spine visibility, etc.)
        
        # Copy Y-axis tick label font size and tick params from lower
        ax_upper.tick_params(axis='y', labelsize=fontsize, direction='in', length=4, width=linewidth, color='black', left=True)

        # Copy y-axis tick locators/formatters so minor interval settings apply to both axes
        try:
            # IMPORTANT: matplotlib Locator/Formatter instances are bound to a single Axis.
            # Reusing the same instance can detach it from ax_lower (causing missing ticks/labels).
            # Clone them for ax_upper.
            lower_major_loc = ax_lower.yaxis.get_major_locator()
            lower_minor_loc = ax_lower.yaxis.get_minor_locator()
            lower_major_fmt = ax_lower.yaxis.get_major_formatter()
            lower_minor_fmt = ax_lower.yaxis.get_minor_formatter()

            ax_upper.yaxis.set_major_locator(copy.deepcopy(lower_major_loc))
            ax_upper.yaxis.set_minor_locator(copy.deepcopy(lower_minor_loc))
            ax_upper.yaxis.set_major_formatter(copy.deepcopy(lower_major_fmt))
            ax_upper.yaxis.set_minor_formatter(copy.deepcopy(lower_minor_fmt))
        except Exception as e:
            try:
                ax_upper.yaxis.set_major_locator(copy.copy(ax_lower.yaxis.get_major_locator()))
                ax_upper.yaxis.set_minor_locator(copy.copy(ax_lower.yaxis.get_minor_locator()))
                ax_upper.yaxis.set_major_formatter(copy.copy(ax_lower.yaxis.get_major_formatter()))
                ax_upper.yaxis.set_minor_formatter(copy.copy(ax_lower.yaxis.get_minor_formatter()))
            except Exception as e2:
                print(f"[DEBUG] Could not copy y-axis locator/formatter to upper axis: {e}; fallback failed: {e2}")
        
        # Copy spine linewidth from lower to upper
        for spine_name, spine in ax_upper.spines.items():
            lower_spine = ax_lower.spines[spine_name]
            spine.set_linewidth(lower_spine.get_linewidth())
            spine.set_color(lower_spine.get_edgecolor())
        
        # Copy spine visibility from lower (except for the break spines)
        ax_upper.spines['left'].set_visible(ax_lower.spines['left'].get_visible())
        ax_upper.spines['right'].set_visible(ax_lower.spines['right'].get_visible())
        ax_upper.spines['top'].set_visible(ax_lower.spines['top'].get_visible())
        
        # Hide the spines between the two axes (the break)
        ax_upper.spines['bottom'].set_visible(False)
        ax_lower.spines['top'].set_visible(False)
        
        # Hide x-axis ticks on upper plot (they show on lower)
        ax_upper.tick_params(axis='x', which='both', bottom=False, labelbottom=False, top=False, labeltop=False)
        ax_upper.set_xlabel('')  # Remove x-label from upper

        # Mirror grid visibility and styling from lower axis
        try:
            y_gridlines_lower = ax_lower.get_ygridlines()
            y_grid_on = any(gl.get_visible() for gl in y_gridlines_lower)
            ax_upper.grid(y_grid_on, axis='y', which='major')
            if y_gridlines_lower:
                src = y_gridlines_lower[0]
                for gl in ax_upper.get_ygridlines():
                    gl.set_color(src.get_color())
                    gl.set_linestyle(src.get_linestyle())
                    gl.set_linewidth(src.get_linewidth())
                    gl.set_alpha(src.get_alpha())
                    gl.set_zorder(src.get_zorder())

            x_gridlines_lower = ax_lower.get_xgridlines()
            x_grid_on = any(gl.get_visible() for gl in x_gridlines_lower)
            ax_upper.grid(x_grid_on, axis='x', which='major')
            if x_gridlines_lower:
                src = x_gridlines_lower[0]
                for gl in ax_upper.get_xgridlines():
                    gl.set_color(src.get_color())
                    gl.set_linestyle(src.get_linestyle())
                    gl.set_linewidth(src.get_linewidth())
                    gl.set_alpha(src.get_alpha())
                    gl.set_zorder(src.get_zorder())
        except Exception as e:
            print(f"[DEBUG] Could not copy grid settings to upper axis: {e}")

        try:
            if hasattr(self, '_ybreak_marker_artists') and self._ybreak_marker_artists:
                for a in list(self._ybreak_marker_artists):
                    try:
                        a.remove()
                    except Exception:
                        pass
                self._ybreak_marker_artists = []
        except Exception:
            pass

        try:
            if hasattr(self, '_ybreak_frame_connector_artists') and self._ybreak_frame_connector_artists:
                for a in list(self._ybreak_frame_connector_artists):
                    try:
                        a.remove()
                    except Exception:
                        pass
                self._ybreak_frame_connector_artists = []
        except Exception:
            pass

        style = None
        try:
            style = self.ybreak_marker_style_var.get() if hasattr(self, 'ybreak_marker_style_var') else None
        except Exception:
            style = None
        style = style or "Diagonal"

        try:
            marker_lw = ax_lower.spines['left'].get_linewidth()
        except Exception:
            marker_lw = linewidth

        if style == "Connected":
            try:
                fig = ax_lower.figure
                lower_pos = ax_lower.get_position().frozen()
                upper_pos = ax_upper.get_position().frozen()
                y_bot_gap = lower_pos.y0 + lower_pos.height
                y_top_gap = upper_pos.y0
                gap_h = max(y_top_gap - y_bot_gap, 0)

                # Connect directly to the broken spine endpoints
                mid_y = (y_bot_gap + y_top_gap) / 2
                dx_fig = max(0.004, min(0.012, lower_pos.width * 0.03))
                dx_axes = (dx_fig / lower_pos.width) if lower_pos.width else 0.03
                dx_axes = max(0.015, min(0.08, dx_axes))

                # x in axes coords (so it hits the y-axis spine exactly), y in figure coords (so it sits in the gap)
                x0 = 0.0
                t = blended_transform_factory(ax_lower.transAxes, fig.transFigure)

                # Single connected 'zig' (like '>')
                ys = [y_top_gap, mid_y, y_bot_gap]
                xs = [x0, x0 + dx_axes, x0]

                line = Line2D(xs, ys, transform=t, color='black', linewidth=marker_lw, clip_on=False)
                try:
                    line.set_solid_capstyle('projecting')
                except Exception:
                    pass
                try:
                    line.set_zorder(ax_lower.spines['left'].get_zorder() + 2)
                except Exception:
                    pass
                fig.add_artist(line)
                self._ybreak_marker_artists = [line]
            except Exception:
                pass
        else:
            # Diagonal slashes, but drawn as fixed-size marker paths (consistent visual angle)
            marker_path = MplPath([(-1, -1), (1, 1)], [MplPath.MOVETO, MplPath.LINETO])
            x_left = 0.0
            x_right = 1.0
            ms = 5.0

            marker_kwargs_upper = dict(
                marker=marker_path,
                markersize=ms,
                markeredgewidth=marker_lw,
                color='black',
                linestyle='none',
                transform=ax_upper.transAxes,
                clip_on=False
            )
            ax_upper.plot([x_left], [0], **marker_kwargs_upper)
            if show_frame:
                ax_upper.plot([x_right], [0], **marker_kwargs_upper)

            marker_kwargs_lower = marker_kwargs_upper.copy()
            marker_kwargs_lower['transform'] = ax_lower.transAxes
            ax_lower.plot([x_left], [1], **marker_kwargs_lower)
            if show_frame:
                ax_lower.plot([x_right], [1], **marker_kwargs_lower)

        # Close the frame gap on the right side (opposite of the y-axis break) when frame is enabled.
        try:
            if show_frame and ax_lower.spines['right'].get_visible():
                fig = ax_lower.figure
                lower_pos = ax_lower.get_position().frozen()
                upper_pos = ax_upper.get_position().frozen()
                y_bot_gap = lower_pos.y0 + lower_pos.height
                y_top_gap = upper_pos.y0

                x_right = upper_pos.x0 + upper_pos.width
                spine = ax_lower.spines['right']
                spine_lw = spine.get_linewidth()
                spine_color = spine.get_edgecolor()

                conn = Line2D([x_right, x_right], [y_bot_gap, y_top_gap],
                              transform=fig.transFigure, color=spine_color,
                              linewidth=spine_lw, clip_on=False)
                try:
                    conn.set_solid_capstyle('projecting')
                except Exception:
                    pass
                try:
                    conn.set_zorder(spine.get_zorder() + 1)
                except Exception:
                    pass
                fig.add_artist(conn)
                self._ybreak_frame_connector_artists = [conn]
        except Exception:
            pass
        
        # Handle legend - move from lower to upper axis
        lower_legend = ax_lower.get_legend()
        if lower_legend:
            # Get handles and labels from lower axis legend
            handles = lower_legend.legend_handles
            labels = [t.get_text() for t in lower_legend.get_texts()]
            fontsize = lower_legend.get_texts()[0].get_fontsize() if lower_legend.get_texts() else 10
            
            # Remove legend from lower axis
            lower_legend.remove()

            # Remove any existing legend on upper axis, then re-place using the same helper
            # as non-broken plots for consistent layout.
            try:
                if ax_upper.get_legend():
                    ax_upper.get_legend().remove()
            except Exception:
                pass

            if hasattr(self, 'place_legend'):
                self.place_legend(ax_upper, handles, labels)
            else:
                ax_upper.legend(handles, labels, loc='best', frameon=False, fontsize=fontsize)
        
        # Handle Y-axis label: center it across BOTH axes.
        # Remove per-axis ylabels and draw a figure-level ylabel at the midpoint.
        try:
            fig = ax_lower.figure
            ylabel_text = ax_lower.get_ylabel()
            if ylabel_text:
                if hasattr(self, '_ybreak_ylabel_artist') and self._ybreak_ylabel_artist is not None:
                    try:
                        self._ybreak_ylabel_artist.remove()
                    except Exception:
                        pass
                    self._ybreak_ylabel_artist = None

                label_artist = ax_lower.yaxis.label
                x_fig = None
                try:
                    if hasattr(fig, 'canvas') and fig.canvas is not None:
                        fig.canvas.draw()
                        renderer = fig.canvas.get_renderer()
                        bbox = label_artist.get_window_extent(renderer=renderer)
                        x_center_disp = bbox.x0 + bbox.width / 2
                        x_fig, _ = fig.transFigure.inverted().transform((x_center_disp, 0))
                except Exception:
                    x_fig = None

                if x_fig is None:
                    label_pos_axes = label_artist.get_position()
                    x_disp, _ = ax_lower.transAxes.transform(label_pos_axes)
                    x_fig, _ = fig.transFigure.inverted().transform((x_disp, 0))

                lower_pos = ax_lower.get_position().frozen()
                upper_pos = ax_upper.get_position().frozen()
                y_center = (lower_pos.y0 + (upper_pos.y0 + upper_pos.height)) / 2

                self._ybreak_ylabel_artist = fig.text(
                    x_fig,
                    y_center,
                    ylabel_text,
                    rotation=90,
                    va='center',
                    ha='center',
                    fontsize=label_artist.get_fontsize(),
                    color=label_artist.get_color()
                )

            ax_lower.set_ylabel('')
            ax_upper.set_ylabel('')
        except Exception as e:
            print(f"[DEBUG] Could not center y-axis label for axis break: {e}")

    def display_preview(self):
        """Display an interactive preview using matplotlib's native tkinter widget.
        
        This allows draggable legends - drag the legend to reposition it,
        then click 'Save PDF' to export with the new legend position.
        """
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        
        # Clear existing widgets
        for widget in self.canvas_frame.winfo_children():
            widget.destroy()
            
        # Show or hide the statistical details button if statistics were calculated
        if self.use_stats_var.get() and hasattr(self, 'latest_stats') and self.latest_stats:
            try:
                self.stats_details_btn.grid()
            except Exception as e:
                print(f"Error showing stats details button: {e}")
        else:
            try:
                self.stats_details_btn.grid_remove()
            except Exception as e:
                print(f"Error hiding stats details button: {e}")
        
        if not hasattr(self, 'fig') or self.fig is None:
            # No figure available - show placeholder
            placeholder = ttk.Label(self.canvas_frame, text="No plot generated yet")
            placeholder.pack(pady=50)
            return
            
        try:
            fig_dpi = self.fig.dpi

            if (not hasattr(self, '_preview_base_geometry') or
                self._preview_base_geometry is None or
                self._preview_base_geometry.get('fig_id') != id(self.fig)):
                self._preview_base_geometry = {
                    'fig_id': id(self.fig),
                    'fig_size': tuple(self.fig.get_size_inches()),
                    'axes_positions': [ax.get_position().bounds for ax in self.fig.axes]
                }
            else:
                try:
                    base_size = self._preview_base_geometry.get('fig_size')
                    base_axes = self._preview_base_geometry.get('axes_positions')
                    if base_size and base_axes and len(base_axes) == len(self.fig.axes):
                        self.fig.set_size_inches(base_size[0], base_size[1])
                        for ax, pos in zip(self.fig.axes, base_axes):
                            ax.set_position(pos)
                except Exception:
                    pass

            base_w_in, base_h_in = self._preview_base_geometry.get('fig_size')
            base_w_px = int(base_w_in * fig_dpi)
            base_h_px = int(base_h_in * fig_dpi)

            legend_position = self.legend_position_var.get() if hasattr(self, 'legend_position_var') else 'best'

            avg_dim = (base_w_px + base_h_px) / 2
            pad_px = max(int(avg_dim * 0.60), 600)
            if legend_position == 'outside right':
                pad_px = max(int(avg_dim * 0.90), 900)
            elif legend_position == 'outside top':
                pad_px = max(int(avg_dim * 0.75), 800)

            new_w_px = base_w_px + (2 * pad_px)
            new_h_px = base_h_px + (2 * pad_px)
            new_w_in = new_w_px / fig_dpi
            new_h_in = new_h_px / fig_dpi

            self.fig.set_size_inches(new_w_in, new_h_in)

            try:
                base_axes = self._preview_base_geometry.get('axes_positions')
                if base_axes and len(base_axes) == len(self.fig.axes):
                    for ax, b in zip(self.fig.axes, base_axes):
                        left_px = b[0] * base_w_px
                        bottom_px = b[1] * base_h_px
                        width_px = b[2] * base_w_px
                        height_px = b[3] * base_h_px
                        ax.set_position([
                            (left_px + pad_px) / new_w_px,
                            (bottom_px + pad_px) / new_h_px,
                            width_px / new_w_px,
                            height_px / new_h_px
                        ])
            except Exception:
                pass

            try:
                if getattr(self, 'ybreak_enabled', False) and getattr(self, 'ax_lower', None) is not None and getattr(self, 'ax_upper', None) is not None:
                    try:
                        lower_pos = self.ax_lower.get_position().frozen()
                        upper_pos = self.ax_upper.get_position().frozen()

                        x0 = min(lower_pos.x0, upper_pos.x0)
                        y0 = min(lower_pos.y0, upper_pos.y0)
                        x1 = max(lower_pos.x0 + lower_pos.width, upper_pos.x0 + upper_pos.width)
                        y1 = max(lower_pos.y0 + lower_pos.height, upper_pos.y0 + upper_pos.height)

                        bbox_cls = type(lower_pos)
                        self._ybreak_base_pos = bbox_cls.from_extents(x0, y0, x1, y1).frozen()
                    except Exception:
                        self._ybreak_base_pos = self.ax_lower.get_position().frozen()

                    if hasattr(self, '_ybreak_ylabel_artist') and self._ybreak_ylabel_artist is not None:
                        try:
                            old_x, _ = self._ybreak_ylabel_artist.get_position()
                            new_x = (old_x * base_w_px + pad_px) / new_w_px

                            lower_pos = self.ax_lower.get_position().frozen()
                            upper_pos = self.ax_upper.get_position().frozen()
                            y_center = (lower_pos.y0 + (upper_pos.y0 + upper_pos.height)) / 2

                            self._ybreak_ylabel_artist.set_position((new_x, y_center))
                        except Exception:
                            pass

                    try:
                        if (hasattr(self, '_ybreak_marker_artists') and self._ybreak_marker_artists and
                            hasattr(self, 'ybreak_marker_style_var') and self.ybreak_marker_style_var.get() == 'Connected'):
                            lower_pos = self.ax_lower.get_position().frozen()
                            upper_pos = self.ax_upper.get_position().frozen()
                            y_bot_gap = lower_pos.y0 + lower_pos.height
                            y_top_gap = upper_pos.y0
                            gap_h = max(y_top_gap - y_bot_gap, 0)
                            mid_y = (y_bot_gap + y_top_gap) / 2
                            dx_fig = max(0.004, min(0.012, lower_pos.width * 0.03))
                            dx_axes = (dx_fig / lower_pos.width) if lower_pos.width else 0.03
                            dx_axes = max(0.015, min(0.08, dx_axes))
                            x0 = 0.0

                            ys = [y_top_gap, mid_y, y_bot_gap]
                            xs = [x0, x0 + dx_axes, x0]

                            for a in list(self._ybreak_marker_artists):
                                try:
                                    try:
                                        a.set_transform(blended_transform_factory(self.ax_lower.transAxes, self.fig.transFigure))
                                    except Exception:
                                        pass
                                    a.set_data(xs, ys)
                                    try:
                                        a.set_solid_capstyle('projecting')
                                    except Exception:
                                        pass
                                    try:
                                        a.set_zorder(self.ax_lower.spines['left'].get_zorder() + 2)
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    try:
                        if (hasattr(self, '_ybreak_frame_connector_artists') and self._ybreak_frame_connector_artists and
                            getattr(self, 'ax_lower', None) is not None and getattr(self, 'ax_upper', None) is not None):
                            lower_pos = self.ax_lower.get_position().frozen()
                            upper_pos = self.ax_upper.get_position().frozen()
                            y_bot_gap = lower_pos.y0 + lower_pos.height
                            y_top_gap = upper_pos.y0
                            x_right = upper_pos.x0 + upper_pos.width

                            for a in list(self._ybreak_frame_connector_artists):
                                try:
                                    a.set_data([x_right, x_right], [y_bot_gap, y_top_gap])
                                    try:
                                        a.set_transform(self.fig.transFigure)
                                    except Exception:
                                        pass
                                except Exception:
                                    pass
                    except Exception:
                        pass
            except Exception:
                pass

            fig_width_inches, fig_height_inches = self.fig.get_size_inches()
            canvas_width = int(fig_width_inches * fig_dpi)
            canvas_height = int(fig_height_inches * fig_dpi)

            pad_left = 0
            pad_top = 0
            padded_width = canvas_width
            padded_height = canvas_height

            padding_frame = tk.Frame(self.canvas_frame, bg='white', width=padded_width, height=padded_height)
            padding_frame.pack(fill='none', expand=False, anchor='nw')
            padding_frame.pack_propagate(False)
            
            # Create the matplotlib canvas widget (keeps figure at native DPI)
            self.mpl_canvas = FigureCanvasTkAgg(self.fig, master=padding_frame)
            canvas_widget = self.mpl_canvas.get_tk_widget()
            
            # Set fixed size on canvas widget to prevent tkinter from resizing it
            canvas_widget.configure(width=canvas_width, height=canvas_height)
            
            canvas_widget.place(x=pad_left, y=pad_top)
            
            # Draw the canvas
            self.mpl_canvas.draw()

            try:
                if getattr(self, 'ybreak_enabled', False) and legend_position == "outside top":
                    legend_visible = self.legend_visible_var.get() if hasattr(self, 'legend_visible_var') else True
                    legend_ncol = self.legend_ncol_var.get() if hasattr(self, 'legend_ncol_var') else 0
                    fontsize = self.fontsize.get() if hasattr(self, 'fontsize') else 10

                    for ax in self.fig.axes:
                        old_legend = ax.get_legend()
                        if not old_legend:
                            continue

                        handles = old_legend.legend_handles
                        labels = [t.get_text() for t in old_legend.get_texts()]
                        old_legend.remove()

                        if not legend_visible or not handles:
                            continue

                        if legend_ncol > 0:
                            ncol = legend_ncol
                        else:
                            ncol = self.optimize_legend_layout(ax, handles, labels, fontsize=fontsize)

                        if hasattr(self, '_ybreak_base_pos') and self._ybreak_base_pos is not None:
                            base_pos = self._ybreak_base_pos
                            x_center = base_pos.x0 + base_pos.width / 2
                            y_top = base_pos.y0 + base_pos.height
                            new_legend = ax.legend(
                                handles, labels,
                                bbox_to_anchor=(x_center, y_top + base_pos.height * 0.15),
                                bbox_transform=self.fig.transFigure,
                                loc='upper center',
                                borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol
                            )
                        else:
                            new_legend = ax.legend(handles, labels, bbox_to_anchor=(0.5, 1.15), loc='upper center',
                                                  borderaxespad=0., frameon=False, fontsize=fontsize, ncol=ncol)

                        if new_legend:
                            new_legend.set_draggable(True)
            except Exception:
                pass

            self.preview_scroll_canvas.config(scrollregion=(0, 0, padded_width, padded_height))

            # Center the plot in the viewport (important when window is not maximized)
            def _center_preview_view():
                try:
                    self.preview_scroll_canvas.update_idletasks()
                    viewport_w = self.preview_scroll_canvas.winfo_width()
                    viewport_h = self.preview_scroll_canvas.winfo_height()

                    max_x = max(padded_width - viewport_w, 0)
                    max_y = max(padded_height - viewport_h, 0)

                    target_x = max((pad_left + (canvas_width / 2)) - (viewport_w / 2), 0)
                    target_y = max((pad_top + (canvas_height / 2)) - (viewport_h / 2), 0)

                    # xview_moveto/yview_moveto expect a fraction of the total scrollregion
                    if padded_width > 0 and max_x > 0:
                        frac_x = min(max(min(target_x, max_x) / padded_width, 0.0), 1.0)
                        self.preview_scroll_canvas.xview_moveto(frac_x)
                    else:
                        self.preview_scroll_canvas.xview_moveto(0)

                    if padded_height > 0 and max_y > 0:
                        frac_y = min(max(min(target_y, max_y) / padded_height, 0.0), 1.0)
                        self.preview_scroll_canvas.yview_moveto(frac_y)
                    else:
                        self.preview_scroll_canvas.yview_moveto(0)
                except Exception as e:
                    print(f"Error centering preview view: {e}")

            self.root.after(80, _center_preview_view)
            
            # Make legends draggable
            for ax in self.fig.axes:
                legend = ax.get_legend()
                if legend:
                    legend.set_draggable(True)
            
            # Add hint about draggable legends
            hint_label = ttk.Label(
                self.canvas_frame,
                text="💡 Drag legends to reposition, then Save PDF",
                font=('TkDefaultFont', 9)
            )
            hint_label.pack(pady=(2, 0))
            
            # Store reference
            self.preview_label = canvas_widget
            
            print(f"Preview: {canvas_width}x{canvas_height}px @ {fig_dpi}dpi (draggable legends enabled)")
            
        except Exception as e:
            print(f"Error creating interactive preview: {e}")
            import traceback
            traceback.print_exc()
            # Fallback to static image
            placeholder = ttk.Label(self.canvas_frame, text=f"Preview error: {e}")
            placeholder.pack(pady=50)

    def save_pdf(self):
        """Save the current figure to PDF at original dimensions, including any legend repositioning done by the user."""
        file_path = filedialog.asksaveasfilename(
            defaultextension='.pdf', 
            filetypes=[("PDF files", "*.pdf")], 
            initialfile='plot_output.pdf'
        )
        if file_path:
            try:
                if hasattr(self, 'fig') and self.fig is not None:
                    # Save directly from the current figure state (preserves dragged legend positions)
                    # PDF is vector format so DPI doesn't affect output quality
                    self.fig.savefig(file_path, format='pdf', bbox_inches='tight')
                    if hasattr(self, 'original_fig_width') and hasattr(self, 'original_fig_height'):
                        print(f"PDF saved to: {file_path} (size: {self.original_fig_width}\" x {self.original_fig_height}\")")
                    else:
                        print(f"PDF saved to: {file_path}")
                elif os.path.exists(self.temp_pdf):
                    # Fallback to temp PDF if figure not available
                    os.replace(self.temp_pdf, file_path)
            except Exception as e:
                print(f"Error saving PDF: {e}")
                from tkinter import messagebox
                messagebox.showerror("Save Error", f"Could not save PDF: {e}")

    def manage_colors_palettes(self):
        window = tk.Toplevel(self.root)
        window.title("Manage Colors & Palettes")
        window.geometry("550x500")
        window.resizable(True, True)
        
        notebook = ttk.Notebook(window)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        colors_tab = ttk.Frame(notebook)
        palettes_tab = ttk.Frame(notebook)
        notebook.add(colors_tab, text="  Single Colors  ")
        notebook.add(palettes_tab, text="  Palettes  ")
        
        # ==================== COLORS TAB ====================
        colors_tab.columnconfigure(0, weight=1)
        colors_tab.rowconfigure(1, weight=1)
        
        ttk.Label(colors_tab, text="Click a color to edit, or use buttons below", font=(None, 9)).grid(row=0, column=0, pady=(10, 5), sticky='w', padx=10)
        
        colors_list_frame = ttk.Frame(colors_tab)
        colors_list_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
        colors_list_frame.columnconfigure(0, weight=1)
        colors_list_frame.rowconfigure(0, weight=1)
        
        color_tree = ttk.Treeview(colors_list_frame, columns=('hex',), show='headings', height=10, selectmode='browse')
        color_tree.heading('#1', text='Color Name', anchor='w')
        color_tree.column('#1', width=200, stretch=tk.YES)
        color_tree.grid(row=0, column=0, sticky='nsew')
        
        color_scrollbar = ttk.Scrollbar(colors_list_frame, orient='vertical', command=color_tree.yview)
        color_scrollbar.grid(row=0, column=1, sticky='ns')
        color_tree.configure(yscrollcommand=color_scrollbar.set)
        
        color_preview_frame = ttk.Frame(colors_tab)
        color_preview_frame.grid(row=2, column=0, pady=10, padx=10, sticky='ew')
        ttk.Label(color_preview_frame, text="Preview:").pack(side='left', padx=(0, 10))
        color_preview = tk.Canvas(color_preview_frame, width=80, height=30, highlightthickness=1, highlightbackground='gray', bg='white')
        color_preview.pack(side='left')
        color_hex_label = ttk.Label(color_preview_frame, text="", font=('Courier', 10))
        color_hex_label.pack(side='left', padx=10)
        
        def get_selected_color_name():
            sel = color_tree.selection()
            if sel:
                return color_tree.item(sel[0])['values'][0] if color_tree.item(sel[0])['values'] else None
            return None
        
        def show_color_preview(event=None):
            color_preview.delete('all')
            sel = color_tree.selection()
            if sel:
                item = color_tree.item(sel[0])
                name = item['tags'][0] if item['tags'] else None
                if name and name in self.custom_colors:
                    hexcode = self.custom_colors[name]
                    color_preview.create_rectangle(2, 2, 78, 28, fill=hexcode, outline='black')
                    color_hex_label.config(text=hexcode)
                else:
                    color_hex_label.config(text="")
            else:
                color_hex_label.config(text="")
        color_tree.bind('<<TreeviewSelect>>', show_color_preview)
        
        def refresh_color_list(select_name=None):
            color_tree.delete(*color_tree.get_children())
            for name, hexcode in self.custom_colors.items():
                iid = color_tree.insert('', 'end', values=(f"  ■  {name}",), tags=(name,))
                color_tree.tag_configure(name, foreground=hexcode)
            if select_name:
                for item in color_tree.get_children():
                    if color_tree.item(item)['tags'][0] == select_name:
                        color_tree.selection_set(item)
                        break
            show_color_preview()
        refresh_color_list()
        
        def add_color():
            result = colorchooser.askcolor(title="Pick a new color")
            if result[1]:
                hexcode = result[1]
                popup = tk.Toplevel(window)
                popup.title("New Color Name")
                popup.geometry("300x100")
                popup.transient(window)
                popup.grab_set()
                ttk.Label(popup, text="Enter a name for this color:").pack(pady=(15, 5))
                name_entry = ttk.Entry(popup, width=30)
                name_entry.pack(pady=5)
                name_entry.focus_set()
                def save():
                    name = name_entry.get().strip()
                    if not name:
                        messagebox.showerror("Error", "Please enter a color name.", parent=popup)
                        return
                    self.custom_colors[name] = hexcode
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    refresh_color_list(select_name=name)
                    popup.destroy()
                name_entry.bind('<Return>', lambda e: save())
                ttk.Button(popup, text="Save", command=save).pack(pady=5)
        
        def add_color_by_hex():
            popup = tk.Toplevel(window)
            popup.title("Add Color by Hex Code")
            popup.geometry("350x180")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text="Color Name:").pack(pady=(15, 5))
            name_entry = ttk.Entry(popup, width=35)
            name_entry.pack(pady=2)
            ttk.Label(popup, text="Hex Code (e.g. #FF5733 or FF5733):").pack(pady=(10, 5))
            hex_entry = ttk.Entry(popup, width=35)
            hex_entry.pack(pady=2)
            hex_preview = tk.Canvas(popup, width=60, height=25, highlightthickness=1, highlightbackground='gray', bg='white')
            hex_preview.pack(pady=5)
            def update_preview(event=None):
                hex_preview.delete('all')
                hexcode = hex_entry.get().strip()
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                try:
                    hex_preview.create_rectangle(2, 2, 58, 23, fill=hexcode, outline='black')
                except:
                    pass
            hex_entry.bind('<KeyRelease>', update_preview)
            name_entry.focus_set()
            def save():
                name = name_entry.get().strip()
                hexcode = hex_entry.get().strip()
                if not name:
                    messagebox.showerror("Error", "Please enter a color name.", parent=popup)
                    return
                if not hexcode:
                    messagebox.showerror("Error", "Please enter a hex code.", parent=popup)
                    return
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                hexcode = self._to_hex(hexcode)
                self.custom_colors[name] = hexcode
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_color_list(select_name=name)
                popup.destroy()
            hex_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Save", command=save).pack(pady=5)
        
        def edit_color():
            sel = color_tree.selection()
            if not sel:
                messagebox.showinfo("Edit Color", "Please select a color to edit.", parent=window)
                return
            name = color_tree.item(sel[0])['tags'][0]
            current_hex = self.custom_colors.get(name, '#000000')
            result = colorchooser.askcolor(color=current_hex, title=f"Edit '{name}'")
            if result[1]:
                self.custom_colors[name] = result[1]
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_color_list(select_name=name)
        
        def edit_color_by_hex():
            sel = color_tree.selection()
            if not sel:
                messagebox.showinfo("Edit by Hex", "Please select a color to edit.", parent=window)
                return
            name = color_tree.item(sel[0])['tags'][0]
            current_hex = self.custom_colors.get(name, '#000000')
            popup = tk.Toplevel(window)
            popup.title(f"Edit '{name}' by Hex Code")
            popup.geometry("350x150")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text=f"Current: {current_hex}", font=(None, 9)).pack(pady=(15, 5))
            ttk.Label(popup, text="New Hex Code:").pack(pady=(5, 2))
            hex_entry = ttk.Entry(popup, width=35)
            hex_entry.insert(0, current_hex)
            hex_entry.pack(pady=2)
            hex_entry.select_range(0, 'end')
            hex_preview = tk.Canvas(popup, width=60, height=25, highlightthickness=1, highlightbackground='gray', bg='white')
            hex_preview.pack(pady=5)
            hex_preview.create_rectangle(2, 2, 58, 23, fill=current_hex, outline='black')
            def update_preview(event=None):
                hex_preview.delete('all')
                hexcode = hex_entry.get().strip()
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                try:
                    hex_preview.create_rectangle(2, 2, 58, 23, fill=hexcode, outline='black')
                except:
                    pass
            hex_entry.bind('<KeyRelease>', update_preview)
            hex_entry.focus_set()
            def save():
                hexcode = hex_entry.get().strip()
                if not hexcode:
                    messagebox.showerror("Error", "Please enter a hex code.", parent=popup)
                    return
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                hexcode = self._to_hex(hexcode)
                self.custom_colors[name] = hexcode
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_color_list(select_name=name)
                popup.destroy()
            hex_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Save", command=save).pack(pady=5)
        
        def rename_color():
            sel = color_tree.selection()
            if not sel:
                messagebox.showinfo("Rename Color", "Please select a color to rename.", parent=window)
                return
            old_name = color_tree.item(sel[0])['tags'][0]
            hexcode = self.custom_colors.get(old_name)
            popup = tk.Toplevel(window)
            popup.title("Rename Color")
            popup.geometry("300x100")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text=f"Enter new name for '{old_name}':").pack(pady=(15, 5))
            name_entry = ttk.Entry(popup, width=30)
            name_entry.insert(0, old_name)
            name_entry.pack(pady=5)
            name_entry.select_range(0, 'end')
            name_entry.focus_set()
            def save():
                new_name = name_entry.get().strip()
                if not new_name:
                    messagebox.showerror("Error", "Please enter a color name.", parent=popup)
                    return
                if new_name != old_name:
                    del self.custom_colors[old_name]
                    self.custom_colors[new_name] = hexcode
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    refresh_color_list(select_name=new_name)
                popup.destroy()
            name_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Save", command=save).pack(pady=5)
        
        def remove_color():
            sel = color_tree.selection()
            if not sel:
                return
            name = color_tree.item(sel[0])['tags'][0]
            if messagebox.askyesno("Remove Color", f"Remove '{name}'?", parent=window):
                del self.custom_colors[name]
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_color_list()
        
        color_tree.bind('<Double-1>', lambda e: edit_color())
        
        color_buttons_frame = ttk.Frame(colors_tab)
        color_buttons_frame.grid(row=3, column=0, pady=10, padx=10, sticky='ew')
        ttk.Button(color_buttons_frame, text="Add", command=add_color).pack(side='left', padx=2)
        ttk.Button(color_buttons_frame, text="Add Hex", command=add_color_by_hex).pack(side='left', padx=2)
        ttk.Button(color_buttons_frame, text="Edit", command=edit_color).pack(side='left', padx=2)
        ttk.Button(color_buttons_frame, text="Edit Hex", command=edit_color_by_hex).pack(side='left', padx=2)
        ttk.Button(color_buttons_frame, text="Rename", command=rename_color).pack(side='left', padx=2)
        ttk.Button(color_buttons_frame, text="Remove", command=remove_color).pack(side='left', padx=2)
        
        # ==================== PALETTES TAB ====================
        palettes_tab.columnconfigure(0, weight=1)
        palettes_tab.rowconfigure(1, weight=1)
        
        ttk.Label(palettes_tab, text="Select a palette, then edit colors below", font=(None, 9)).grid(row=0, column=0, pady=(10, 5), sticky='w', padx=10)
        
        palettes_list_frame = ttk.Frame(palettes_tab)
        palettes_list_frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=5)
        palettes_list_frame.columnconfigure(0, weight=1)
        palettes_list_frame.rowconfigure(0, weight=1)
        
        palette_tree = ttk.Treeview(palettes_list_frame, columns=('name',), show='headings', height=6, selectmode='browse')
        palette_tree.heading('#1', text='Palette Name', anchor='w')
        palette_tree.column('#1', width=200, stretch=tk.YES)
        palette_tree.grid(row=0, column=0, sticky='nsew')
        
        palette_scrollbar = ttk.Scrollbar(palettes_list_frame, orient='vertical', command=palette_tree.yview)
        palette_scrollbar.grid(row=0, column=1, sticky='ns')
        palette_tree.configure(yscrollcommand=palette_scrollbar.set)
        
        palette_preview_frame = ttk.LabelFrame(palettes_tab, text="Palette Colors (click to edit)")
        palette_preview_frame.grid(row=2, column=0, pady=10, padx=10, sticky='ew')
        
        palette_colors_canvas = tk.Canvas(palette_preview_frame, height=50, highlightthickness=0, bg='#f0f0f0')
        palette_colors_canvas.pack(fill='x', padx=5, pady=10)
        
        current_palette_name = tk.StringVar(value="")
        current_palette_colors = []
        
        def get_selected_palette_name():
            sel = palette_tree.selection()
            if sel:
                return palette_tree.item(sel[0])['tags'][0] if palette_tree.item(sel[0])['tags'] else None
            return None
        
        def draw_palette_preview():
            palette_colors_canvas.delete('all')
            name = current_palette_name.get()
            if not name or name not in self.custom_palettes:
                return
            colors = self.custom_palettes[name]
            canvas_width = palette_colors_canvas.winfo_width()
            if canvas_width < 50:
                canvas_width = 500
            n = len(colors)
            if n == 0:
                return
            swatch_w = min(40, (canvas_width - 20) // max(n, 1))
            swatch_h = 35
            start_x = 10
            for i, hexcode in enumerate(colors):
                x0 = start_x + i * (swatch_w + 4)
                x1 = x0 + swatch_w
                y0 = 5
                y1 = y0 + swatch_h
                rect_id = palette_colors_canvas.create_rectangle(x0, y0, x1, y1, fill=hexcode, outline='black', width=1)
                palette_colors_canvas.tag_bind(rect_id, '<Button-1>', lambda e, idx=i: edit_palette_color(idx))
                palette_colors_canvas.tag_bind(rect_id, '<Enter>', lambda e, rid=rect_id: palette_colors_canvas.itemconfig(rid, width=2))
                palette_colors_canvas.tag_bind(rect_id, '<Leave>', lambda e, rid=rect_id: palette_colors_canvas.itemconfig(rid, width=1))
            add_x = start_x + n * (swatch_w + 4)
            add_rect = palette_colors_canvas.create_rectangle(add_x, 5, add_x + swatch_w, 5 + swatch_h, fill='#e0e0e0', outline='gray', dash=(2, 2))
            palette_colors_canvas.create_text(add_x + swatch_w//2, 5 + swatch_h//2, text="+", font=(None, 16, 'bold'), fill='gray')
            palette_colors_canvas.tag_bind(add_rect, '<Button-1>', lambda e: add_color_to_palette())
        
        def edit_palette_color(idx):
            name = current_palette_name.get()
            if not name or name not in self.custom_palettes:
                return
            colors = self.custom_palettes[name]
            if idx >= len(colors):
                return
            current_hex = colors[idx]
            popup = tk.Toplevel(window)
            popup.title(f"Edit Color #{idx+1}")
            popup.geometry("320x200")
            popup.transient(window)
            popup.grab_set()
            preview_canvas = tk.Canvas(popup, width=60, height=30, highlightthickness=1, highlightbackground='gray')
            preview_canvas.pack(pady=(10, 5))
            preview_canvas.create_rectangle(2, 2, 58, 28, fill=current_hex, outline='black')
            hex_frame = ttk.Frame(popup)
            hex_frame.pack(pady=5)
            ttk.Label(hex_frame, text="Hex:").pack(side='left', padx=(0, 5))
            hex_entry = ttk.Entry(hex_frame, width=12)
            hex_entry.insert(0, current_hex)
            hex_entry.pack(side='left')
            def update_preview_from_hex(event=None):
                preview_canvas.delete('all')
                hexcode = hex_entry.get().strip()
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                try:
                    preview_canvas.create_rectangle(2, 2, 58, 28, fill=hexcode, outline='black')
                except:
                    pass
            hex_entry.bind('<KeyRelease>', update_preview_from_hex)
            def save_hex():
                hexcode = hex_entry.get().strip()
                if not hexcode:
                    return
                if not hexcode.startswith('#'):
                    hexcode = '#' + hexcode
                hexcode = self._to_hex(hexcode)
                self.custom_palettes[name][idx] = hexcode
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                draw_palette_preview()
                popup.destroy()
            ttk.Button(hex_frame, text="Apply", command=save_hex).pack(side='left', padx=5)
            def pick_new():
                result = colorchooser.askcolor(color=current_hex, title="Pick new color")
                if result[1]:
                    self.custom_palettes[name][idx] = result[1]
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    draw_palette_preview()
                    popup.destroy()
            def remove_this():
                if len(self.custom_palettes[name]) > 1:
                    del self.custom_palettes[name][idx]
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    draw_palette_preview()
                    popup.destroy()
                else:
                    messagebox.showwarning("Cannot Remove", "Palette must have at least one color.", parent=popup)
            btn_frame = ttk.Frame(popup)
            btn_frame.pack(pady=10)
            ttk.Button(btn_frame, text="Color Picker", command=pick_new).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Remove", command=remove_this).pack(side='left', padx=5)
            ttk.Button(popup, text="Cancel", command=popup.destroy).pack(pady=5)
        
        def add_color_to_palette():
            name = current_palette_name.get()
            if not name or name not in self.custom_palettes:
                messagebox.showinfo("Add Color", "Please select a palette first.", parent=window)
                return
            result = colorchooser.askcolor(title="Add color to palette")
            if result[1]:
                self.custom_palettes[name].append(result[1])
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                draw_palette_preview()
        
        def show_palette_preview(event=None):
            sel = palette_tree.selection()
            if sel:
                name = palette_tree.item(sel[0])['tags'][0]
                current_palette_name.set(name)
            else:
                current_palette_name.set("")
            draw_palette_preview()
        palette_tree.bind('<<TreeviewSelect>>', show_palette_preview)
        palette_colors_canvas.bind('<Configure>', lambda e: draw_palette_preview())
        
        def refresh_palette_list(select_name=None):
            palette_tree.delete(*palette_tree.get_children())
            for name in self.custom_palettes.keys():
                palette_tree.insert('', 'end', values=(name,), tags=(name,))
            if select_name:
                for item in palette_tree.get_children():
                    if palette_tree.item(item)['tags'][0] == select_name:
                        palette_tree.selection_set(item)
                        break
            show_palette_preview()
        refresh_palette_list()
        
        def add_palette():
            popup = tk.Toplevel(window)
            popup.title("New Palette")
            popup.geometry("300x100")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text="Enter a name for the new palette:").pack(pady=(15, 5))
            name_entry = ttk.Entry(popup, width=30)
            name_entry.pack(pady=5)
            name_entry.focus_set()
            def save():
                name = name_entry.get().strip()
                if not name:
                    messagebox.showerror("Error", "Please enter a palette name.", parent=popup)
                    return
                result = colorchooser.askcolor(title="Pick first color for palette")
                if result[1]:
                    self.custom_palettes[name] = [result[1]]
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    refresh_palette_list(select_name=name)
                    popup.destroy()
                else:
                    messagebox.showerror("Error", "Please pick at least one color.", parent=popup)
            name_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Next: Pick Color", command=save).pack(pady=5)
        
        def add_palette_from_hex():
            popup = tk.Toplevel(window)
            popup.title("Create Palette from Hex Codes")
            popup.geometry("450x220")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text="Palette Name:").pack(pady=(15, 5))
            name_entry = ttk.Entry(popup, width=40)
            name_entry.pack(pady=2)
            ttk.Label(popup, text="Hex Codes (comma or space separated):").pack(pady=(10, 5))
            ttk.Label(popup, text="e.g.: #FF5733, #33FF57, #3357FF  or  FF5733 33FF57 3357FF", font=(None, 8)).pack()
            hex_entry = ttk.Entry(popup, width=50)
            hex_entry.pack(pady=5)
            preview_canvas = tk.Canvas(popup, height=30, highlightthickness=0, bg='#f0f0f0')
            preview_canvas.pack(fill='x', padx=20, pady=5)
            def update_preview(event=None):
                preview_canvas.delete('all')
                text = hex_entry.get().strip()
                if not text:
                    return
                parts = [p.strip() for p in text.replace(',', ' ').split() if p.strip()]
                colors = []
                for p in parts:
                    if not p.startswith('#'):
                        p = '#' + p
                    try:
                        colors.append(self._to_hex(p))
                    except:
                        pass
                for i, c in enumerate(colors[:12]):
                    x0 = 10 + i * 32
                    try:
                        preview_canvas.create_rectangle(x0, 2, x0 + 28, 26, fill=c, outline='black')
                    except:
                        pass
            hex_entry.bind('<KeyRelease>', update_preview)
            name_entry.focus_set()
            def save():
                name = name_entry.get().strip()
                text = hex_entry.get().strip()
                if not name:
                    messagebox.showerror("Error", "Please enter a palette name.", parent=popup)
                    return
                if not text:
                    messagebox.showerror("Error", "Please enter at least one hex code.", parent=popup)
                    return
                parts = [p.strip() for p in text.replace(',', ' ').split() if p.strip()]
                colors = []
                for p in parts:
                    if not p.startswith('#'):
                        p = '#' + p
                    try:
                        colors.append(self._to_hex(p))
                    except:
                        pass
                if not colors:
                    messagebox.showerror("Error", "No valid hex codes found.", parent=popup)
                    return
                self.custom_palettes[name] = colors
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_palette_list(select_name=name)
                popup.destroy()
            hex_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Create Palette", command=save).pack(pady=10)
        
        def rename_palette():
            sel = palette_tree.selection()
            if not sel:
                messagebox.showinfo("Rename Palette", "Please select a palette to rename.", parent=window)
                return
            old_name = palette_tree.item(sel[0])['tags'][0]
            colors = self.custom_palettes.get(old_name, [])
            popup = tk.Toplevel(window)
            popup.title("Rename Palette")
            popup.geometry("300x100")
            popup.transient(window)
            popup.grab_set()
            ttk.Label(popup, text=f"Enter new name for '{old_name}':").pack(pady=(15, 5))
            name_entry = ttk.Entry(popup, width=30)
            name_entry.insert(0, old_name)
            name_entry.pack(pady=5)
            name_entry.select_range(0, 'end')
            name_entry.focus_set()
            def save():
                new_name = name_entry.get().strip()
                if not new_name:
                    messagebox.showerror("Error", "Please enter a palette name.", parent=popup)
                    return
                if new_name != old_name:
                    del self.custom_palettes[old_name]
                    self.custom_palettes[new_name] = colors
                    self.save_custom_colors_palettes()
                    self.update_color_palette_dropdowns()
                    refresh_palette_list(select_name=new_name)
                popup.destroy()
            name_entry.bind('<Return>', lambda e: save())
            ttk.Button(popup, text="Save", command=save).pack(pady=5)
        
        def duplicate_palette():
            sel = palette_tree.selection()
            if not sel:
                messagebox.showinfo("Duplicate Palette", "Please select a palette to duplicate.", parent=window)
                return
            old_name = palette_tree.item(sel[0])['tags'][0]
            colors = self.custom_palettes.get(old_name, [])[:]
            new_name = f"{old_name} (copy)"
            i = 2
            while new_name in self.custom_palettes:
                new_name = f"{old_name} (copy {i})"
                i += 1
            self.custom_palettes[new_name] = colors
            self.save_custom_colors_palettes()
            self.update_color_palette_dropdowns()
            refresh_palette_list(select_name=new_name)
        
        def remove_palette():
            sel = palette_tree.selection()
            if not sel:
                return
            name = palette_tree.item(sel[0])['tags'][0]
            if messagebox.askyesno("Remove Palette", f"Remove '{name}'?", parent=window):
                del self.custom_palettes[name]
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
                refresh_palette_list()
        
        palette_buttons_frame = ttk.Frame(palettes_tab)
        palette_buttons_frame.grid(row=3, column=0, pady=10, padx=10, sticky='ew')
        ttk.Button(palette_buttons_frame, text="Add", command=add_palette).pack(side='left', padx=2)
        ttk.Button(palette_buttons_frame, text="From Hex", command=add_palette_from_hex).pack(side='left', padx=2)
        ttk.Button(palette_buttons_frame, text="Rename", command=rename_palette).pack(side='left', padx=2)
        ttk.Button(palette_buttons_frame, text="Duplicate", command=duplicate_palette).pack(side='left', padx=2)
        ttk.Button(palette_buttons_frame, text="Remove", command=remove_palette).pack(side='left', padx=2)
        
        ttk.Button(window, text="Close", command=window.destroy).pack(pady=10)

    def cleanup(self):
        """Clean up resources before exiting the application."""
        try:
            # Close any open matplotlib figures
            if hasattr(self, 'figure'):
                try:
                    import matplotlib.pyplot as plt
                    plt.close(self.figure)
                except Exception as e:
                    print(f"Error closing figure: {e}")
            
            # Clear any data
            if hasattr(self, 'data'):
                if isinstance(self.data, pd.DataFrame):
                    self.data = pd.DataFrame()
            
            # Clear any cached data
            if hasattr(self, 'cached_data'):
                self.cached_data = {}
                
            # Clear any plot elements
            if hasattr(self, 'canvas'):
                try:
                    self.canvas.get_tk_widget().destroy()
                except Exception as e:
                    print(f"Error destroying canvas: {e}")
            
            # Clear any other resources
            if hasattr(self, 'temp_files'):
                for temp_file in self.temp_files:
                    try:
                        if os.path.exists(temp_file):
                            os.unlink(temp_file)
                    except Exception as e:
                        print(f"Error removing temp file {temp_file}: {e}")
            
            # Force garbage collection
            import gc
            gc.collect()
            
            # Destroy the root window if it exists
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.quit()
                self.root.destroy()
            
            # Force exit on macOS
            import os
            import sys
            if sys.platform == 'darwin':
                os._exit(0)
                
        except Exception as e:
            print(f"Error during cleanup: {e}")
            # Still try to exit even if cleanup fails
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.quit()
                self.root.destroy()
            import os
            os._exit(0)

def on_closing(app):
    """Handle window close event."""
    app.cleanup()

if __name__ == '__main__':
    root = tk.Tk()
    app = ExPlotApp(root)
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(app))
    root.mainloop()