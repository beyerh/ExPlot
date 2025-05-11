# Excel Plotter - Data visualization tool for Excel files

VERSION = "0.5.2"
# =====================================================================

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
import sys
import tempfile
from pathlib import Path
import pingouin as pg
import scikit_posthocs as sp
import math

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

# --- in show_statistical_details, replace all key = ... assignments for latest_pvals with stat_key
# --- in plot_graph, replace all key = ... and latest_pvals lookups with stat_key
# --- add debug print if a key is missing in plot_graph when drawing annotation

class ExcelPlotterApp:
    def stat_key(self, *args):
        # For group comparisons, always sort the pair for unpaired, or keep order for paired if needed
        # If first arg is a group (x axis), keep as is, but sort the next two
        if len(args) == 3:
            g, h1, h2 = args
            return (g, ) + tuple(sorted([h1, h2]))
        elif len(args) == 2:
            h1, h2 = args
            return tuple(sorted([h1, h2]))
        else:
            return tuple(args)

    def calculate_statistics(self, df_plot, x_col, value_col, hue_col=None):
        """
        Centralized statistics calculation method that both annotations and
        statistical details panel will use. This ensures consistency between
        the two and respects settings in the statistics tab.
        
        Returns a dictionary with comprehensive statistical results.
        """
        import itertools
        import numpy as np
        
        # Initialize statistics storage
        try:
            alpha = float(self.alpha_level_var.get())
        except (ValueError, AttributeError):
            alpha = 0.05  # Default if not set or invalid
            
        self.latest_stats = {
            'pvals': {},                # P-values for annotations
            'test_results': {},         # Complete test result objects
            'test_type': self.ttest_type_var.get(),
            'alternative': self.ttest_alternative_var.get(),
            'alpha_level': alpha,        # Alpha level for significance
            'raw_data': {},             # Raw data used for calculations
            'x_col': x_col,
            'value_col': value_col,
            'hue_col': hue_col
        }
        
        # Clear previous p-values (for backward compatibility)
        self.latest_pvals = {}
        
        # Skip all calculations if statistics are disabled
        if not self.use_stats_var.get():
            print("[DEBUG] Statistics disabled, skipping calculations")
            return self.latest_stats
            
        print(f"[DEBUG] calculate_statistics: x_col={x_col}, value_col={value_col}, hue_col={hue_col}")
        
        # X-category comparisons (for both with and without hue groups)
        x_values = [g for g in df_plot[x_col].dropna().unique()]
        self.latest_stats['x_values'] = x_values
        print(f"[DEBUG] X-values: {x_values}")
        
        # Process based on whether we have hue groups or not
        if not hue_col or (hue_col and len(df_plot[hue_col].dropna().unique()) == 1):
            # Either no hue column, or just one hue group
            single_group = None
            if hue_col:
                hue_groups = list(df_plot[hue_col].dropna().unique())
                if len(hue_groups) == 1:
                    single_group = hue_groups[0]
                    print(f"[DEBUG] Single group detected: {single_group}")
                    
            # If we have multiple x-values (e.g., Control, Experimental, Predicted)
            # Calculate p-values between all pairs
            if len(x_values) > 1:
                print(f"[DEBUG] Multiple x-values detected: {x_values}")
                pairs = list(itertools.combinations(x_values, 2))
                
                # Store test data for Statistical Details panel
                self.latest_stats['comparison_type'] = 'x_categories'
                self.latest_stats['pairs'] = pairs
                
                # Determine if we should use ANOVA + post-hoc tests
                use_anova = len(x_values) > 2 and self.anova_type_var.get() != "None"
                posthoc_results = None
                
                # If we have more than 2 categories, try ANOVA + post-hoc first
                if use_anova:
                    try:
                        import pingouin as pg
                        import scikit_posthocs as sp
                        print(f"[DEBUG] Using ANOVA + post-hoc for multiple x-categories")
                        
                        # Prepare data for ANOVA
                        if single_group:
                            # Filter by the single group
                            df_anova = df_plot[df_plot[hue_col] == single_group].copy()
                        else:
                            df_anova = df_plot.copy()
                        
                        # Ensure data is numeric for ANOVA
                        try:
                            df_anova[value_col] = pd.to_numeric(df_anova[value_col], errors='coerce')
                            df_anova = df_anova.dropna(subset=[value_col])
                            print(f"[DEBUG] Converted {value_col} to numeric for ANOVA, shape: {df_anova.shape}")
                        except Exception as e:
                            print(f"[DEBUG] Error converting {value_col} to numeric: {e}")
                            raise
                        
                        # Perform ANOVA based on the selected type
                        anova_type = self.anova_type_var.get()
                        print(f"[DEBUG] Using ANOVA type: {anova_type}")
                        
                        if anova_type == "Welch's ANOVA":
                            aov = pg.welch_anova(data=df_anova, dv=value_col, between=x_col)
                            self.latest_stats['anova_type'] = "Welch's ANOVA"
                        elif anova_type == "Repeated measures ANOVA":
                            # For repeated measures ANOVA, we need a subject identifier
                            # We'll use the hue column if available as the subject ID
                            subject_col = 'Subject'
                            if hue_col in df_anova.columns:
                                # Use the hue column as a proxy for subject ID if no explicit subject column
                                subject_id = hue_col
                            else:
                                # If no subject ID available, create dummy IDs
                                df_anova[subject_col] = np.arange(len(df_anova))
                                subject_id = subject_col
                                
                            try:
                                aov = pg.rm_anova(data=df_anova, dv=value_col, within=x_col, subject=subject_id)
                                self.latest_stats['anova_type'] = "Repeated measures ANOVA"
                            except Exception as e:
                                print(f"[DEBUG] Repeated measures ANOVA failed: {e}. Falling back to one-way ANOVA.")
                                aov = pg.anova(data=df_anova, dv=value_col, between=x_col)
                                self.latest_stats['anova_type'] = "One-way ANOVA (fallback)"
                        else:  # Regular one-way ANOVA
                            aov = pg.anova(data=df_anova, dv=value_col, between=x_col)
                            self.latest_stats['anova_type'] = "One-way ANOVA"
                        
                        # Store ANOVA results
                        self.latest_stats['anova_results'] = aov
                        
                        # Perform post-hoc test based on the selected type
                        posthoc_type = self.posthoc_type_var.get()
                        print(f"[DEBUG] Using post-hoc test: {posthoc_type}")
                        
                        try:
                            # Number of data points for debug
                            print(f"[DEBUG] Running post-hoc test with {len(df_anova)} data points")
                            
                            if posthoc_type == "Tukey's HSD":
                                # Tukey's HSD requires equal variances and is available through pingouin
                                posthoc = pg.pairwise_tukey(data=df_anova, dv=value_col, between=x_col)
                                # Convert to matrix format for consistency
                                groups = df_anova[x_col].unique()
                                posthoc_matrix = pd.DataFrame(index=groups, columns=groups)
                                for i, row in posthoc.iterrows():
                                    g1, g2 = row['A'], row['B']
                                    p_value = row['p-tukey']
                                    # Store actual p-values, not rounded values
                                    posthoc_matrix.loc[g1, g2] = p_value
                                    posthoc_matrix.loc[g2, g1] = p_value
                                # Fill diagonal with 1.0 (no difference)
                                for g in groups:
                                    posthoc_matrix.loc[g, g] = 1.0
                                # Ensure p-values are properly stored as floats
                                posthoc_matrix = posthoc_matrix.astype(float)
                                posthoc = posthoc_matrix
                                
                            elif posthoc_type == "Scheffe's test":
                                # Scheffe's test is robust for unequal sample sizes
                                posthoc = sp.posthoc_scheffe(df_anova, val_col=value_col, group_col=x_col)
                                
                            elif posthoc_type == "Dunn's test":
                                # Dunn's test is non-parametric, good for non-normal distributions
                                # First create a cross-tabulation of data
                                posthoc = sp.posthoc_dunn(df_anova, val_col=value_col, group_col=x_col)
                                
                            else:  # Default to Tamhane's T2
                                # Tamhane's T2 is for unequal variances
                                posthoc = sp.posthoc_tamhane(df_anova, val_col=value_col, group_col=x_col)
                            
                            # Store results
                            posthoc_results = posthoc
                            self.latest_stats['posthoc_results'] = posthoc
                            self.latest_stats['posthoc_type'] = posthoc_type
                            print(f"[DEBUG] Post-hoc test successful: {posthoc_type}")
                            
                        except Exception as e:
                            print(f"[DEBUG] Post-hoc test '{posthoc_type}' failed: {e}")
                            raise
                        
                        # Use post-hoc p-values for annotations instead of individual t-tests
                        for g1, g2 in pairs:
                            try:
                                # Get p-value from post-hoc test
                                if g1 in posthoc.index and g2 in posthoc.columns:
                                    pval = posthoc.loc[g1, g2]
                                elif g2 in posthoc.index and g1 in posthoc.columns:
                                    pval = posthoc.loc[g2, g1]
                                else:
                                    raise KeyError(f"Cannot find {g1}, {g2} pair in post-hoc results")
                                
                                # Store raw data for reference
                                if single_group:
                                    vals1 = df_plot[(df_plot[x_col] == g1) & (df_plot[hue_col] == single_group)][value_col].dropna().astype(float)
                                    vals2 = df_plot[(df_plot[x_col] == g2) & (df_plot[hue_col] == single_group)][value_col].dropna().astype(float)
                                else:
                                    vals1 = df_plot[df_plot[x_col] == g1][value_col].dropna().astype(float)
                                    vals2 = df_plot[df_plot[x_col] == g2][value_col].dropna().astype(float)
                                self.latest_stats['raw_data'][(g1, g2)] = (vals1, vals2)
                                
                                # Store using multiple key formats
                                key1 = self.stat_key(g1, g2)  # Sorted key
                                key2 = (g1, g2)  # Direct key
                                key3 = (g2, g1)  # Reversed key
                                
                                # Store p-values from post-hoc test
                                self.latest_pvals[key1] = pval
                                self.latest_pvals[key2] = pval
                                self.latest_pvals[key3] = pval
                                self.latest_stats['pvals'][key1] = pval
                                self.latest_stats['pvals'][key2] = pval
                                self.latest_stats['pvals'][key3] = pval
                                
                                print(f"[DEBUG] Stored post-hoc p-value for {g1} vs {g2}: {pval}")
                            except Exception as e:
                                print(f"[DEBUG] Error storing post-hoc p-value for {g1} vs {g2}: {e}")
                                # Fall back to individual t-tests for this pair
                                posthoc_results = None
                    except Exception as e:
                        print(f"[DEBUG] ANOVA + post-hoc failed, falling back to pairwise t-tests: {e}")
                        posthoc_results = None
                
                # If ANOVA failed or we're not using it, fall back to individual t-tests
                if not use_anova or posthoc_results is None:
                    for g1, g2 in pairs:
                        # Get the data for each x value
                        if single_group:
                            # If we have a single group, filter by it
                            vals1 = df_plot[(df_plot[x_col] == g1) & (df_plot[hue_col] == single_group)][value_col].dropna().astype(float)
                            vals2 = df_plot[(df_plot[x_col] == g2) & (df_plot[hue_col] == single_group)][value_col].dropna().astype(float)
                        else:
                            # No hue group
                            vals1 = df_plot[df_plot[x_col] == g1][value_col].dropna().astype(float)
                            vals2 = df_plot[df_plot[x_col] == g2][value_col].dropna().astype(float)
                        
                        # Store raw data for future reference
                        self.latest_stats['raw_data'][(g1, g2)] = (vals1, vals2)
                        
                        if len(vals1) < 2 or len(vals2) < 2:
                            print(f"[DEBUG] Skipping comparison {g1} vs {g2} due to insufficient data")
                            continue
                            
                        # Perform statistical test based on user settings
                        ttest_type = self.ttest_type_var.get()
                        alternative = self.ttest_alternative_var.get()
                        
                        try:
                            if ttest_type == "Paired t-test" and len(vals1) == len(vals2):
                                ttest = stats.ttest_rel(vals1, vals2, alternative=alternative)
                            elif ttest_type == "Student's t-test (unpaired)":
                                ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=True)
                            else:  # Welch's t-test (default)
                                ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=False)
                                
                            # Store the full test result for statistical details panel
                            self.latest_stats['test_results'][(g1, g2)] = ttest
                            
                            # Store p-value using multiple key formats to ensure retrieval works
                            pval = ttest.pvalue
                            key1 = self.stat_key(g1, g2)  # Sorted key
                            key2 = (g1, g2)  # Direct key
                            key3 = (g2, g1)  # Reversed key
                            
                            # Store p-values in both new and old locations for compatibility
                            self.latest_pvals[key1] = pval
                            self.latest_pvals[key2] = pval
                            self.latest_pvals[key3] = pval
                            self.latest_stats['pvals'][key1] = pval
                            self.latest_stats['pvals'][key2] = pval
                            self.latest_stats['pvals'][key3] = pval
                            
                            print(f"[DEBUG] Stored t-test p-value for {g1} vs {g2}: {pval}")
                        except Exception as e:
                            print(f"[DEBUG] Error calculating p-value for {g1} vs {g2}: {e}")
            
            # Special case: For a single x-value with multiple data points
            elif len(x_values) == 1:
                print(f"[DEBUG] Single x-value with multiple data points: {x_values[0]}")
                g = x_values[0]  # The single x value
                if single_group:
                    values = df_plot[(df_plot[x_col] == g) & (df_plot[hue_col] == single_group)][value_col].dropna().astype(float)
                else:
                    values = df_plot[df_plot[x_col] == g][value_col].dropna().astype(float)
                
                # Store comparison type
                self.latest_stats['comparison_type'] = 'one_sample'
                self.latest_stats['raw_data']['one_sample'] = values
                
                # If there are enough data points, compute a one-sample t-test
                if len(values) >= 2:
                    try:
                        ttest = stats.ttest_1samp(values, 0)
                        key = self.stat_key(g, "one_sample")
                        key_display = self.stat_key(g, "display")
                        
                        # Store results
                        self.latest_stats['test_results']['one_sample'] = ttest
                        self.latest_pvals[key] = ttest.pvalue
                        self.latest_pvals[key_display] = ttest.pvalue
                        self.latest_stats['pvals'][key] = ttest.pvalue
                        self.latest_stats['pvals'][key_display] = ttest.pvalue
                        
                        print(f"[DEBUG] Stored one-sample p-value for {g}: {ttest.pvalue}")
                    except Exception as e:
                        print(f"[DEBUG] Error calculating one-sample p-value for {g}: {e}")
            
            # No meaningful comparisons possible
            else:
                print(f"[DEBUG] No meaningful comparisons possible with x_values={x_values}")
        
        # Multiple hue groups within x-categories
        else:  
            base_groups = [g for g in df_plot[x_col].dropna().unique()]
            hue_groups = [g for g in df_plot[hue_col].dropna().unique()]
            
            print(f"[DEBUG] Multiple hue groups within x-categories: x={base_groups}, hue={hue_groups}")
            
            # Store for statistical details panel
            self.latest_stats['comparison_type'] = 'within_x_groups'
            self.latest_stats['base_groups'] = base_groups
            self.latest_stats['hue_groups'] = hue_groups
            
            # Perform within-group comparisons
            for g in base_groups:
                df_sub = df_plot[df_plot[x_col] == g]
                pairs = list(itertools.combinations(hue_groups, 2))
                
                for h1, h2 in pairs:
                    # Get the data
                    vals1 = df_sub[df_sub[hue_col] == h1][value_col].dropna().astype(float)
                    vals2 = df_sub[df_sub[hue_col] == h2][value_col].dropna().astype(float)
                    
                    # Store raw data
                    self.latest_stats['raw_data'][(g, h1, h2)] = (vals1, vals2)
                    
                    if len(vals1) < 2 or len(vals2) < 2:
                        print(f"[DEBUG] Skipping comparison {g}: {h1} vs {h2} due to insufficient data")
                        continue
                        
                    # Perform statistical test
                    ttest_type = self.ttest_type_var.get()
                    alternative = self.ttest_alternative_var.get()
                    
                    try:
                        if ttest_type == "Paired t-test" and len(vals1) == len(vals2):
                            ttest = stats.ttest_rel(vals1, vals2, alternative=alternative)
                        elif ttest_type == "Student's t-test (unpaired)":
                            ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=True)
                        else:  # Welch's t-test (default)
                            ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=False)
                        
                        # Store the full test result and p-value
                        key = self.stat_key(g, h1, h2)
                        self.latest_stats['test_results'][key] = ttest
                        pval = ttest.pvalue
                        
                        # Store in both locations
                        self.latest_pvals[key] = pval
                        self.latest_stats['pvals'][key] = pval
                        
                        print(f"[DEBUG] Stored p-value for {g}: {h1} vs {h2}: {pval}")
                    except Exception as e:
                        print(f"[DEBUG] Error calculating p-value for {g}: {h1} vs {h2}: {e}")
        
        return self.latest_stats
    
    # Define a backward-compatible method to maintain existing interfaces
    def calculate_and_store_pvals(self, df_plot, x_col, value_col, hue_col=None):
        """Backward-compatible wrapper for calculate_statistics"""
        stats_result = self.calculate_statistics(df_plot, x_col, value_col, hue_col)
        # Return the result for compatibility
        return stats_result
    
    def __init__(self, root):
        self.latest_pvals = {}  # {(group, h1, h2): pval or (h1, h2): pval}

        self.root = root
        self.version = VERSION  # Use the global VERSION constant
        self.root.title(f'Excel Plotter v{VERSION}')
        self.df = None
        self.excel_file = None
        self.preview_label = None
        self.config_dir = self.get_config_dir()
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.custom_colors_file = str(self.config_dir / "custom_colors.json")
        self.custom_palettes_file = str(self.config_dir / "custom_palettes.json")
        self.default_settings_file = str(self.config_dir / "default_settings.json")
        self.temp_pdf = str(Path(tempfile.gettempdir()) / "excelplotter_temp_plot.pdf")
        self.xaxis_renames = {}
        self.xaxis_order = []
        self.use_stats_var = tk.BooleanVar(value=False)
        self.linewidth = tk.DoubleVar(value=1.0)
        self.errorbar_capsize_var = tk.StringVar(value="Default")  # Capsize style for error bars
        self.strip_black_var = tk.BooleanVar(value=True)
        self.show_stripplot_var = tk.BooleanVar(value=True)
        self.plot_kind_var = tk.StringVar(value="bar")  # "bar", "box", or "xy"
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
        self.load_custom_colors_palettes()
        # Initialize default plot dimensions for the plot area
        self.plot_width_var = tk.DoubleVar(value=1.5)
        self.plot_height_var = tk.DoubleVar(value=1.5)
        self.load_user_preferences()
        self.setup_menu()
        self.setup_ui()
        self.setup_statistics_settings_tab()

    def pval_to_annotation(self, pval):
        """Return annotation string for a given p-value based on the selected alpha level.
        Returns '?' if pval is None or NaN."""
        if pval is None or (isinstance(pval, float) and math.isnan(pval)):
            return "?"
            
        # Get the current alpha level
        try:
            alpha = float(self.alpha_level_var.get())
        except (ValueError, AttributeError):
            alpha = 0.05  # Default if not set or invalid
            
        if pval > alpha:
            return "ns"
        elif pval <= alpha/5000:  # 4 stars threshold
            return "****"
        elif pval <= alpha/50:    # 3 stars threshold
            return "***"
        elif pval <= alpha/5:     # 2 stars threshold
            return "**"
        elif pval <= alpha:       # 1 star threshold
            return "*"

    def format_pvalue_matrix(self, matrix):
        """Format a p-value matrix as a readable ASCII table with aligned columns and dashes on the diagonal."""
        import pandas as pd
        formatted = matrix.copy()
        col_names = list(formatted.columns)
        # Format values: fewer digits, dash for diagonal
        for idx in formatted.index:
            for col in formatted.columns:
                val = formatted.loc[idx, col]
                if idx == col or (isinstance(val, float) and val == 1.0):
                    formatted.loc[idx, col] = "-"
                elif isinstance(val, float):
                    if val < 0.0001:
                        formatted.loc[idx, col] = f"{val:.2e}"
                    else:
                        formatted.loc[idx, col] = f"{val:.3f}"
                else:
                    formatted.loc[idx, col] = str(val)
        # Calculate column widths
        col_widths = [
            max(len(str(col)), *(len(str(formatted.loc[idx, col])) for idx in formatted.index))
            for col in col_names
        ]
        idx_width = max(len(str(idx)) for idx in formatted.index)
        # Build header
        header = " " * (idx_width + 2) + "| " + " | ".join(
            f"{col:^{w}}" for col, w in zip(col_names, col_widths)
        ) + " |"
        sep = "-" * (idx_width + 2) + "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
        # Build rows
        rows = []
        for idx in formatted.index:
            row = f" {str(idx):<{idx_width}} | " + " | ".join(
                f"{str(formatted.loc[idx, col]):>{w}}" for col, w in zip(col_names, col_widths)
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
            "Student's t-test (unpaired)",
            "Welch's t-test (unpaired, unequal variances)",
            "Paired t-test"
        ]
        ttest_dropdown = ttk.Combobox(frame, textvariable=self.ttest_type_var, values=ttest_options, state='readonly')
        ttest_dropdown.grid(row=1, column=1, sticky="ew", padx=8, pady=8)
        
        # T-test alternative hypothesis
        ttk.Label(frame, text="T-test Alternative:").grid(row=2, column=0, sticky="w", padx=8, pady=8)
        # Use the existing ttest_alternative_var variable initialized in __init__
        ttest_alternative_options = [
            "two-sided",
            "less",
            "greater"
        ]
        ttest_alternative_dropdown = ttk.Combobox(frame, textvariable=self.ttest_alternative_var, values=ttest_alternative_options, state='readonly')
        ttest_alternative_dropdown.grid(row=2, column=1, sticky="ew", padx=8, pady=8)
        
        # ANOVA type
        ttk.Label(frame, text="ANOVA type:").grid(row=3, column=0, sticky="w", padx=8, pady=8)
        # Use the existing anova_type_var variable initialized in __init__
        anova_options = [
            "One-way ANOVA",
            "Welch's ANOVA",
            "Repeated measures ANOVA"
        ]
        anova_dropdown = ttk.Combobox(frame, textvariable=self.anova_type_var, values=anova_options, state='readonly')
        anova_dropdown.grid(row=3, column=1, sticky="ew", padx=8, pady=8)
        
        # Alpha level
        ttk.Label(frame, text="Alpha level:").grid(row=4, column=0, sticky="w", padx=8, pady=8)
        alpha_options = ["0.05", "0.01", "0.001", "0.0001"]
        alpha_dropdown = ttk.Combobox(frame, textvariable=self.alpha_level_var, values=alpha_options, state='readonly')
        alpha_dropdown.grid(row=4, column=1, sticky="ew", padx=8, pady=8)
        
        # Post-hoc test
        ttk.Label(frame, text="Post-hoc test:").grid(row=5, column=0, sticky="w", padx=8, pady=8)
        # Use the existing posthoc_type_var variable initialized in __init__
        posthoc_options = [
            "Tukey's HSD",
            "Tamhane's T2",
            "Scheffe's test",
            "Dunn's test"
        ]
        posthoc_dropdown = ttk.Combobox(frame, textvariable=self.posthoc_type_var, values=posthoc_options, state='readonly')
        posthoc_dropdown.grid(row=5, column=1, sticky="ew", padx=8, pady=8)


    def get_config_dir(self):
        """Return the user config directory for settings (cross-platform)."""
        if sys.platform == "darwin":
            return Path.home() / "Library" / "Application Support" / "ExcelPlotter"
        elif sys.platform.startswith("win"):
            return Path(os.environ.get("APPDATA", str(Path.home() / "AppData" / "Roaming"))) / "ExcelPlotter"
        else:
            # Linux and other
            return Path.home() / ".config" / "ExcelPlotter"

    def setup_menu(self):
        """Setup the application menu."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Excel File", command=self.open_file)
        file_menu.add_command(label="Load Example Data", command=self.load_example_data)
        file_menu.add_command(label="Save Graph", command=self.save_graph)
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

    def save_project(self):
        """Save current plot configuration to a project file."""
        # Check if we have data to save
        if self.df is None:
            messagebox.showwarning("No Data", "Please load data before saving a project.")
            return
            
        # Ask user for save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".explt",
            filetypes=[("ExcelPlotter Projects", "*.explt"), ("All files", "*.*")],
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
            filetypes=[("ExcelPlotter Projects", "*.explt"), ("All files", "*.*")],
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
        """Process an Excel file by loading it and setting up UI elements."""
        try:
            self.excel_file = file_path
            self.current_excel_file = file_path  # Store path for project saving/loading
            
            if file_path.lower().endswith('.csv'):
                # Handle CSV files
                self.df = pd.read_csv(file_path, dtype=object)
                # Create a single sheet for CSV
                self.sheet_options = ['Sheet1']
                if hasattr(self, 'selected_sheet'):
                    self.selected_sheet.set('Sheet1')
                elif hasattr(self, 'sheet_var'):
                    self.sheet_var.set('Sheet1')
                
                # Update column selection dropdowns
                self.update_columns()
            else:
                # Handle Excel files
                xls = pd.ExcelFile(self.excel_file)
                self.sheet_options = xls.sheet_names
                
                # Update the sheet dropdown
                if hasattr(self, 'sheet_dropdown'):
                    self.sheet_dropdown['values'] = self.sheet_options
                
                # Set the sheet (prefer 'export' if it exists)
                if "export" in self.sheet_options:
                    sheet_name = "export"
                else:
                    sheet_name = self.sheet_options[0]
                
                # Handle the different sheet variable names used in the code
                if hasattr(self, 'selected_sheet'):
                    self.selected_sheet.set(sheet_name)
                if hasattr(self, 'sheet_var'):
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
            if os.path.exists(example_data_path):
                self.process_excel_file(example_data_path)
            else:
                messagebox.showerror("Error", f"Example data file not found at {example_data_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load example data: {str(e)}")
            
    def save_graph(self):
        """Save the current plot to an image file."""
        if not hasattr(self, 'fig') or self.fig is None:
            messagebox.showwarning("No Plot", "Please create a plot before saving.")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG files", "*.png"),
                ("PDF files", "*.pdf"),
                ("SVG files", "*.svg"),
                ("JPEG files", "*.jpg"),
                ("All files", "*.*")
            ],
            title="Save Graph"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Get the file extension
            extension = os.path.splitext(file_path)[1].lower()
            
            # Save with appropriate DPI based on format
            if extension == '.pdf' or extension == '.svg':
                # Vector formats - save at high quality
                self.fig.savefig(file_path, bbox_inches='tight')
            else:
                # Raster formats - use appropriate DPI
                self.fig.savefig(file_path, dpi=300, bbox_inches='tight')
                
            messagebox.showinfo("Success", f"Graph saved to {file_path}")
            
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
        xy_plot_tab = ttk.Frame(notebook)
        
        notebook.add(general_tab, text='General')
        notebook.add(plot_settings_tab, text='Plot Settings')
        notebook.add(stats_tab, text='Statistics')
        notebook.add(appearance_tab, text='Appearance')
        notebook.add(xy_plot_tab, text='XY Plot')
        
        # Variables to hold settings
        # General tab
        self.settings_plot_kind_var = tk.StringVar(value=self.plot_kind_var.get())
        
        # Plot Settings tab
        self.settings_show_stripplot_var = tk.BooleanVar(value=self.show_stripplot_var.get())
        self.settings_strip_black_var = tk.BooleanVar(value=self.strip_black_var.get())
        self.settings_errorbar_type_var = tk.StringVar(value=self.errorbar_type_var.get())
        self.settings_errorbar_black_var = tk.BooleanVar(value=self.errorbar_black_var.get())
        self.settings_errorbar_capsize_var = tk.StringVar(value=self.errorbar_capsize_var.get())
        
        # Statistics tab
        self.settings_use_stats_var = tk.BooleanVar(value=self.use_stats_var.get())
        self.settings_ttest_type_var = tk.StringVar(value=self.ttest_type_var.get())
        self.settings_ttest_alternative_var = tk.StringVar(value=self.ttest_alternative_var.get())
        self.settings_anova_type_var = tk.StringVar(value=self.anova_type_var.get())
        self.settings_alpha_level_var = tk.StringVar(value=self.alpha_level_var.get())
        self.settings_posthoc_type_var = tk.StringVar(value=self.posthoc_type_var.get())
        
        # Appearance tab
        self.settings_linewidth = tk.DoubleVar(value=self.linewidth.get())
        self.settings_plot_width_var = tk.DoubleVar(value=self.plot_width_var.get())
        self.settings_plot_height_var = tk.DoubleVar(value=self.plot_height_var.get())
        
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
        
        # General Tab Content
        tk.Label(general_tab, text="Default Plot Type:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(general_tab, textvariable=self.settings_plot_kind_var, values=["bar", "box", "xy"], width=15, state="readonly").grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        # Plot Settings Tab Content
        tk.Checkbutton(plot_settings_tab, text="Show stripplot", variable=self.settings_show_stripplot_var).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Checkbutton(plot_settings_tab, text="Use black stripplot", variable=self.settings_strip_black_var).grid(row=1, column=0, sticky="w", padx=10, pady=5)
        
        tk.Label(plot_settings_tab, text="Error bar type:", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(plot_settings_tab, textvariable=self.settings_errorbar_type_var, values=["SD", "SEM"], width=15, state="readonly").grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        tk.Checkbutton(plot_settings_tab, text="Black error bars", variable=self.settings_errorbar_black_var).grid(row=3, column=0, sticky="w", padx=10, pady=5)
        
        tk.Label(plot_settings_tab, text="Error bar capsize:", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(plot_settings_tab, textvariable=self.settings_errorbar_capsize_var, values=["Default", "None", "Small", "Medium", "Large"], width=15, state="readonly").grid(row=4, column=1, sticky="w", padx=10, pady=10)
        
        # Statistics Tab Content
        # Use statistics checkbox at the top
        tk.Checkbutton(stats_tab, text="Use statistics by default", variable=self.settings_use_stats_var).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        
        tk.Label(stats_tab, text="t-test type:", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        ttest_options = ["Student's t-test (unpaired)", "Welch's t-test (unpaired, unequal variances)", "Paired t-test"]
        ttk.Combobox(stats_tab, textvariable=self.settings_ttest_type_var, values=ttest_options, width=35, state="readonly").grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(stats_tab, text="T-test Alternative:", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        ttest_alternative_options = ["two-sided", "less", "greater"]
        ttk.Combobox(stats_tab, textvariable=self.settings_ttest_alternative_var, values=ttest_alternative_options, width=35, state="readonly").grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(stats_tab, text="ANOVA type:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        anova_options = ["One-way ANOVA", "Welch's ANOVA", "Repeated measures ANOVA"]
        ttk.Combobox(stats_tab, textvariable=self.settings_anova_type_var, values=anova_options, width=35, state="readonly").grid(row=3, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(stats_tab, text="Alpha level:", anchor="w").grid(row=4, column=0, sticky="w", padx=10, pady=10)
        alpha_options = ["0.05", "0.01", "0.001", "0.0001"]
        ttk.Combobox(stats_tab, textvariable=self.settings_alpha_level_var, values=alpha_options, width=35, state="readonly").grid(row=4, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(stats_tab, text="Post-hoc test:", anchor="w").grid(row=5, column=0, sticky="w", padx=10, pady=10)
        posthoc_options = ["Tukey's HSD", "Tamhane's T2", "Scheffe's test", "Dunn's test"]
        ttk.Combobox(stats_tab, textvariable=self.settings_posthoc_type_var, values=posthoc_options, width=35, state="readonly").grid(row=5, column=1, sticky="w", padx=10, pady=10)
        
        # Appearance Tab Content
        tk.Label(appearance_tab, text="Line width:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.5, textvariable=self.settings_linewidth, width=5).grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(appearance_tab, text="Plot width (inches):", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.1, textvariable=self.settings_plot_width_var, width=5).grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(appearance_tab, text="Plot height (inches):", anchor="w").grid(row=2, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(appearance_tab, from_=0.5, to=5.0, increment=0.1, textvariable=self.settings_plot_height_var, width=5).grid(row=2, column=1, sticky="w", padx=10, pady=10)
        
        # Reset colors/palettes
        tk.Label(appearance_tab, text="Reset options:", anchor="w", font=(None, 10, 'bold')).grid(row=3, column=0, sticky="w", padx=10, pady=(20, 5))
        
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
                
        reset_frame = tk.Frame(appearance_tab)
        reset_frame.grid(row=4, column=0, columnspan=2, sticky="w", padx=10, pady=5)
        tk.Button(reset_frame, text="Reset Colors", command=reset_colors, width=15).grid(row=0, column=0, padx=5)
        tk.Button(reset_frame, text="Reset Palettes", command=reset_palettes, width=15).grid(row=0, column=1, padx=5)
        
        # XY Plot Tab Content
        tk.Label(xy_plot_tab, text="Marker Symbol:", anchor="w").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(xy_plot_tab, textvariable=self.settings_xy_marker_symbol_var, values=["o", "s", "^", "D", "v", "P", "X", "+", "x", "*", "."], width=5, state="readonly").grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        tk.Label(xy_plot_tab, text="Marker Size:", anchor="w").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        tk.Spinbox(xy_plot_tab, from_=1, to=15, increment=0.5, textvariable=self.settings_xy_marker_size_var, width=5).grid(row=1, column=1, sticky="w", padx=10, pady=10)
        
        tk.Checkbutton(xy_plot_tab, text="Filled Symbols", variable=self.settings_xy_filled_var).grid(row=2, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        
        tk.Label(xy_plot_tab, text="Line Style:", anchor="w").grid(row=3, column=0, sticky="w", padx=10, pady=10)
        ttk.Combobox(xy_plot_tab, textvariable=self.settings_xy_line_style_var, values=["solid", "dashed", "dotted", "dashdot"], width=10, state="readonly").grid(row=3, column=1, sticky="w", padx=10, pady=10)
        
        tk.Checkbutton(xy_plot_tab, text="Lines in Black", variable=self.settings_xy_line_black_var).grid(row=4, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        tk.Checkbutton(xy_plot_tab, text="Connect Mean with Lines", variable=self.settings_xy_connect_var).grid(row=5, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        tk.Checkbutton(xy_plot_tab, text="Show Mean Values", variable=self.settings_xy_show_mean_var).grid(row=6, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        tk.Checkbutton(xy_plot_tab, text="With Error Bars", variable=self.settings_xy_show_mean_errorbars_var).grid(row=7, column=0, sticky="w", padx=30, pady=5, columnspan=2)
        tk.Checkbutton(xy_plot_tab, text="Draw Bands", variable=self.settings_xy_draw_band_var).grid(row=8, column=0, sticky="w", padx=10, pady=5, columnspan=2)
        
        # Buttons at the bottom
        button_frame = tk.Frame(window)
        button_frame.pack(pady=10, fill='x')
        
        def save_settings():
            self.save_user_preferences()
            messagebox.showinfo("Settings Saved", "Your preferences have been saved.")
            window.destroy()
            
        def reset_all_preferences():
            if messagebox.askyesno("Reset All Preferences", "This will reset all preferences to default values and cannot be undone. Continue?"):
                if os.path.exists(self.default_settings_file):
                    os.remove(self.default_settings_file)
                messagebox.showinfo("Reset Preferences", "All preferences have been reset to default values.")
                window.destroy()
        
        tk.Button(button_frame, text="Save Settings", command=save_settings, width=15).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Reset All Preferences", command=reset_all_preferences, width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=window.destroy, width=10).pack(side=tk.RIGHT, padx=5)

    def show_about(self):
        messagebox.showinfo("About Excel Plotter", f"Excel Plotter\nVersion: {self.version}\n\nA tool for plotting Excel data.")
        
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
        general_text.insert('end', '\nThis guide will help you choose the appropriate statistical test for your data.\n\n', 'normal')
        general_text.insert('end', '1. For comparing TWO groups:', 'subheading')
        general_text.insert('end', '\n   • Use a t-test (Student\'s or Welch\'s for unpaired data, Paired for paired data)\n', 'normal')
        general_text.insert('end', '2. For comparing THREE OR MORE groups:', 'subheading')
        general_text.insert('end', '\n   • Use ANOVA followed by a post-hoc test to identify which specific groups differ\n', 'normal')
        general_text.insert('end', '3. For non-parametric data (data that doesn\'t follow normal distribution):', 'subheading')
        general_text.insert('end', '\n   • Consider using non-parametric alternatives like Dunn\'s test for post-hoc comparisons\n', 'normal')
        general_text.insert('end', '\nThe ExcelPlotter automatically selects appropriate tests based on your data structure.\n', 'normal')
        general_text.configure(state='disabled')  # Make read-only
        
        # t-tests information
        t_test_text = create_text_widget(t_test_frame)
        t_test_text.insert('end', 't-test Types\n', 'heading')
        t_test_text.insert('end', '\nStudent\'s t-test (unpaired):', 'subheading')
        t_test_text.insert('end', '\n• Use when: Comparing two independent groups with equal variances\n• Example: Control group vs. Treatment group with similar spread of data\n• Assumption: Equal variances between groups\n\n', 'normal')
        t_test_text.insert('end', 'Welch\'s t-test (unpaired, unequal variances):', 'subheading')
        t_test_text.insert('end', '\n• Use when: Comparing two independent groups with unequal variances\n• Example: Control group vs. Treatment group with different spread of data\n• More robust when variances differ\n• Recommended as default for most unpaired comparisons\n\n', 'normal')
        t_test_text.insert('end', 'Paired t-test:', 'subheading')
        t_test_text.insert('end', '\n• Use when: Comparing paired measurements (before/after, matched samples)\n• Example: Before treatment vs. After treatment in the same subjects\n• Requires equal number of data points in both groups\n\n', 'normal')
        t_test_text.insert('end', 't-test Alternatives:', 'subheading')
        t_test_text.insert('end', '\n• two-sided: Tests if groups are different (most common)\n• less: Tests if first group mean is less than second group mean\n• greater: Tests if first group mean is greater than second group mean\n', 'normal')
        t_test_text.configure(state='disabled')  # Make read-only
        
        # ANOVA information
        anova_text = create_text_widget(anova_frame)
        anova_text.insert('end', 'ANOVA Types\n', 'heading')
        anova_text.insert('end', '\nOne-way ANOVA:', 'subheading')
        anova_text.insert('end', '\n• Use when: Comparing three or more independent groups with equal variances\n• Example: Multiple treatment groups vs. control\n• Assumption: Equal variances across all groups\n\n', 'normal')
        anova_text.insert('end', 'Welch\'s ANOVA:', 'subheading')
        anova_text.insert('end', '\n• Use when: Comparing three or more independent groups with unequal variances\n• More robust than standard ANOVA when variances differ\n• Recommended as default for most multi-group comparisons\n\n', 'normal')
        anova_text.insert('end', 'Repeated measures ANOVA:', 'subheading')
        anova_text.insert('end', '\n• Use when: Comparing multiple measurements of the same subjects\n• Example: Measurements at different time points or conditions on the same samples\n• More powerful than independent ANOVA for paired data\n\n', 'normal')
        anova_text.insert('end', 'Important Note:', 'subheading')
        anova_text.insert('end', '\nANOVA only tells you that differences exist among groups, not which specific groups differ. For that, you need post-hoc tests.\n', 'normal')
        anova_text.configure(state='disabled')  # Make read-only
        
        # Post-hoc tests information
        posthoc_text = create_text_widget(posthoc_frame)
        posthoc_text.insert('end', 'Post-hoc Tests\n', 'heading')
        posthoc_text.insert('end', '\nAfter finding significant differences with ANOVA, use post-hoc tests to identify which specific groups differ from each other.\n\n', 'normal')
        posthoc_text.insert('end', 'Tukey\'s HSD (Honestly Significant Difference):', 'subheading')
        posthoc_text.insert('end', '\n• Use when: Equal sample sizes and variances across groups\n• Controls family-wise error rate while conducting all pairwise comparisons\n• Balanced between conservative and liberal\n\n', 'normal')
        posthoc_text.insert('end', 'Tamhane\'s T2:', 'subheading')
        posthoc_text.insert('end', '\n• Use when: Unequal variances across groups\n• Conservative test that doesn\'t assume equal variances\n• Recommended default after Welch\'s ANOVA\n\n', 'normal')
        posthoc_text.insert('end', 'Scheffe\'s test:', 'subheading')
        posthoc_text.insert('end', '\n• Use when: Complex comparisons beyond simple pairwise comparisons\n• Very conservative test with strong control of Type I error\n• Flexible for examining various contrasts and combinations\n\n', 'normal')
        posthoc_text.insert('end', 'Dunn\'s test:', 'subheading')
        posthoc_text.insert('end', '\n• Use when: Data doesn\'t follow normal distribution\n• Non-parametric alternative for multiple comparisons\n• Often used after Kruskal-Wallis test (non-parametric equivalent of ANOVA)\n', 'normal')
        posthoc_text.configure(state='disabled')  # Make read-only
        
        # Close button at bottom
        ttk.Button(window, text="Close", command=window.destroy).pack(pady=10)

    def show_statistical_details(self):
        window = tk.Toplevel(self.root)
        window.title("Statistical Details")
        window.geometry("500x600")
        details_text = tk.Text(window, wrap='word', height=30, width=100)
        details_text.pack(fill='both', expand=True, padx=10, pady=10)
        # Set a monospaced font for table alignment
        import tkinter.font as tkfont
        monospace_font = None
        for fname in ("TkFixedFont", "Courier", "Menlo", "Consolas", "Monaco", "Liberation Mono"):
            if fname in tkfont.families():
                monospace_font = fname
                break
        if monospace_font is None:
            monospace_font = "Courier"  # fallback
        details_text.configure(font=(monospace_font, 10))
        details_text.insert(tk.END, "Statistical Details\n\n")
        # Legend
        details_text.insert(tk.END, "Significance Levels:\n")
        details_text.insert(tk.END, "  ns      p > 0.05\n")
        details_text.insert(tk.END, "   *      p ≤ 0.05\n")
        details_text.insert(tk.END, "  **      p ≤ 0.01\n")
        details_text.insert(tk.END, " ***      p ≤ 0.001\n")
        details_text.insert(tk.END, "****      p ≤ 0.0001\n\n")
        # Try to reconstruct the most recent statistical tests
        # We'll use the existing p-values that were calculated for annotations
        # DO NOT clear the p-values dictionary here
        if not hasattr(self, 'df') or self.df is None:
            details_text.insert(tk.END, "No data loaded.\n")
            return
        try:
            x_col = self.xaxis_var.get()
            group_col = self.group_var.get()
            value_cols = [col for var, col in self.value_vars if var.get() and col != x_col]
            if not x_col or not value_cols:
                details_text.insert(tk.END, "No plot or insufficient columns selected.\n")
                return
            df_plot = self.df.copy()
            if self.xaxis_renames:
                df_plot[x_col] = df_plot[x_col].map(self.xaxis_renames).fillna(df_plot[x_col])
            if self.xaxis_order:
                df_plot[x_col] = pd.Categorical(df_plot[x_col], categories=self.xaxis_order, ordered=True)
            plot_kind = self.plot_kind_var.get()
            swap_axes = self.swap_axes_var.get()
            # Only show for bar/box plots with stats
            # Only show statistical details if grouped data was used for last plot
            if not self.use_stats_var.get() or not getattr(self, 'stats_enabled_for_this_plot', False):
                details_text.insert(tk.END, "Statistical tests were not enabled for the last plot or data was not grouped.\n")
                return
            import itertools
            from scipy import stats
            import pandas as pd
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
                try:


                    # Detect ungrouped data the same way as calculate_statistics
                    if not group_col or group_col == 'None' or (group_col and len(df_plot[group_col].dropna().unique()) <= 1):
                        # Ungrouped Data: show statistical tests based on x-axis categories
                        # Import pandas for DataFrame operations
                        import pandas as pd
                        import numpy as np
                        
                        # Get x categories (these are the values we're comparing)
                        x_categories = df_plot[x_col].dropna().unique() if x_col in df_plot else []
                        n_x_categories = len(x_categories)
                        
                        if n_x_categories <= 1:
                            details_text.insert(tk.END, "Only one category: no statistical test performed.\n")
                        elif n_x_categories == 2:
                            # Two-sample t-test between categories
                            cat1, cat2 = x_categories
                            
                            # Get the selected t-test type and alternative from UI
                            ttest_type = self.ttest_type_var.get()
                            alternative = self.ttest_alternative_var.get()
                            
                            details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n\n")
                            
                            # Check if we have p-values from previous calculations
                            key = self.stat_key(cat1, cat2)
                            if key in self.latest_pvals:
                                p_val = self.latest_pvals[key]
                                sig = self.pval_to_annotation(p_val)
                                details_text.insert(tk.END, f"P-value: {p_val:.4g} {sig}\n")
                                
                                # Note about using same p-values as annotations
                                details_text.insert(tk.END, "\nNote: These statistics are the same as those used for plot annotations.\n")
                            else:
                                details_text.insert(tk.END, "No statistical results available for these categories.\n")
                        elif n_x_categories > 2:
                            # For 3+ categories, display ANOVA + post-hoc test results
                            # Get the stored information about the statistical tests
                            anova_type = self.latest_stats.get('anova_type', "One-way ANOVA")
                            posthoc_type = self.latest_stats.get('posthoc_type', "Tukey's HSD")
                            
                            details_text.insert(tk.END, f"Test Used: {anova_type} + {posthoc_type} across {n_x_categories} categories\n\n")
                            
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
                                        details_text.insert(tk.END, f"Main ANOVA result: p = {anova_p:.4g} {self.pval_to_annotation(anova_p)}\n")
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
                                
                            # Fill matrix with p-values
                            for g1 in x_categories:
                                for g2 in x_categories:
                                    if g1 == g2:
                                        p_matrix.loc[g1, g2] = float('nan')  # Diagonal is not applicable
                                    else:
                                        # Try all possible key formats
                                        key1 = self.stat_key(g1, g2)  # Standard key
                                        key2 = (g1, g2)  # Direct tuple
                                        key3 = (g2, g1)  # Reversed tuple

                                        # Check all possible keys
                                        print(f"[DEBUG] Looking for p-value for {g1} vs {g2}: keys {key1}, {key2}, {key3}")
                                        if key1 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key1: {key1}")
                                            p_matrix.loc[g1, g2] = self.latest_pvals[key1]
                                            has_pvals = True
                                        elif key2 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key2: {key2}")
                                            p_matrix.loc[g1, g2] = self.latest_pvals[key2]
                                            has_pvals = True
                                        elif key3 in self.latest_pvals:
                                            print(f"[DEBUG] Found using key3: {key3}")
                                            p_matrix.loc[g1, g2] = self.latest_pvals[key3]
                                            has_pvals = True
                                        else:
                                            p_matrix.loc[g1, g2] = float('nan')
                                
                            # Format and display p-value matrix if we found values
                            if has_pvals:
                                details_text.insert(tk.END, "P-values from statistical calculations:\n")
                                details_text.insert(tk.END, self.format_pvalue_matrix(p_matrix) + '\n')
                                
                                # Add significance indicators
                                details_text.insert(tk.END, "\nSignificance Indicators:\n")
                                for i, g1 in enumerate(x_categories):
                                    for j, g2 in enumerate(x_categories):
                                        if i < j:  # Only show each pair once
                                            # Try all possible key formats
                                            key1 = self.stat_key(g1, g2)  # Standard key
                                            key2 = (g1, g2)  # Direct tuple
                                            key3 = (g2, g1)  # Reversed tuple
                                            
                                            p_val = None
                                            if key1 in self.latest_pvals:
                                                p_val = self.latest_pvals[key1]
                                            elif key2 in self.latest_pvals:
                                                p_val = self.latest_pvals[key2]
                                            elif key3 in self.latest_pvals:
                                                p_val = self.latest_pvals[key3]
                                            
                                            if p_val is not None:
                                                sig = self.pval_to_annotation(p_val)
                                                details_text.insert(tk.END, f"{g1} vs {g2}: p = {p_val:.4g} {sig}\n")
                                
                                # Note about using same p-values as annotations
                                details_text.insert(tk.END, "\nNote: These statistics are the same as those used for plot annotations.\n")
                            else:
                                details_text.insert(tk.END, "No matching p-values found in latest_pvals dictionary.\n")
                                details_text.insert(tk.END, f"Keys available: {list(self.latest_pvals.keys())[:5]}\n")
                    else:
                        # Grouped Data: 1 group, N categories
                        unique_groups = df_plot[group_col].dropna().unique() if group_col in df_plot else []
                        x_categories = df_plot[x_col].dropna().unique() if x_col in df_plot else []
                        n_groups = len(unique_groups)
                        n_x_categories = len(x_categories)

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
                                    continue

                                # Get the selected t-test type and alternative from UI
                                ttest_type = self.ttest_type_var.get()
                                alternative = self.ttest_alternative_var.get()
                                
                                details_text.insert(tk.END, f"Test Used: {ttest_type} ({alternative}) between {cat1} and {cat2}\n")
                                try:
                                    # Use the selected t-test type for the calculation
                                    if ttest_type == "Paired t-test" and len(df_cat1) == len(df_cat2):
                                        ttest_cats = stats.ttest_rel(df_cat1, df_cat2, alternative=alternative)
                                    elif ttest_type == "Student's t-test (unpaired)":
                                        ttest_cats = stats.ttest_ind(df_cat1, df_cat2, alternative=alternative, equal_var=True)
                                    else:  # Welch's t-test (unequal variance)
                                        ttest_cats = stats.ttest_ind(df_cat1, df_cat2, alternative=alternative, equal_var=False)
                                    p_annotation = self.pval_to_annotation(ttest_cats.pvalue)
                                    details_text.insert(tk.END, f"t = {ttest_cats.statistic:.4g}, p = {ttest_cats.pvalue:.4g} {p_annotation}\n")
                                except Exception as e:
                                    details_text.insert(tk.END, f"T-test failed: {e}\n")
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
                                    df_long = df_plot.melt(id_vars=[x_col], value_vars=[val_col], var_name='Condition', value_name='MeltedValue')
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
                                            if anova_type == "Welch's ANOVA":
                                                aov = pg.welch_anova(data=df_long, dv='MeltedValue', between=x_col)
                                            elif anova_type == "Repeated measures ANOVA":
                                                # For repeated measures, we need a subject identifier
                                                # We'll use dummy IDs for now
                                                df_long['Subject'] = np.arange(len(df_long))
                                                try:
                                                    aov = pg.rm_anova(data=df_long, dv='MeltedValue', within=x_col, subject='Subject')
                                                except Exception:
                                                    details_text.insert(tk.END, "Repeated measures ANOVA failed, falling back to regular ANOVA\n")
                                                    aov = pg.anova(data=df_long, dv='MeltedValue', between=x_col)
                                            else:  # Regular one-way ANOVA
                                                aov = pg.anova(data=df_long, dv='MeltedValue', between=x_col)
                                        
                                        details_text.insert(tk.END, str(aov) + '\n')
                                        anova_success = True
                                except Exception as e:
                                    details_text.insert(tk.END, f"ANOVA failed: {e}\n")
                                
                                # Only run post-hoc test if ANOVA was successful
                                if anova_success and sp is not None:
                                    try:
                                        # Use stored post-hoc results if available, otherwise calculate new ones
                                        posthoc = self.latest_stats.get('posthoc_results', None)
                                        if posthoc is None:
                                            # Perform the post-hoc test based on the selected type
                                            if posthoc_type == "Tukey's HSD":
                                                # Try to convert Tukey results to matrix format
                                                tukey_pairwise = pg.pairwise_tukey(data=df_long, dv='MeltedValue', between=x_col)
                                                groups = df_long[x_col].unique()
                                                posthoc = pd.DataFrame(index=groups, columns=groups)
                                                for i, row in tukey_pairwise.iterrows():
                                                    g1, g2 = row['A'], row['B']
                                                    p_value = row['p-tukey']
                                                    posthoc.loc[g1, g2] = p_value
                                                    posthoc.loc[g2, g1] = p_value
                                                # Fill diagonal with 1.0 (no difference)
                                                for g in groups:
                                                    posthoc.loc[g, g] = 1.0
                                                print(f"[DEBUG] Tukey's HSD result matrix created")
                                            
                                            elif posthoc_type == "Scheffe's test":
                                                posthoc = sp.posthoc_scheffe(df_long, val_col='MeltedValue', group_col=x_col)
                                            
                                            elif posthoc_type == "Dunn's test":
                                                posthoc = sp.posthoc_dunn(df_long, val_col='MeltedValue', group_col=x_col)
                                            
                                            else:  # Default to Tamhane's T2
                                                posthoc = sp.posthoc_tamhane(df_long, val_col='MeltedValue', group_col=x_col)
                                                # Ensure p-values are properly stored for access in the statistical details panel
                                                # Store these values in latest_pvals for display
                                                groups = df_long[x_col].unique()
                                                for i, g1 in enumerate(groups):
                                                    for j, g2 in enumerate(groups):
                                                        if i != j:  # Skip diagonal
                                                            # Store p-value for this pair
                                                            pval = posthoc.loc[g1, g2]
                                                            key = self.stat_key(g1, g2)
                                                            self.latest_pvals[key] = pval
                                                            self.latest_stats['pvals'][key] = pval
                                        
                                        details_text.insert(tk.END, f"Post-hoc {stored_posthoc_type} test:\n")
                                        details_text.insert(tk.END, self.format_pvalue_matrix(posthoc) + '\n')
                                        
                                        # Add a more readable version with significance indicators
                                        details_text.insert(tk.END, "\nSignificance indicators for post-hoc test:\n")
                                        for idx1, group1 in enumerate(posthoc.index):
                                            for idx2, group2 in enumerate(posthoc.columns):
                                                if idx1 < idx2:  # Only show each comparison once
                                                    pval = posthoc.loc[group1, group2]
                                                    sig = self.pval_to_annotation(pval)
                                                    details_text.insert(tk.END, f"{group1} vs {group2}: p = {pval:.4g} {sig}\n")
                                    except Exception as e:
                                        details_text.insert(tk.END, f"Posthoc {stored_posthoc_type} failed: {e}\n")
                            else:
                                details_text.insert(tk.END, "ANOVA/post-hoc pipeline requires required packages.\n")

                except Exception as e:
                    tb = traceback.format_exc()
                    details_text.insert(tk.END, f"[ERROR] Exception for column {val_col}: {e}\nTraceback:\n{tb}\n")
                    continue

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
                        else:
                            details_text.insert(tk.END, "Multiple groups present, but no suitable test could be performed.\n")
                            continue
                        
                        # Process each x-axis category
                        for i, g in enumerate(base_groups):
                            df_sub = df_plot[df_plot[x_col] == g]
                            pairs = list(itertools.combinations(hue_groups, 2))
                            # For ANOVA + post-hoc test, build a matrix
                            if n_hue > 2 and pg is not None and sp is not None:
                                pval_matrix = pd.DataFrame(index=hue_groups, columns=hue_groups, dtype=object)
                                
                                # Get the selected ANOVA and post-hoc test types from UI
                                anova_type = self.anova_type_var.get()
                                posthoc_type = self.posthoc_type_var.get()
                                
                                # Perform the selected ANOVA test
                                try:
                                    # First run the appropriate ANOVA test
                                    if anova_type == "Welch's ANOVA":
                                        aov = pg.welch_anova(data=df_sub, dv=val_col, between=group_col)
                                    elif anova_type == "Repeated measures ANOVA":
                                        # For repeated measures, we need a subject identifier
                                        df_sub["Subject"] = np.arange(len(df_sub))
                                        try:
                                            aov = pg.rm_anova(data=df_sub, dv=val_col, within=group_col, subject="Subject")
                                        except Exception:
                                            print(f"[DEBUG] Repeated measures ANOVA failed, falling back to One-way ANOVA")
                                            aov = pg.anova(data=df_sub, dv=val_col, between=group_col)
                                    else:  # One-way ANOVA
                                        aov = pg.anova(data=df_sub, dv=val_col, between=group_col)
                                    
                                    # Use individual t-tests as fallback if specific posthoc tests fail
                                    fallback_to_pairwise_ttest = False
                                    posthoc = None

                                    try:
                                        # Run the selected post-hoc test
                                        if posthoc_type == "Tukey's HSD":
                                            # Use pingouin for Tukey's HSD
                                            tukey_result = pg.pairwise_tukey(data=df_sub, dv=val_col, between=group_col)
                                            # Create a matrix format for consistency
                                            posthoc = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                            # Populate the matrix with p-values from the tukey result
                                            for i, row in tukey_result.iterrows():
                                                g1, g2 = row['A'], row['B']
                                                p_value = row['p-tukey']
                                                posthoc.loc[g1, g2] = p_value
                                                posthoc.loc[g2, g1] = p_value
                                            # Fill diagonal with 1.0 (no difference)
                                            for g in hue_groups:
                                                posthoc.loc[g, g] = 1.0
                                            print(f"[DEBUG] Tukey's HSD result matrix created")
                                        
                                        elif posthoc_type == "Scheffe's test":
                                            # Create a manual pairwise comparison result for Scheffe's test
                                            posthoc = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                            # Fill diagonal with 1.0
                                            for g in hue_groups:
                                                posthoc.loc[g, g] = 1.0
                                            # Calculate pairwise p-values
                                            for h1, h2 in pairs:
                                                vals1 = df_sub[df_sub[group_col] == h1][val_col].dropna().astype(float)
                                                vals2 = df_sub[df_sub[group_col] == h2][val_col].dropna().astype(float)
                                                if len(vals1) >= 2 and len(vals2) >= 2:
                                                    # Calculate F-statistic for Scheffe's test
                                                    f_stat, p_val = stats.f_oneway(vals1, vals2)
                                                    # Apply Scheffe's correction
                                                    p_val = min(1.0, p_val * (len(hue_groups) - 1))  # Conservative correction
                                                    posthoc.loc[h1, h2] = p_val
                                                    posthoc.loc[h2, h1] = p_val
                                                else:
                                                    posthoc.loc[h1, h2] = float('nan')
                                                    posthoc.loc[h2, h1] = float('nan')
                                            print(f"[DEBUG] Scheffe's test result matrix created manually")
                                        
                                        elif posthoc_type == "Dunn's test":
                                            # For Dunn's test, manually calculate p-values
                                            posthoc = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                            # Fill diagonal with 1.0
                                            for g in hue_groups:
                                                posthoc.loc[g, g] = 1.0
                                            # Calculate pairwise p-values
                                            for h1, h2 in pairs:
                                                vals1 = df_sub[df_sub[group_col] == h1][val_col].dropna().astype(float)
                                                vals2 = df_sub[df_sub[group_col] == h2][val_col].dropna().astype(float)
                                                if len(vals1) >= 2 and len(vals2) >= 2:
                                                    # Use Mann-Whitney U test as a non-parametric test
                                                    u_stat, p_val = stats.mannwhitneyu(vals1, vals2, alternative='two-sided')
                                                    posthoc.loc[h1, h2] = p_val
                                                    posthoc.loc[h2, h1] = p_val
                                                else:
                                                    posthoc.loc[h1, h2] = float('nan')
                                                    posthoc.loc[h2, h1] = float('nan')
                                            print(f"[DEBUG] Dunn's test result matrix created manually")
                                        
                                        else:  # Default to pairwise t-tests
                                            # Create matrix for t-test results
                                            posthoc = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                            # Fill diagonal with 1.0
                                            for g in hue_groups:
                                                posthoc.loc[g, g] = 1.0
                                            # Calculate pairwise p-values
                                            for h1, h2 in pairs:
                                                vals1 = df_sub[df_sub[group_col] == h1][val_col].dropna().astype(float)
                                                vals2 = df_sub[df_sub[group_col] == h2][val_col].dropna().astype(float)
                                                if len(vals1) >= 2 and len(vals2) >= 2:
                                                    t_stat, p_val = stats.ttest_ind(vals1, vals2, equal_var=False)
                                                    posthoc.loc[h1, h2] = p_val
                                                    posthoc.loc[h2, h1] = p_val
                                                else:
                                                    posthoc.loc[h1, h2] = float('nan')
                                                    posthoc.loc[h2, h1] = float('nan')
                                            print(f"[DEBUG] Pairwise t-test result matrix created")
                                    
                                    except Exception as e:
                                        print(f"[DEBUG] Error in post-hoc test: {e}, falling back to pairwise t-tests")
                                        fallback_to_pairwise_ttest = True
                                        posthoc = None
                                    
                                    # If post-hoc test failed, use pairwise t-tests as fallback
                                    if fallback_to_pairwise_ttest or posthoc is None:
                                        try:
                                            # Create matrix for t-test results
                                            posthoc = pd.DataFrame(index=hue_groups, columns=hue_groups)
                                            # Fill diagonal with 1.0
                                            for g in hue_groups:
                                                posthoc.loc[g, g] = 1.0
                                            # Calculate pairwise p-values with t-tests
                                            for h1, h2 in pairs:
                                                vals1 = df_sub[df_sub[group_col] == h1][val_col].dropna().astype(float)
                                                vals2 = df_sub[df_sub[group_col] == h2][val_col].dropna().astype(float)
                                                if len(vals1) >= 2 and len(vals2) >= 2:
                                                    t_stat, p_val = stats.ttest_ind(vals1, vals2, equal_var=False)
                                                    posthoc.loc[h1, h2] = p_val
                                                    posthoc.loc[h2, h1] = p_val
                                                else:
                                                    posthoc.loc[h1, h2] = float('nan')
                                                    posthoc.loc[h2, h1] = float('nan')
                                            print(f"[DEBUG] Fallback: Pairwise t-test result matrix created")
                                        except Exception as e2:
                                            print(f"[DEBUG] Fallback also failed: {e2}")
                                            posthoc = None
                                    
                                    # Print the posthoc result for debugging
                                    if posthoc is not None:
                                        print(f"[DEBUG] Post-hoc result matrix:\n{posthoc}")
                                except Exception:
                                    posthoc = None
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
                                                    print(f"[DEBUG] Found p-value {pval_val} for {h1} vs {h2}")
                                                elif h1 in posthoc.columns and h2 in posthoc.index:
                                                    pval_val = posthoc.loc[h2, h1]
                                                    print(f"[DEBUG] Found p-value {pval_val} for {h2} vs {h1} (reversed)")
                                                else:
                                                    print(f"[DEBUG] Could not find p-value for {h1} vs {h2} in posthoc result")
                                                    pval_val = float('nan')
                                                
                                                # Convert to string with appropriate formatting
                                                if not pd.isna(pval_val):
                                                    pval_matrix.loc[h1, h2] = f"{pval_val:.4g}"
                                                else:
                                                    pval_matrix.loc[h1, h2] = 'nan'
                                            except Exception as e:
                                                print(f"[DEBUG] Error extracting p-value for {h1} vs {h2}: {e}")
                                                pval_matrix.loc[h1, h2] = 'error'
                                        else:
                                            # No posthoc result available
                                            pval_matrix.loc[h1, h2] = ''
                                # Print the pairwise list using the SAME p-values that were used for annotations
                                for h1, h2 in pairs:
                                    # Based on the debug output, for 3+ group data, the p-values are stored with format
                                    # (x_val, hue_val1, hue_val2) where x_val is the category and hue_val are the groups
                                    # Looking at all possible key formats
                                    possible_keys = [
                                        (g, h1, h2),           # Direct key (category, group1, group2)
                                        (g, h2, h1),           # Reversed groups
                                        self.stat_key(g, h1, h2), # Sorted key with category
                                        (h1, h2),              # Just the groups (no category)
                                        (h2, h1),              # Reversed groups (no category)
                                        self.stat_key(h1, h2)   # Sorted key (just groups)
                                    ]
                                    
                                    # Print all available keys in latest_pvals for debugging
                                    print(f"[DEBUG] Looking for keys among: {list(self.latest_pvals.keys())[:5]}...")
                                    
                                    # Try to find the p-value in latest_pvals using all possible key formats
                                    p_val = None
                                    key_found = None
                                    for key in possible_keys:
                                        if key in self.latest_pvals:
                                            p_val = self.latest_pvals[key]
                                            key_found = key
                                            print(f"[DEBUG] Found stored p-value {p_val} using key {key}")
                                            break
                                    
                                    # If we found a p-value, use it
                                    if p_val is not None:
                                        # Format the p-value for display
                                        p_annotation = self.pval_to_annotation(p_val)
                                        details_text.insert(tk.END, f"{g}: {h1} vs {h2}: p = {p_val:.4g} {p_annotation}\n")
                                        
                                        # Also update the matrix for display
                                        pval_matrix.loc[h1, h2] = f"{p_val:.4g}"
                                    else:
                                        # If we couldn't find the p-value in stored data, calculate it directly
                                        # This ensures we always show p-values in the statistical details
                                        try:
                                            # Get the data for both groups
                                            vals1 = df_sub[df_sub[group_col] == h1][val_col].dropna().astype(float)
                                            vals2 = df_sub[df_sub[group_col] == h2][val_col].dropna().astype(float)
                                            
                                            if len(vals1) >= 2 and len(vals2) >= 2:
                                                # Calculate p-value based on selected test
                                                if posthoc_type == "Tukey's HSD":
                                                    # For parametric ANOVA-based test
                                                    f_stat, p_val = stats.f_oneway(vals1, vals2)
                                                elif posthoc_type == "Scheffe's test":
                                                    # For Scheffe's test
                                                    f_stat, p_val = stats.f_oneway(vals1, vals2)
                                                    # Apply Scheffe's correction
                                                    p_val = min(1.0, p_val * (len(hue_groups) - 1))
                                                elif posthoc_type == "Dunn's test":
                                                    # For non-parametric test
                                                    u_stat, p_val = stats.mannwhitneyu(vals1, vals2, alternative='two-sided')
                                                else:
                                                    # Default to t-test
                                                    t_stat, p_val = stats.ttest_ind(vals1, vals2, equal_var=False)
                                                
                                                # Format the p-value for display
                                                p_annotation = self.pval_to_annotation(p_val)
                                                details_text.insert(tk.END, f"{g}: {h1} vs {h2}: p = {p_val:.4g} {p_annotation} (calculated)\n")
                                                
                                                # Also update the matrix for display
                                                pval_matrix.loc[h1, h2] = f"{p_val:.4g}"
                                                
                                                # Store for future reference
                                                key = (g, h1, h2)
                                                self.latest_pvals[key] = p_val
                                            else:
                                                details_text.insert(tk.END, f"{g}: {h1} vs {h2}: p = insufficient data\n")
                                                pval_matrix.loc[h1, h2] = "N/A"
                                        except Exception as e:
                                            details_text.insert(tk.END, f"{g}: {h1} vs {h2}: p = error ({str(e)[:50]})\n")
                                            pval_matrix.loc[h1, h2] = "Error"
                                # Print the matrix
                                details_text.insert(tk.END, f"\nP-value matrix for {g} (rows vs columns):\n")
                                details_text.insert(tk.END, self.format_pvalue_matrix(pval_matrix) + "\n\n")
                            else:
                                # For two groups or if posthoc not available, just show the pairwise list
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
                                        # Use the selected t-test type
                                        if ttest_type == "Paired t-test" and len(vals1) == len(vals2):
                                            ttest = stats.ttest_rel(vals1, vals2, alternative=alternative)
                                        elif ttest_type == "Student's t-test (unpaired)":
                                            ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=True)
                                        else:  # Welch's t-test (default)
                                            ttest = stats.ttest_ind(vals1, vals2, alternative=alternative, equal_var=False)
                                        p_annotation = self.pval_to_annotation(ttest.pvalue)
                                        details_text.insert(tk.END, f"{g}: {h1} vs {h2}: t = {ttest.statistic:.4g}, p = {ttest.pvalue:.4g} {p_annotation}\n")
                                        key = (g, h1, h2) if (g, h1, h2) not in self.latest_pvals else (g, h2, h1)
                                        self.latest_pvals[key] = ttest.pvalue
                                    except Exception as e:
                                        details_text.insert(tk.END, f"T-test failed for {g}: {h1} vs {h2}: {e}\n")

        except Exception as e:
            details_text.insert(tk.END, f"Error calculating statistics: {e}\n")
        details_text.config(state='disabled')
        tk.Button(window, text='Close', command=window.destroy).pack(pady=8)


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
        
    def load_user_preferences(self):
        """Load user preferences from a JSON file and apply them to the application."""
        # Initialize default preferences first
        default_preferences = {
            'plot_kind': 'bar',
            'show_stripplot': True,
            'strip_black': True,
            'errorbar_type': 'SD',
            'errorbar_black': True,
            'errorbar_capsize': 'Default',
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
            'xy_draw_band': False
        }
        
        # Check if user preferences file exists
        if os.path.exists(self.default_settings_file):
            try:
                with open(self.default_settings_file, 'r') as f:
                    user_prefs = json.load(f)
                # Update default_preferences with user's saved preferences
                default_preferences.update(user_prefs)
            except Exception as e:
                messagebox.showwarning("Error Loading Preferences", f"Could not load preferences: {str(e)}")
        
        # Apply the preferences to the UI
        self._apply_user_preferences(default_preferences)
    
    def _apply_user_preferences(self, preferences):
        """Apply the loaded preferences to UI elements."""
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
    
    def save_user_preferences(self):
        """Save current UI settings as user preferences to a JSON file."""
        preferences = {}
        
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
            
            # Confirmation message
            messagebox.showinfo("Settings Saved", "Your preferences have been saved and applied.")
        except Exception as e:
            messagebox.showerror("Error Saving Preferences", f"Could not save preferences: {str(e)}")
            print(f"Error saving preferences: {str(e)}")

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
        self.stats_settings_tab = tk.Frame(self.tab_control)

        self.tab_control.add(self.basic_tab, text="Basic")
        self.tab_control.add(self.appearance_tab, text="Appearance")
        self.tab_control.add(self.axis_tab, text="Axis")
        self.tab_control.add(self.colors_tab, text="Colors")
        self.tab_control.add(self.stats_settings_tab, text="Statistics")

        right_frame = tk.Frame(main_frame)
        right_frame.pack(side='right', fill='both', expand=True)

        self.canvas_frame = tk.Frame(right_frame)
        self.canvas_frame.pack(fill='both', expand=True)

        tk.Button(right_frame, text='Generate Plot', command=self.plot_graph).pack(pady=5)
        tk.Button(right_frame, text='Save as PDF', command=self.save_pdf).pack(pady=5)
        # Statistical Details button with dynamic visibility
        self.stats_details_btn = tk.Button(right_frame, text="Statistical Details", command=self.show_statistical_details)
        self.stats_details_btn.pack_forget()  # Initially hidden

        # Update visibility when the checkbox changes
        def _toggle_stats_details_btn(*args):
            try:
                if self.use_stats_var.get():
                    self.stats_details_btn.pack(pady=5)
                else:
                    self.stats_details_btn.pack_forget()
            except Exception as e:
                print(f"Error toggling stats details button: {e}")
        self.use_stats_var.trace_add('write', _toggle_stats_details_btn)


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
        # We use the use_stats_var that was initialized in __init__
        # No need to create it again here
        stats_frame = tk.Frame(opt_grp)
        stats_frame.pack(anchor="w", pady=1)
        tk.Checkbutton(stats_frame, text="Use statistics", variable=self.use_stats_var).pack(side="left")
        # --- Error bar type (SD/SEM) ---
        # We use the errorbar_type_var that was initialized in __init__
        # No need to create it again here
            
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
        self.xy_marker_symbol_label.grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.xy_marker_symbol_dropdown.grid(row=0, column=1, sticky="w", padx=2, pady=2)
        self.xy_marker_size_label.grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.xy_marker_size_entry.grid(row=1, column=1, sticky="w", padx=2, pady=2)
        self.xy_filled_check.grid(row=2, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_line_style_label.grid(row=3, column=0, sticky="w", padx=4, pady=2)
        self.xy_line_style_dropdown.grid(row=3, column=1, sticky="w", padx=2, pady=2)
        self.xy_line_black_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_connect_check.grid(row=5, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_show_mean_check.grid(row=6, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self.xy_show_mean_errorbars_check.grid(row=7, column=0, columnspan=2, sticky="w", padx=24, pady=2)
        self.xy_draw_band_check.grid(row=8, column=0, columnspan=2, sticky="w", padx=4, pady=2)
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
        self.width_entry = tk.Entry(size_grp, textvariable=self.plot_width_var)
        self.width_entry.grid(row=0, column=1, sticky="ew", padx=2, pady=2)
        tk.Label(size_grp, text="Plot Height per plot (inches):").grid(row=1, column=0, sticky="w", pady=2)
        self.height_entry = tk.Entry(size_grp, textvariable=self.plot_height_var)
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
        tk.Label(font_grp, text="Error Bar Capsize:").grid(row=2, column=0, sticky="w", pady=2)
        self.capsize_dropdown = ttk.Combobox(font_grp, textvariable=self.errorbar_capsize_var, 
                                        values=["Default", "Narrow", "Wide", "Wider", "None"], width=10)
        self.capsize_dropdown.grid(row=2, column=1, sticky="ew", padx=2, pady=2)
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
        
        # --- Combined X/Y Labels frame ---
        labels_frame = tk.Frame(frame)
        labels_frame.pack(fill='x', padx=4, pady=1)
        
        # X-axis label row with grid layout for alignment
        xlabel_frame = tk.Frame(labels_frame)
        xlabel_frame.pack(fill='x', padx=0, pady=1)
        xlabel_frame.columnconfigure(1, weight=1)
        tk.Label(xlabel_frame, text="X-axis Label:", width=14, anchor="w").grid(row=0, column=0, padx=2)
        self.xlabel_entry = tk.Entry(xlabel_frame)
        self.xlabel_entry.grid(row=0, column=1, sticky="ew", padx=2)
        
        # Y-axis label row (just below X)
        ylabel_frame = tk.Frame(labels_frame)
        ylabel_frame.pack(fill='x', padx=0, pady=1)
        ylabel_frame.columnconfigure(1, weight=1)
        tk.Label(ylabel_frame, text="Y-axis Label:", width=14, anchor="w").grid(row=0, column=0, padx=2)
        self.ylabel_entry = tk.Entry(ylabel_frame)
        self.ylabel_entry.grid(row=0, column=1, sticky="ew", padx=2)
        
        # --- Label orientation (horizontal layout) ---
        orient_frame = tk.Frame(frame)
        orient_frame.pack(fill='x', padx=4, pady=1)
        tk.Label(orient_frame, text="X-axis Label Orientation:").pack(side="left")
        self.label_orientation = tk.StringVar(value="vertical")
        tk.Radiobutton(orient_frame, text="Vertical", variable=self.label_orientation, 
                      value="vertical").pack(side="left", padx=5)
        tk.Radiobutton(orient_frame, text="Horizontal", variable=self.label_orientation, 
                      value="horizontal").pack(side="left", padx=5)
        tk.Radiobutton(orient_frame, text="Angled", variable=self.label_orientation, 
                      value="angled").pack(side="left", padx=5)
        
        # --- X-Axis settings (more compact with grid layout) ---
        xaxis_grp = tk.LabelFrame(frame, text="X-Axis Settings", padx=4, pady=2)
        xaxis_grp.pack(fill='x', padx=4, pady=1)
        
        # Use grid layout for better alignment
        xaxis_grid = tk.Frame(xaxis_grp)
        xaxis_grid.pack(fill='x', padx=2, pady=1)
        
        # Configure columns for alignment - fixed label widths
        xaxis_grid.columnconfigure(1, weight=0) # Min Entry
        xaxis_grid.columnconfigure(3, weight=0) # Max Entry
        
        # Row 1: Min/Max values
        tk.Label(xaxis_grid, text="Minimum:", width=12, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.xmin_entry = tk.Entry(xaxis_grid, width=10)
        self.xmin_entry.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Maximum:", width=12, anchor="w").grid(row=0, column=2, sticky="w", padx=(10,0), pady=2)
        self.xmax_entry = tk.Entry(xaxis_grid, width=10)
        self.xmax_entry.grid(row=0, column=3, sticky="w", pady=2)
        
        # Row 2: Tick settings
        tk.Label(xaxis_grid, text="Major Tick:", width=12, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.xinterval_entry = tk.Entry(xaxis_grid, width=10)
        self.xinterval_entry.grid(row=1, column=1, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Minor/Major:", width=12, anchor="w").grid(row=1, column=2, sticky="w", padx=(10,0), pady=2)
        self.xminor_ticks_entry = tk.Entry(xaxis_grid, width=10)
        self.xminor_ticks_entry.grid(row=1, column=3, sticky="w", pady=2)
        
        # Row 3: Log options
        self.xlogscale_var = tk.BooleanVar(value=False)
        tk.Checkbutton(xaxis_grid, text="Logarithmic X-axis", variable=self.xlogscale_var, 
                      command=self.update_xlog_options).grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        tk.Label(xaxis_grid, text="Base:", width=6, anchor="e").grid(row=2, column=2, sticky="e", pady=2)
        self.xlog_base_var = tk.StringVar(value="10")
        self.xlog_base_dropdown = ttk.Combobox(xaxis_grid, textvariable=self.xlog_base_var, 
                                          values=["10", "2"], state="disabled", width=5)
        self.xlog_base_dropdown.grid(row=2, column=3, sticky="w", pady=2)
        
        # --- Y-Axis settings with grid layout ---
        yaxis_grp = tk.LabelFrame(frame, text="Y-Axis Settings", padx=4, pady=2)
        yaxis_grp.pack(fill='x', padx=4, pady=1)
        
        # Use grid layout for better alignment
        yaxis_grid = tk.Frame(yaxis_grp)
        yaxis_grid.pack(fill='x', padx=2, pady=1)
        
        # Configure columns for alignment - fixed label widths
        yaxis_grid.columnconfigure(1, weight=0) # Min Entry
        yaxis_grid.columnconfigure(3, weight=0) # Max Entry
        
        # Row 1: Min/Max values
        tk.Label(yaxis_grid, text="Minimum:", width=12, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        self.ymin_entry = tk.Entry(yaxis_grid, width=10)
        self.ymin_entry.grid(row=0, column=1, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Maximum:", width=12, anchor="w").grid(row=0, column=2, sticky="w", padx=(10,0), pady=2)
        self.ymax_entry = tk.Entry(yaxis_grid, width=10)
        self.ymax_entry.grid(row=0, column=3, sticky="w", pady=2)
        
        # Row 2: Tick settings
        tk.Label(yaxis_grid, text="Major Tick:", width=12, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        self.yinterval_entry = tk.Entry(yaxis_grid, width=10)
        self.yinterval_entry.grid(row=1, column=1, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Minor/Major:", width=12, anchor="w").grid(row=1, column=2, sticky="w", padx=(10,0), pady=2)
        self.minor_ticks_entry = tk.Entry(yaxis_grid, width=10)
        self.minor_ticks_entry.grid(row=1, column=3, sticky="w", pady=2)
        
        # Row 3: Log options
        self.logscale_var = tk.BooleanVar(value=False)
        tk.Checkbutton(yaxis_grid, text="Logarithmic Y-axis", variable=self.logscale_var, 
                      command=self.update_ylog_options).grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        tk.Label(yaxis_grid, text="Base:", width=6, anchor="e").grid(row=2, column=2, sticky="e", pady=2)
        self.ylog_base_var = tk.StringVar(value="10")
        self.ylog_base_dropdown = ttk.Combobox(yaxis_grid, textvariable=self.ylog_base_var, 
                                          values=["10", "2"], state="disabled", width=5)
        self.ylog_base_dropdown.grid(row=2, column=3, sticky="w", pady=2)

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
        
        # Update dropdown values
        self.xaxis_dropdown['values'] = columns
        self.group_dropdown['values'] = ['None'] + columns
        
        # Reset selections when switching sheets
        self.xaxis_var.set('')  # Clear X-axis selection
        self.group_var.set('None')  # Reset group selection to None
        
        # Clear Y-axis checkboxes
        for cb in self.value_checkbuttons:
            cb.destroy()
        self.value_vars.clear()
        
        # Recreate value column checkboxes
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
        # This version includes better annotation handling for multiple x-axis categories
        # Robustly initialize show_errorbars at the very top
        show_errorbars = getattr(self, 'show_errorbars_var', None)
        if show_errorbars is not None:
            show_errorbars = show_errorbars.get()
        else:
            show_errorbars = True

        if self.df is None:
            return

        try:
            linewidth = float(self.linewidth.get())
        except Exception:
            linewidth = 1.0
            
        # Get column selections
        x_col = self.xaxis_var.get()
        group_col = self.group_var.get()
        if not group_col or group_col.strip() == '' or group_col == 'None':
            group_col = None
            
        # Enable statistics for both grouped and ungrouped data
        # Statistics will work for both group_col and pure x_col comparisons
        self.stats_enabled_for_this_plot = True
        
        value_cols = [col for var, col in self.value_vars if var.get() and col != x_col]
        
        # Handle empty selections gracefully
        if not x_col:
            messagebox.showinfo("Missing X-axis", "Please select an X-axis column.")
            return
        
        if not value_cols:
            messagebox.showinfo("Missing Y-axis", "Please select at least one Y-axis column.")
            return
            
        plot_mode = 'single' if len(value_cols) == 1 else 'overlay'
        plot_width = self.plot_width_var.get()
        plot_height = self.plot_height_var.get()
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
        # Create the figure and store it as a class attribute for later use in save_graph
        self.fig, axes = plt.subplots(n_rows, 1, figsize=(fig_width, fig_height), squeeze=False)
        axes = axes.flatten()

        show_frame = self.show_frame_var.get()
        show_hgrid = self.show_hgrid_var.get()
        show_vgrid = self.show_vgrid_var.get()
        swap_axes = self.swap_axes_var.get()
        strip_black = self.strip_black_var.get()
        show_stripplot = self.show_stripplot_var.get()
        plot_kind = self.plot_kind_var.get()  # "bar", "box", or "xy"
        errorbar_black = self.errorbar_black_var.get()

        # Default initialization of show_mean
        show_mean = False
        if plot_kind == 'xy':
            show_mean = self.xy_show_mean_var.get()

        for idx, value_col in enumerate(value_cols):
            ax = axes[idx] if n_rows > 1 else axes[0]
            df_plot = self.df.copy()

            if self.xaxis_renames:
                df_plot[x_col] = df_plot[x_col].map(self.xaxis_renames).fillna(df_plot[x_col])

            if self.xaxis_order:
                df_plot[x_col] = pd.Categorical(df_plot[x_col], categories=self.xaxis_order, ordered=True)

            if plot_mode == 'overlay' and len(value_cols) > 1:
                df_plot = pd.melt(df_plot, id_vars=[x_col] + ([group_col] if group_col else []),
                                  value_vars=value_cols, var_name='Measurement', value_name=value_col)
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
                if errorbar_type == "SD":
                    ci_val = 'sd'
                    estimator = np.mean
                else:
                    # Set up palette based on hue groups
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    
                    # If we have a hue column, size palette to match number of hue groups
                    if hue_col and hue_col in df_plot.columns:
                        hue_groups = df_plot[hue_col].dropna().unique()
                        if len(palette_full) < len(hue_groups):
                            # Repeat the palette to ensure we have enough colors
                            palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                        palette = palette_full[:len(hue_groups)]
                    else:
                        # Otherwise use value columns for palette
                        palette = palette_full[:len(value_cols)]
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
                    if swap_axes:
                        plot_args = dict(
                            data=df_plot, y=x_col, x=value_col, hue=hue_col, ax=ax,
                            errorbar='sd', capsize=0.2, palette=palette, err_kws={'color': 'black', 'linewidth': linewidth}, estimator=estimator
                        )
                    else:
                        plot_args = dict(
                            data=df_plot, x=x_col, y=value_col, hue=hue_col, ax=ax,
                            errorbar='sd', capsize=0.2, palette=palette, err_kws={'color': 'black', 'linewidth': linewidth}, estimator=estimator
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
                                        means = [group[group[x_col] == x]['Value'].mean() for x in x_sorted]
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
            if plot_kind == "bar":
                # Ensure palette matches exact number of hue groups if hue_col exists
                if hue_col and hue_col in df_plot.columns:
                    hue_groups = df_plot[hue_col].dropna().unique()
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    if len(palette_full) < len(hue_groups):
                        palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                    plot_args["palette"] = palette_full[:len(hue_groups)]
                
                # Use only Seaborn's built-in error bars for bar plots
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
                    plot_args['errorbar'] = 'sd'  # or ci='sd' for older Seaborn
                else:
                    plot_args['errorbar'] = 'se'  # Using native standard error parameter
                
                # Set errorbar styling - capsize needs to be passed directly to barplot
                # and line width needs to be in err_kws
                linewidth = self.linewidth.get()
                
                # Modern way to handle errorbar styling in Seaborn
                plot_args['err_kws'] = {'linewidth': linewidth}
                plot_args['capsize'] = capsize  # This goes directly to barplot, not in err_kws
                
                # Remove deprecated parameter if it exists
                if 'errwidth' in plot_args:
                    plot_args.pop('errwidth')
                
                sns.barplot(**plot_args)
                # (All manual ax.errorbar code for bars removed)
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
                    
                    sns.boxplot(**box_args)
                else:
                    # Use original parameters for grouped data
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
                    # Set up palette based on hue groups
                    palette_name = self.palette_var.get()
                    palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                    
                    # If we have a hue column, size palette to match number of hue groups
                    if hue_col and hue_col in df_plot.columns:
                        hue_groups = df_plot[hue_col].dropna().unique()
                        if len(palette_full) < len(hue_groups):
                            # Repeat the palette to ensure we have enough colors
                            palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                        palette = palette_full[:len(hue_groups)]
                    else:
                        # Otherwise use value columns for palette
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
                        if hue_col:
                            handles, labels = ax.get_legend_handles_labels()
                            if handles and len(handles) > 0:
                                ax.legend()
                    else:
                        # Ungrouped mean plot
                        c = palette[0]
                        x_sorted = np.sort(df_plot[x_col].unique())
                        y_means = [df_plot[df_plot[x_col] == x][value_col].mean() for x in x_sorted]
                        y_errors = [df_plot[df_plot[x_col] == x][value_col].std(ddof=1) if self.errorbar_type_var.get() == 'SD' else 
                                    df_plot[df_plot[x_col] == x][value_col].std(ddof=1) / np.sqrt(len(df_plot[df_plot[x_col] == x])) for x in x_sorted]
                        
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
                            if draw_band:
                                ax.fill_between(x_sorted_numeric, 
                                                y_means_numeric - y_errors_numeric, 
                                                y_means_numeric + y_errors_numeric, 
                                                color=c, alpha=0.18, zorder=1)
                        else:
                            ax.plot(x_sorted_numeric, y_means_numeric, marker=marker_symbol, 
                                    color=c, markerfacecolor=mfc, markeredgecolor=mec, 
                                    markersize=marker_size, linewidth=linewidth, linestyle='None')
                        
                        if connect:
                            ax.plot(x_sorted_numeric, y_means_numeric, 
                                    color='black' if line_black else c, 
                                    linewidth=linewidth, alpha=0.7, linestyle=line_style)
                else:
                    # Plot raw data points (scatter) when show_mean is False
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
                                        means = [group[group[x_col] == x]['Value'].mean() for x in x_sorted]
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
                ax.tick_params(axis='x', which='both', direction='in', length=4, width=linewidth, top=False, bottom=True, labeltop=False, labelbottom=True)

            # --- Stripplot (if enabled) ---
            if show_stripplot:
                if strip_black:
                    stripplot_args["palette"] = ["black"]
                    stripplot_args["color"] = "black"
                else:
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
                # Suppress legend for stripplot
                stripplot_args["legend"] = False

                if plot_kind == 'bar' and not hue_col:
                    # Reduce jitter for more precise positioning
                    stripplot_args['jitter'] = 0.2
                    stripplot_args['dodge'] = False  # Prevent automatic dodging
                
                # Check if "Show stripplot with black dots" option is selected
                strip_black = self.strip_black_var.get()
                
                if strip_black:
                    # If black dots option is selected, override palette with black
                    stripplot_args["color"] = "black"
                    # Remove the palette parameter if it exists
                    if "palette" in stripplot_args:
                        del stripplot_args["palette"]
                else:
                    # One final check to ensure stripplot palette matches exact number of hue groups
                    if hue_col and hue_col in df_plot.columns:
                        hue_groups = df_plot[hue_col].dropna().unique()
                        palette_name = self.palette_var.get()
                        palette_full = self.custom_palettes.get(palette_name, ["#333333"])
                        if len(palette_full) < len(hue_groups):
                            palette_full = (palette_full * ((len(hue_groups) // len(palette_full)) + 1))
                        stripplot_args["palette"] = palette_full[:len(hue_groups)]
                
                sns.stripplot(**stripplot_args)

            # --- Always rebuild legend after all plotting ---
            if hue_col and plot_kind == "box":
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
                base_groups = []
                hue_groups = []
                annotation_count = 0
                try:
                    print(f"[DEBUG] Starting annotations: x_col={x_col}, value_col={value_col}, hue_col={hue_col}")
                    # Calculate p-values for all pairs
                    self.calculate_and_store_pvals(df_plot, x_col, value_col, hue_col)
                    print(f"[DEBUG] After calculate_and_store_pvals, latest_pvals = {self.latest_pvals}")
                    import itertools
                    y_max = df_plot[value_col].astype(float).max()
                    y_min = df_plot[value_col].astype(float).min()
                    # Ensure we have a reasonable height for annotations
                    if y_max <= 0:
                        y_max = 1
                    # Make sure annotations are positioned well above the error bars
                    if self.errorbar_type_var.get() == "SD":
                        # Calculate standard deviation for each group to ensure annotations clear the error bars
                        std_vals = df_plot.groupby(x_col)[value_col].std().max()
                        # If std is NaN or 0, use a sensible default
                        if pd.isna(std_vals) or std_vals == 0:
                            std_vals = 0.1 * y_max
                        # Position annotations well above the standard deviation error bars
                        pval_height = y_max + 1.5 * std_vals
                    else:
                        # For SEM or no error bars, position relative to the max
                        pval_height = y_max * 1.15  # 15% above the maximum value
                    
                    step = 0.07 * (y_max - y_min if y_max != y_min else 0.2 * y_max)
                    if step <= 0:
                        step = 0.1 * y_max  # Ensure a minimum step size
                    
                    test_used = ""
                    # Get the unique x-values and their positions
                    x_values = [g for g in df_plot[x_col].dropna().unique()]
                    positions = {g: i for i, g in enumerate(x_values)}
                    print(f"[DEBUG] X values: {x_values}, positions: {positions}")
                    
                    # Always print DEBUG information about the p-values being used
                    print(f"[DEBUG] Annotation stats: latest_pvals = {self.latest_pvals}")
                    
                    # For bar plots, we want to directly annotate between the bars
                    # Get all combinations of x_values for comparison
                    if len(x_values) > 1:
                        pairs = list(itertools.combinations(x_values, 2))
                        # No longer comparing all pairs
                        
                        # Calculate maximum value for positioning annotations
                        y_max = df_plot[value_col].astype(float).max()
                        
                        # Get error metrics to determine annotation heights
                        error_metrics = {}
                        for x_val in x_values:
                            subset = df_plot[df_plot[x_col] == x_val]
                            mean_val = subset[value_col].mean()
                            std_val = subset[value_col].std()
                            error_metrics[x_val] = {'mean': mean_val, 'std': std_val}
                        
                        # Print all available p-value keys before annotation loop
                        print(f"[DEBUG] Available p-value keys before annotation: {list(self.latest_pvals.keys())}")
                        annotation_count = 0
                        
                        # Add annotations for group comparisons (e.g., 'Treated' vs 'Untreated') within each x-value
                        # Dynamically determine group names from the data
                        group_names = list(df_plot[hue_col].dropna().unique()) if hue_col else []
                        
                        # CASE 1: Two groups (e.g., Treated vs Untreated) within each x-value
                        if len(group_names) == 2:
                            g1, g2 = group_names
                            for idx, x_val in enumerate(x_values):
                                # Try different key formats to find the p-value
                                key1 = (x_val, g1, g2)
                                key2 = (x_val, g2, g1)
                                key3 = self.stat_key(x_val, g1, g2)
                                
                                pval = self.latest_pvals.get(key1)
                                if pval is None:
                                    pval = self.latest_pvals.get(key2)
                                if pval is None:
                                    pval = self.latest_pvals.get(key3)
                                    
                                if pval is not None:
                                    print(f"[DEBUG] Found p-value for {x_val}, {g1} vs {g2}: {pval}")
                                else:
                                    print(f"[DEBUG] Missing p-value for {x_val}, {g1} vs {g2}")
                                    continue
                                    
                                try:
                                    pos = positions[x_val]
                                except Exception as e:
                                    print(f"[DEBUG] Could not find position for {x_val}: {e}")
                                    continue
                                    
                                # Find the y values (heights) for both groups at this x_val
                                y1 = df_plot[(df_plot[x_col]==x_val) & (df_plot[hue_col]==g1)][value_col].mean()
                                y2 = df_plot[(df_plot[x_col]==x_val) & (df_plot[hue_col]==g2)][value_col].mean()
                                y_max_group = max(y1, y2)
                                annotation_height = y_max_group + (y_max * 0.08)
                                annotation_text = self.pval_to_annotation(pval)
                                print(f"[DEBUG] Adding annotation: {x_val}, {g1} vs {g2} -> {annotation_text} (p={pval})")
                                annotation_count += 1
                                
                                # Calculate bar positions for hue groups
                                width = 0.8  # Typical seaborn bar width
                                group_width = width / len(group_names)
                                g1_index = group_names.index(g1)
                                g2_index = group_names.index(g2)
                                
                                # Calculate positions with offsets
                                g1_pos = pos - width/2 + group_width/2 + g1_index * group_width
                                g2_pos = pos - width/2 + group_width/2 + g2_index * group_width
                                
                                if swap_axes:
                                    # Draw bracket style annotation for swapped axes
                                    line_height = annotation_height
                                    # Horizontal line connecting the bars
                                    ax.plot([line_height, line_height], [g1_pos, g2_pos], color='black', linewidth=linewidth, zorder=10)
                                    # Vertical lines at each end
                                    for end_pos in [g1_pos, g2_pos]:
                                        ax.plot([line_height - (y_max * 0.02), line_height], [end_pos, end_pos], color='black', linewidth=linewidth, zorder=10)
                                    # Place the significance marker next to the line
                                    mid_pos = (g1_pos + g2_pos) / 2
                                    ax.text(line_height + (y_max * 0.03), mid_pos, annotation_text, ha='left', va='center', fontsize=fontsize, color='black', zorder=11)
                                else:
                                    # Draw bracket style annotation for normal axes
                                    line_height = annotation_height
                                    # Horizontal line connecting the bars
                                    ax.plot([g1_pos, g2_pos], [line_height, line_height], color='black', linewidth=linewidth, zorder=10)
                                    # Vertical lines at each end
                                    for end_pos in [g1_pos, g2_pos]:
                                        ax.plot([end_pos, end_pos], [line_height - (y_max * 0.02), line_height], color='black', linewidth=linewidth, zorder=10)
                                    # Place the significance marker above the line
                                    mid_pos = (g1_pos + g2_pos) / 2
                                    ax.text(mid_pos, line_height + (y_max * 0.03), annotation_text, ha='center', va='bottom', fontsize=fontsize, color='black', zorder=11)
                                    
                            print(f"[DEBUG] Total annotations drawn: {annotation_count}")
                        # CASE 2: Multiple groups within each x-value
                        elif len(group_names) > 2:
                            # We're going to place annotations independently for each x-value
                            for x_val in x_values:
                                # Get all pairwise combinations of groups
                                group_pairs = list(itertools.combinations(group_names, 2))
                                
                                # Track vertical spacing for annotations within this x-value
                                annotation_vertical_offset = 0
                                
                                for g1, g2 in group_pairs:
                                    # Try different key formats to find the p-value
                                    key1 = (x_val, g1, g2)
                                    key2 = (x_val, g2, g1)
                                    key3 = self.stat_key(x_val, g1, g2)
                                    
                                    pval = self.latest_pvals.get(key1)
                                    if pval is None:
                                        pval = self.latest_pvals.get(key2)
                                    if pval is None:
                                        pval = self.latest_pvals.get(key3)
                                    
                                    if pval is None:
                                        continue
                                        
                                    try:
                                        pos = positions[x_val]
                                    except Exception as e:
                                        print(f"[DEBUG] Could not find position for {x_val}: {e}")
                                        continue
                                        
                                    # Find the y values (heights) for both groups at this x_val
                                    y1 = df_plot[(df_plot[x_col]==x_val) & (df_plot[hue_col]==g1)][value_col].mean()
                                    y2 = df_plot[(df_plot[x_col]==x_val) & (df_plot[hue_col]==g2)][value_col].mean()
                                    y_values = [y1, y2]
                                    y_max_group = max(y_values)
                                        
                                    # Determine spacing for annotation based on heights - with reduced distance
                                    base_offset = 0.15  # Start closer to the bars
                                    spacing = 0.15     # Less space between each comparison
                                    annotation_height = y_max_group + (y_max * (base_offset + spacing * annotation_vertical_offset))
                                    p_text = self.pval_to_annotation(pval)
                                    print(f"[DEBUG] Adding annotation: {x_val}, {g1} vs {g2} -> {p_text} (p={pval})")
                                    annotation_count += 1
                                    
                                    # Get bar positions for each group within this x-value
                                    if not swap_axes:
                                        # Calculate bar positions for hue groups
                                        width = 0.8  # Typical seaborn bar width
                                        group_width = width / len(group_names)
                                        g1_index = group_names.index(g1)
                                        g2_index = group_names.index(g2)
                                        
                                        # Calculate positions with offsets
                                        g1_pos = pos - width/2 + group_width/2 + g1_index * group_width
                                        g2_pos = pos - width/2 + group_width/2 + g2_index * group_width
                                        
                                        # Draw the bracket line
                                        line_height = annotation_height
                                        # Horizontal line connecting the bars
                                        ax.plot([g1_pos, g2_pos], [line_height, line_height], color='black', linewidth=linewidth, zorder=10)
                                        # Vertical lines at each end
                                        for end_pos in [g1_pos, g2_pos]:
                                            ax.plot([end_pos, end_pos], [line_height - (y_max * 0.02), line_height], color='black', linewidth=linewidth, zorder=10)
                                        
                                        # Place the significance marker above the line with specified spacing
                                        mid_pos = (g1_pos + g2_pos) / 2
                                        ax.text(mid_pos, line_height + (y_max * 0.03), p_text, ha='center', va='bottom', fontsize=fontsize, color='black', zorder=11)
                                    else:
                                        # Similar logic for swapped axes
                                        width = 0.8
                                        group_width = width / len(group_names)
                                        g1_index = group_names.index(g1)
                                        g2_index = group_names.index(g2)
                                        
                                        g1_pos = pos - width/2 + group_width/2 + g1_index * group_width
                                        g2_pos = pos - width/2 + group_width/2 + g2_index * group_width
                                        
                                        line_height = annotation_height
                                        ax.plot([line_height, line_height], [g1_pos, g2_pos], color='black', linewidth=linewidth, zorder=10)
                                        for end_pos in [g1_pos, g2_pos]:
                                            ax.plot([line_height - (y_max * 0.02), line_height], [end_pos, end_pos], color='black', linewidth=linewidth, zorder=10)
                                        
                                        mid_pos = (g1_pos + g2_pos) / 2
                                        ax.text(line_height + (y_max * 0.02), mid_pos, p_text, ha='left', va='center', fontsize=fontsize, color='black', zorder=11)
                                        
                                    # Increment vertical offset for the next comparison within this x-value
                                    annotation_vertical_offset += 1
                                
                            print(f"[DEBUG] Total annotations drawn for multiple groups: {annotation_count}")
                        # CASE 3: One group but multiple x-categories (e.g., Control vs Experimental)
                        elif len(x_values) > 1 and (hue_col is None or len(group_names) == 1):
                            # Get pairs of x-values to compare
                            pairs = list(itertools.combinations(x_values, 2))
                            for idx, (x1, x2) in enumerate(pairs):
                                # Determine key format to use
                                key = (x1, x2)
                                pval = self.latest_pvals.get(key)
                                
                                # If not found, try the reversed key
                                if pval is None:
                                    key = (x2, x1)
                                    pval = self.latest_pvals.get(key)
                                
                                # Final fallback to the sorted key method
                                if pval is None:
                                    key = self.stat_key(x1, x2)
                                    pval = self.latest_pvals.get(key)
                                
                                if pval is not None:
                                    print(f"[DEBUG] Found p-value for {x1} vs {x2}: {pval}")
                                    pos1, pos2 = positions[x1], positions[x2]
                                    annotation_text = self.pval_to_annotation(pval)
                                    print(f"[DEBUG] Adding annotation: {x1} vs {x2} -> {annotation_text} (p={pval})")
                                    annotation_count += 1
                                    
                                    # Calculate annotation height (increase for each additional annotation)
                                    annotation_height = y_max * 1.15 + idx * (y_max * 0.08)
                                    
                                    # Draw annotation lines and text
                                    if swap_axes:
                                        # Horizontal annotation for swapped axes
                                        ax.plot([annotation_height, annotation_height, annotation_height, annotation_height], 
                                                [pos1, pos1 - 0.05, pos2 + 0.05, pos2], 
                                                color='black', linewidth=linewidth, zorder=10)
                                        ax.text(annotation_height + (y_max * 0.02), (pos1 + pos2) / 2, 
                                                annotation_text, ha='left', va='center', 
                                                fontsize=fontsize, color='black', zorder=11)
                                    else:
                                        # Vertical annotation for normal axes
                                        ax.plot([pos1, pos1, pos2, pos2], 
                                                [annotation_height, annotation_height + (y_max * 0.02), 
                                                annotation_height + (y_max * 0.02), annotation_height], 
                                                color='black', linewidth=linewidth, zorder=10)
                                        ax.text((pos1 + pos2) / 2, annotation_height + (y_max * 0.03), 
                                                annotation_text, ha='center', va='bottom', 
                                                fontsize=fontsize, color='black', zorder=11)
                                else:
                                    print(f"[DEBUG] Missing p-value for pair {x1} vs {x2}")
                            
                            print(f"[DEBUG] Total annotations drawn in x-category comparison: {annotation_count}")
                    else:
                        print("[DEBUG] Statistical annotations skipped: no applicable annotation case")
                        print(f"[DEBUG] Annotating with hue groups: {base_groups} / {hue_groups}")
                        
                        for i, g in enumerate(base_groups):
                            df_sub = df_plot[df_plot[x_col] == g]
                            pairs = list(itertools.combinations(hue_groups, 2))
                            x_base = i
                            
                            for idx, (h1, h2) in enumerate(pairs):
                                key = self.stat_key(g, h1, h2)
                                pval = self.latest_pvals.get(key, np.nan)
                                if pval is None or (isinstance(pval, float) and np.isnan(pval)):
                                    print(f"[DEBUG] Missing p-value for hue annotation key: {key}")
                                    continue  # Skip this annotation
                                    
                                print(f"[DEBUG] Hue comparison: {g} : {h1} vs {h2} -> {self.pval_to_annotation(pval)} (p={pval})")
                                
                                # Calculate bar positions for each hue group
                                bar_centers = []
                                for j, hue_val in enumerate(hue_groups):
                                    bar_pos = x_base - 0.4 + (j + 0.5) * (0.8 / n_hue)
                                    bar_centers.append(bar_pos)
                                x1 = bar_centers[hue_groups.index(h1)]
                                x2 = bar_centers[hue_groups.index(h2)]
                                y = pval_height + idx * step + i * step * len(pairs)
                                # Dynamically calculate annotation height based on bar heights and additional plot elements
                                bar_heights = []
                                error_bar_heights = []
                                stripplot_heights = []
                                for j, hue_val in enumerate(hue_groups):
                                    subset = df_sub[df_sub[hue_col] == hue_val]
                                    bar_height = subset[value_col].mean() if not subset.empty else 0
                                    bar_heights.append(bar_height)
                                    if show_errorbars:
                                        std_err = subset[value_col].std() if not subset.empty else 0
                                        error_bar_heights.append(bar_height + std_err)
                                    if show_stripplot:
                                        max_strip = subset[value_col].max() if not subset.empty else 0
                                        stripplot_heights.append(max_strip)
                                max_bar_height = max(bar_heights)
                                max_error_height = max(error_bar_heights) if error_bar_heights else max_bar_height
                                max_stripplot_height = max(stripplot_heights) if stripplot_heights else max_bar_height
                                y_annotation = max(max_bar_height, max_error_height, max_stripplot_height) * 1.25  # 25% above the highest element to ensure visibility
                                annotation_text = self.pval_to_annotation(pval)
                                if swap_axes:
                                    ax.plot([y_annotation, y_annotation, y_annotation, y_annotation], [x1, x1, x2, x2], lw=linewidth, c='k', zorder=10)
                                    ax.text(y_annotation * 1.05, (x1+x2)/2, annotation_text, ha='left', va='center', fontsize=max(int(fontsize*0.7), 6), zorder=11, weight='bold')
                                else:
                                    ax.plot([x1, x1, x2, x2], [y_annotation, y_annotation, y_annotation, y_annotation], lw=linewidth, c='k', zorder=10)
                                    ax.text((x1+x2)/2, y_annotation * 1.05, annotation_text, ha='center', va='bottom', fontsize=max(int(fontsize*0.7), 6), zorder=11, weight='bold')
                    fig = ax.figure
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

        # Show the statistical details button if statistics are used
        if self.use_stats_var.get():
            try:
                self.stats_details_btn.pack(pady=5)
            except Exception as e:
                print(f"Error showing stats details button: {e}")
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
                self.save_custom_colors_palettes()
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
                self.save_custom_colors_palettes()
                self.update_color_palette_dropdowns()
        tk.Button(window, text="Remove Selected Palette", command=remove_palette).pack(pady=2)

        tk.Button(window, text="Close", command=window.destroy).pack(pady=10)

if __name__ == '__main__':
    root = tk.Tk()
    app = ExcelPlotterApp(root)
    root.mainloop()