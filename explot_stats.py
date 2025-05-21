"""
Statistical analysis module for ExPlot

This module provides a clean implementation of all statistical tests needed for ExPlot,
separated from the main UI code for better maintainability.
"""

import numpy as np
import pandas as pd
import itertools
import scipy.stats as stats
import math
import pingouin as pg
import scikit_posthocs as sp
import warnings
import traceback

# Suppress specific warnings from pingouin and scikit-posthocs
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning, message=".*invalid value.*")


def debug(msg):
    """Helper for consistent debug printing"""
    print(f"[DEBUG] {msg}")


def make_stat_key(*args):
    """
    Create a standardized key for statistical results storage
    
    For group comparisons, sorts the arguments for consistency
    unless they need to be kept in a specific order
    """
    if len(args) == 2:
        # For simple comparisons, sort the group names for consistent lookup
        return tuple(sorted([str(args[0]), str(args[1])]))
    elif len(args) == 3:
        # For category/group comparisons, keep the category (first arg)
        # but sort the group names (second and third args)
        cat = args[0]
        g1, g2 = sorted([str(args[1]), str(args[2])])
        return (cat, g1, g2)
    else:
        # For other cases, just return as is
        return args


def pval_to_annotation(p_val, alpha=0.05):
    """
    Convert p-value to significance annotation using the same threshold logic as in explot.py.
    
    Args:
        p_val (float): The p-value to convert
        alpha (float): The significance threshold (default: 0.05)
        
    Returns:
        str: Significance annotation (ns, *, **, ***, ****)
    """
    if p_val is None or (isinstance(p_val, float) and math.isnan(p_val)):
        return "?"
        
    if p_val > alpha:
        return "ns"  # Not significant
    elif p_val <= alpha/5000:  # 4 stars threshold
        return "****"
    elif p_val <= alpha/50:    # 3 stars threshold
        return "***"
    elif p_val <= alpha/5:     # 2 stars threshold
        return "**"
    elif p_val <= alpha:       # 1 star threshold
        return "*"


def format_pvalue_matrix(matrix):
    """
    Format a p-value matrix as a readable ASCII table with aligned columns and dashes on the diagonal.
    
    Args:
        matrix (DataFrame): Matrix of p-values
        
    Returns:
        str: Formatted ASCII table
    """
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
    sep = "-" * (idx_width + 2) + "+" + "+".join("-" * (w + 2) for w in zip(col_widths)) + "+"
    
    # Build rows
    rows = []
    for idx in formatted.index:
        row = f"{idx:>{idx_width}} | " + " | ".join(
            f"{formatted.loc[idx, col]:^{w}}" for col, w in zip(col_names, col_widths)
        ) + " |"
        rows.append(row)
    
    # Combine all parts
    table = header + "\n" + sep + "\n" + "\n".join(rows)
    return table


def run_anova(df, value_col, category_col, anova_type="One-way ANOVA", subject_col=None):
    """
    Run an ANOVA test based on the specified type.
    
    Args:
        df (DataFrame): The data
        value_col (str): Column with values to compare
        category_col (str): Column with categories
        anova_type (str): Type of ANOVA test
        subject_col (str, optional): Column with subject IDs for repeated measures
        
    Returns:
        DataFrame: ANOVA results
    """
    try:
        if anova_type == "Welch's ANOVA":
            result = pg.welch_anova(data=df, dv=value_col, between=category_col)
        elif anova_type == "Repeated measures ANOVA" and subject_col:
            result = pg.rm_anova(data=df, dv=value_col, within=category_col, subject=subject_col)
        else:
            result = pg.anova(data=df, dv=value_col, between=category_col)
        
        return result
    except Exception as e:
        debug(f"Error running ANOVA: {e}")
        # Fall back to one-way ANOVA if other methods fail
        try:
            return pg.anova(data=df, dv=value_col, between=category_col)
        except Exception as e2:
            debug(f"Fallback ANOVA also failed: {e2}")
            # Return empty DataFrame on complete failure
            return pd.DataFrame({'Source': ['Error'], 'SS': [0], 'DF': [0], 'MS': [0], 'F': [0], 'p-unc': [1.0]})


def run_posthoc(df, value_col, category_col, posthoc_type="Tukey's HSD"):
    """
    Run a post-hoc test based on the specified type.
    
    Args:
        df (DataFrame): The data
        value_col (str): Column with values to compare
        category_col (str): Column with categories
        posthoc_type (str): Type of post-hoc test
        
    Returns:
        DataFrame: Matrix of post-hoc p-values
    """
    try:
        if posthoc_type == "Tukey's HSD":
            # Using pingouin for Tukey's test
            posthoc = pg.pairwise_tukey(data=df, dv=value_col, between=category_col)
            
            # Convert to matrix format for consistency
            # CRITICAL: Preserve the original order from the dataframe
            groups = list(df[category_col].unique())
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
                
            return posthoc_matrix.astype(float)
            
        elif posthoc_type == "Scheffe's test":
            posthoc = sp.posthoc_scheffe(df, val_col=value_col, group_col=category_col)
        elif posthoc_type == "Dunn's test":
            posthoc = sp.posthoc_dunn(df, val_col=value_col, group_col=category_col)
        else:  # Default to Tamhane's T2
            # For Tamhane's T2, we need to process the result similar to Tukey's HSD
            # Ensure we're working with a Series when calling unique()
            if isinstance(category_col, str):
                groups = list(df[category_col].unique())
            else:
                # If category_col is not a string, try to get the first column if it's a DataFrame
                groups = list(df[category_col[0]].unique()) if hasattr(category_col, '__getitem__') else []
                
            posthoc = sp.posthoc_tamhane(df, val_col=value_col, group_col=category_col)
            
            # Convert to matrix format for consistency with Tukey's HSD
            posthoc_matrix = pd.DataFrame(index=groups, columns=groups)
            
            # Fill the matrix with p-values from the posthoc test
            for i, g1 in enumerate(groups):
                for j, g2 in enumerate(groups):
                    if i == j:
                        posthoc_matrix.loc[g1, g2] = 1.0  # Diagonal is 1.0 (no difference with self)
                    else:
                        try:
                            p_val = posthoc.loc[g1, g2]
                            # Ensure p-value is not exactly 0 (which can happen with floating point precision)
                            if p_val == 0:
                                # Use a very small non-zero value instead of 0
                                p_val = 1e-100
                            posthoc_matrix.loc[g1, g2] = p_val
                        except KeyError:
                            # If the comparison is missing, set to 1.0 (no significance)
                            posthoc_matrix.loc[g1, g2] = 1.0
            
            return posthoc_matrix.astype(float)
        
        # For other test types, ensure diagonal is 1.0 (no difference with self)
        for g in posthoc.index:
            if g in posthoc.columns:
                posthoc.loc[g, g] = 1.0
                
        return posthoc
        
    except Exception as e:
        debug(f"Error running post-hoc test: {e}")
        # Fall back to Tamhane's T2 if other methods fail
        try:
            posthoc = sp.posthoc_tamhane(df, val_col=value_col, group_col=category_col)
            for g in posthoc.index:
                if g in posthoc.columns:
                    posthoc.loc[g, g] = 1.0
            return posthoc
        except Exception as e2:
            debug(f"Fallback post-hoc test also failed: {e2}")
            # Return empty DataFrame if all else fails
            try:
                if isinstance(category_col, str):
                    groups = sorted(df[category_col].unique())
                else:
                    # If category_col is not a string, try to get the first column if it's a DataFrame
                    groups = sorted(df[category_col[0]].unique()) if hasattr(category_col, '__getitem__') else []
                return pd.DataFrame(1.0, index=groups, columns=groups)
            except Exception as e3:
                debug(f"Error creating fallback matrix: {e3}")
                return pd.DataFrame()


def run_ttest(df, value_col, group1, group2, category_col, test_type="Welch's t-test (unpaired, unequal variances)", alternative="two-sided"):
    """
    Run a t-test between two groups.
    
    Args:
        df (DataFrame): The data
        value_col (str): Column with values to compare
        group1 (str): First group to compare
        group2 (str): Second group to compare
        category_col (str): Column containing the groups
        test_type (str): Type of t-test
        alternative (str): Alternative hypothesis
        
    Returns:
        tuple: (p-value, test result object)
    """
    debug(f"Requested t-test type: '{test_type}' for {group1} vs {group2} in column {category_col}")
    actual_test_used = None
    
    try:
        # Get values for each group
        g1 = df[df[category_col] == group1][value_col].dropna().values
        g2 = df[df[category_col] == group2][value_col].dropna().values
        
        # Print group stats for debugging
        debug(f"Group {group1}: n={len(g1)}, mean={np.mean(g1):.4f}, std={np.std(g1):.4f}")
        debug(f"Group {group2}: n={len(g2)}, mean={np.mean(g2):.4f}, std={np.std(g2):.4f}")
        
        # Skip if not enough data
        if len(g1) < 2 or len(g2) < 2:
            debug(f"Not enough data for t-test between {group1} and {group2}")
            return 1.0, {"test_used": "Insufficient data", "p-val": 1.0}  # Return not significant if not enough data
        
        # Choose test based on test_type
        if test_type == "Paired t-test":
            if len(g1) != len(g2):
                debug(f"Cannot run paired t-test: group sizes differ ({len(g1)} vs {len(g2)})")
                actual_test_used = "Paired t-test (failed - unequal samples)"
                return 1.0, {"test_used": actual_test_used, "p-val": 1.0}  # Return not significant on error
                
            # Use pingouin for paired t-test which provides more comprehensive results
            debug("Executing Paired t-test using pingouin")
            actual_test_used = "Paired t-test"
            result_df = pg.ttest(g1, g2, paired=True, alternative=alternative)
            p_val = result_df['p-val'].iloc[0]
            result = result_df.to_dict('records')[0]
            result["test_used"] = actual_test_used
            
        elif test_type == "Student's t-test (unpaired, equal variances)":
            # Use scipy for Student's t-test
            debug("Executing Student's t-test (equal variances) using scipy")
            actual_test_used = "Student's t-test (equal variances)"
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=True, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2, "test_used": actual_test_used}
            
        elif test_type == "Welch's t-test (unpaired, unequal variances)":
            # Use scipy for Welch's t-test
            debug("Executing Welch's t-test (unequal variances) using scipy")
            actual_test_used = "Welch's t-test (unequal variances)"
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=False, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2, "test_used": actual_test_used}
            
        elif test_type == "Mann-Whitney U test (non-parametric)":
            # Use pingouin for Mann-Whitney test
            debug("Executing Mann-Whitney U test using pingouin")
            actual_test_used = "Mann-Whitney U test"
            result_df = pg.mwu(g1, g2, alternative=alternative)
            p_val = result_df['p-val'].iloc[0]
            result = result_df.to_dict('records')[0]
            result["test_used"] = actual_test_used
            
        else:  # Fall back to Welch's t-test (most robust default)
            debug(f"Unknown t-test type: '{test_type}', falling back to Welch's t-test")
            actual_test_used = f"Welch's t-test (fallback, unknown type: {test_type})"
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=False, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2, "test_used": actual_test_used}
            
        debug(f"COMPLETED TEST: {actual_test_used} between {group1} and {group2}: p = {p_val:.4f}")
        return p_val, result
        
    except Exception as e:
        debug(f"Error running t-test: {e}")
        traceback.print_exc()
        return 1.0, {"test_used": f"Error: {str(e)}", "p-val": 1.0}  # Return 1.0 (not significant) on error


def calculate_statistics(df_plot, x_col, value_col, hue_col=None, app_settings=None, comparison_type=None):
    """
    Master statistical calculation function for ExPlot.
    
    This function implements the statistical test logic as follows:
    
    For ungrouped data (no hue column or only one hue group):
    - With one category: No test performed
    - With two categories: Perform the selected t-test
    - With more than two categories: Perform ANOVA with post-hoc tests
    
    For grouped data (multiple hue groups):
    - With one group: Same as ungrouped data
    - With two groups: Perform t-test between the groups
    - With more than two groups: Perform ANOVA with post-hoc tests
    
    Args:
        df_plot (DataFrame): The data used for plotting
        x_col (str): Column for x-axis categories
        value_col (str): Column for values to analyze
        hue_col (str, optional): Column for grouping
        app_settings (dict): Settings from the ExPlot app
        comparison_type (str, optional): Type of comparison to perform. 
            Can be 'within_groups', 'across_categories', or None to show dialog.
        
    Returns:
        dict: Comprehensive results dictionary containing all statistical results
    """
    # If no settings provided, use defaults
    if app_settings is None:
        app_settings = {
            'use_stats': True,
            'alpha_level': 0.05,
            'test_type': "Welch's t-test (unpaired, unequal variances)",
            'alternative': "two-sided",
            'anova_type': "Welch's ANOVA",
            'posthoc_type': "Tamhane's T2"
        }
    
    # ==========================================================
    # 1. INITIALIZATION & EARLY EXITS
    # ==========================================================
    
    # Initialize storage for results
    results = {
        'pvals': {},
        'test_info': {},  # Store detailed test information
        'test_type': None,
        'posthoc_type': None,
        'anova_type': None,
        'anova_result': None,
        'x_col': x_col,
        'value_col': value_col,
        'hue_col': hue_col,
        'alpha_level': app_settings.get('alpha_level', 0.05),
        'test_settings': {
            'test_type': app_settings.get('test_type', "Welch's t-test (unpaired, unequal variances)"),
            'alternative': app_settings.get('alternative', "two-sided"),
            'anova_type': app_settings.get('anova_type', "Welch's ANOVA"),
            'posthoc_type': app_settings.get('posthoc_type', "Tamhane's T2"),
        },
        'posthoc_type': app_settings.get('posthoc_type', "Tamhane's T2"),
        
        # Results containers
        'pvals': {},                # P-values for annotations and display
        'test_results': {},         # Complete test result objects
        'raw_data': {},             # Raw data used for calculations
        'anova_results': None,      # ANOVA results if applicable
        'posthoc_results': None,    # Post-hoc test results
        'posthoc_matrix': None,     # Post-hoc test p-value matrix
        'comparison_type': None,    # Type of comparison performed
        'data_structure': None,     # Structure of the data
        'pairs': [],                # List of comparison pairs
        'x_values': [],             # List of x categories
        'hue_values': [],           # List of hue groups
        'main_anova_p': None,       # Main ANOVA p-value if applicable
        'summary': None,            # Summary of test results
    }
    
    # Exit if statistics are disabled
    if not app_settings.get('use_stats', True):
        debug("Statistics disabled, skipping calculations")
        results['summary'] = "Statistical calculations are disabled."
        return results
    
    # ==========================================================
    # 2. DATA PREPARATION
    # ==========================================================
    debug(f"Statistical analysis: x_col={x_col}, value_col={value_col}, hue_col={hue_col}")
    
    # Make a clean copy of the dataframe
    df_plot = df_plot.copy()
    
    # Convert value column to numeric
    try:
        df_plot[value_col] = pd.to_numeric(df_plot[value_col], errors='coerce')
        df_plot = df_plot.dropna(subset=[value_col])
        debug(f"Converted values to numeric, shape after cleaning: {df_plot.shape}")
    except Exception as e:
        debug(f"Error converting values to numeric: {e}")
        results['summary'] = f"Error preparing data: {str(e)}"
        return results
    
    # ==========================================================
    # 3. DATA STRUCTURE ANALYSIS
    # ==========================================================
    
    # Get x categories - preserve the order in the dataframe
    x_values = [g for g in df_plot[x_col].dropna().unique()]
    results['x_values'] = x_values
    n_x_categories = len(x_values)
    results['n_x_categories'] = n_x_categories
    debug(f"X categories (preserving order): {x_values} (count: {n_x_categories})")
    
    # Get hue groups if present
    if hue_col and hue_col in df_plot.columns:
        hue_values = sorted([g for g in df_plot[hue_col].dropna().unique()])
        n_hue_groups = len(hue_values)
        results['hue_values'] = hue_values
        debug(f"Hue groups: {hue_values} (count: {n_hue_groups})")
    else:
        hue_values = []
        n_hue_groups = 0
        hue_col = None  # Ensure it's None if not in columns
        results['hue_values'] = []
    
    # Skip calculations if we only have one category (nothing to compare)
    if n_x_categories <= 1 and n_hue_groups <= 1:
        debug("Only one category detected, skipping statistical calculations")
        results['summary'] = "Only one category detected. No statistical test performed."
        return results
    
    # ==========================================================
    # 4. STATISTICAL ANALYSIS
    # ==========================================================
    alpha = results['alpha_level']
    
    # CASE A: Ungrouped data (no hue groups or just one)
    if n_hue_groups <= 1:
        results['data_structure'] = 'ungrouped'
        results['comparison_type'] = 'x_categories'
        debug("Data structure: ungrouped (comparisons between x categories)")
        
        # Generate all pairwise combinations for comparison
        pairs = list(itertools.combinations(x_values, 2))
        results['pairs'] = pairs
        debug(f"Generated {len(pairs)} pairwise comparisons")
        
        # Clear the p-values dictionary
        results['pvals'] = {}
        
        # Determine test type based on number of categories
        if n_x_categories == 1:
            # With only one category, there's nothing to compare
            debug("Only one category detected, no tests needed")
            results['summary'] = "Only one X-axis category detected. No statistical test performed."
            return results
            
        elif n_x_categories == 2:
            # With exactly two categories, perform t-test
            debug("Two categories detected, performing t-test")
            results['test_method'] = 'T-test'
            
            # Get the test type and alternative from settings
            test_type = app_settings.get('test_type', "Welch's t-test (unpaired, unequal variances)")
            alternative = app_settings.get('alternative', "two-sided")
            
            debug(f"Using {test_type} with {alternative} alternative")
            
            # Get the two categories
            g1, g2 = x_values
            
            # Run t-test
            p_val, test_result = run_ttest(df_plot, value_col, g1, g2, x_col, test_type, alternative)
            
            # Store p-value in multiple formats for flexibility
            key = make_stat_key(g1, g2)
            results['pvals'][key] = p_val
            results['pvals'][(g1, g2)] = p_val
            results['pvals'][(g2, g1)] = p_val
            
            # Store detailed test information
            results['test_info'][key] = test_result
            results['test_info'][(g1, g2)] = test_result
            results['test_info'][(g2, g1)] = test_result
            
            debug(f"T-test result: p = {p_val:.4g}")
            alpha = app_settings.get('alpha_level', 0.05) if app_settings else 0.05
            results['summary'] = f"{test_type} result: p = {p_val:.4g} {pval_to_annotation(p_val, alpha=alpha)}"
            
        else:  # n_x_categories > 2
            # With more than two categories, perform ANOVA with post-hoc tests
            debug(f"Multiple categories ({n_x_categories}) detected, performing ANOVA")
            results['test_method'] = 'ANOVA with post-hoc'
            
            # Get the ANOVA and post-hoc test types from settings
            anova_type = app_settings.get('anova_type', "Welch's ANOVA")
            posthoc_type = app_settings.get('posthoc_type', "Tamhane's T2")
            
            debug(f"Using {anova_type} with {posthoc_type} post-hoc test")
        
            # Run the ANOVA analysis for multiple categories
            try:
                # Filter data if a single group is present
                if n_hue_groups == 1:
                    single_group = hue_values[0]
                    debug(f"Filtering for single group: {single_group}")
                    df_analysis = df_plot[df_plot[hue_col] == single_group].copy()
                else:
                    df_analysis = df_plot.copy()
                
                # Determine if we need a subject identifier for repeated measures
                subject_col = None
                if anova_type == "Repeated measures ANOVA":
                    # Try to find a subject column
                    for potential_col in df_plot.columns:
                        if 'subject' in potential_col.lower() or 'id' in potential_col.lower():
                            subject_col = potential_col
                            break
                            
                    if not subject_col:
                        # Create synthetic subject IDs
                        df_analysis['subject_id'] = np.arange(len(df_analysis))
                        subject_col = 'subject_id'
                        debug("Created synthetic subject IDs for repeated measures ANOVA")
                
                # Run the ANOVA test
                anova_results = run_anova(df_analysis, value_col, x_col, anova_type, subject_col)
                results['anova_results'] = anova_results
                results['anova_type'] = anova_type
                
                # Get the main p-value from ANOVA
                if 'p-unc' in anova_results.columns:
                    main_p = anova_results['p-unc'].values[0] if len(anova_results) > 0 else 1.0
                elif 'p' in anova_results.columns:
                    main_p = anova_results['p'].values[0] if len(anova_results) > 0 else 1.0
                else:
                    main_p = 1.0
                    
                results['main_anova_p'] = main_p
                debug(f"ANOVA p-value: {main_p:.4g}")
                
                # Calculate post-hoc tests regardless of ANOVA significance
                # This is important because some fields prioritize pairwise comparisons
                # even when the omnibus test is not significant
                debug(f"Running post-hoc test: {posthoc_type}")
                
                # Get post-hoc p-values matrix
                posthoc_matrix = run_posthoc(df_analysis, value_col, x_col, posthoc_type)
                results['posthoc_matrix'] = posthoc_matrix
                results['posthoc_type'] = posthoc_type
                debug(f"Post-hoc test completed with {posthoc_matrix.shape[0]} groups")
                
                # Store all pairwise p-values to ensure annotations match bar order
                debug(f"Storing p-values in exact dataframe order: {x_values}")
                
                # Generate a meaningful summary for the statistical details panel
                if main_p <= alpha:
                    alpha = app_settings.get('alpha_level', 0.05) if app_settings else 0.05
                    results['summary'] = f"{anova_type} result: F = {anova_results['F'].iloc[0]:.3f}, p = {main_p:.4g} {pval_to_annotation(main_p, alpha=alpha)}\n" \
                                        f"Post-hoc test: {posthoc_type}"
                else:
                    results['summary'] = f"{anova_type} result: F = {anova_results['F'].iloc[0]:.3f}, p = {main_p:.4g} (not significant)\n" \
                                        f"Post-hoc tests performed but should be interpreted with caution."
                
                # Store all pairwise p-values from the post-hoc matrix
                for i in range(len(x_values)):
                    for j in range(i+1, len(x_values)):
                        g1, g2 = x_values[i], x_values[j]
                        
                        if g1 in posthoc_matrix.index and g2 in posthoc_matrix.columns:
                            # Get p-value from matrix
                            p_val = posthoc_matrix.loc[g1, g2]
                            
                            # Store p-value with both ordered and sorted keys for compatibility
                            ordered_key = (g1, g2)  # Original order (critical for bar positions)
                            sorted_key = make_stat_key(g1, g2)  # Sorted key for flexible lookup
                            
                            results['pvals'][ordered_key] = p_val
                            results['pvals'][sorted_key] = p_val
                            results['pvals'][(g2, g1)] = p_val  # Also store reversed
                            
                            debug(f"Stored post-hoc p-value for {g1} vs {g2}: {p_val:.4g}")
                            
            except Exception as e:
                debug(f"Error in ANOVA analysis: {e}")
                traceback.print_exc()
                results['summary'] = f"Error in ANOVA: {str(e)}. Check your data structure."
        
    # CASE B: Grouped data (multiple hue groups)
    elif n_hue_groups > 1:
        results['data_structure'] = 'grouped'
        debug(f"Data structure: grouped ({n_hue_groups} groups within each x-category)")
        
        # Clear p-values to ensure a fresh start
        results['pvals'] = {}
        
        # Generate hue group pairs for comparison
        hue_pairs = list(itertools.combinations(hue_values, 2))
        debug(f"Generated {len(hue_pairs)} hue group pairs for comparison")
        
        # Get the test type and alternative from settings
        test_type = app_settings.get('test_type', "Welch's t-test (unpaired, unequal variances)")
        alternative = app_settings.get('alternative', "two-sided")
        debug(f"Using {test_type} with {alternative} alternative for group comparisons")
        
        # Determine comparison approach based on number of groups
        if n_hue_groups == 1:
            # With only one group, treat as ungrouped data (this should never happen here)
            debug("Only one hue group detected, should be handled by ungrouped case")
            results['comparison_type'] = 'x_categories'
            results['summary'] = "Only one group detected. Treating as ungrouped data."
            return results
            
        elif n_hue_groups == 2:
            # With exactly two groups, perform t-test within each x-category
            debug("Two groups detected, performing t-tests between groups within each category")
            results['comparison_type'] = 'within_groups'
            results['test_method'] = 'T-tests between groups'
            
            # For each x-category, compare the two groups
            all_comparisons = []
            
            for x_val in x_values:
                debug(f"Processing x-category: {x_val}")
                g1, g2 = hue_values
                
                # Get data for this x-category
                df_category = df_plot[df_plot[x_col] == x_val].copy()
                
                # Skip if either group has insufficient data
                if len(df_category[df_category[hue_col] == g1]) < 2 or len(df_category[df_category[hue_col] == g2]) < 2:
                    debug(f"Not enough data for {g1} vs {g2} in category {x_val}")
                    continue
                    
                # Run t-test between the two groups for this x-category
                p_val, test_result = run_ttest(df_category, value_col, g1, g2, hue_col, test_type, alternative)
                
                # Store p-value with keys that include the x-category
                key = make_stat_key(x_val, g1, g2)  # Sorted key
                results['pvals'][key] = p_val
                results['pvals'][(x_val, g1, g2)] = p_val  # Original order
                results['pvals'][(x_val, g2, g1)] = p_val  # Reversed order
                
                # Store test information consistently for all key formats
                results['test_info'][key] = test_result
                results['test_info'][(x_val, g1, g2)] = test_result  # Original order
                results['test_info'][(x_val, g2, g1)] = test_result  # Reversed order
                
                all_comparisons.append((x_val, g1, g2, p_val))
                debug(f"T-test for {x_val}: {g1} vs {g2}: p = {p_val:.4g}")
            
            # Generate a summary of the comparisons
            if all_comparisons:
                summary_parts = [f"Comparison type: {test_type} between groups within each category"]
                for x_val, g1, g2, p_val in all_comparisons:
                    alpha = app_settings.get('alpha_level', 0.05) if app_settings else 0.05
                    summary_parts.append(f"{x_val}: {g1} vs {g2}: p = {p_val:.4g} {pval_to_annotation(p_val, alpha=alpha)}")
                results['summary'] = "\n".join(summary_parts)
            else:
                results['summary'] = "No valid comparisons could be performed. Check your data."
                
        else:  # n_hue_groups > 2
            # With more than two groups, perform ANOVA within each x-category
            debug(f"Multiple groups ({n_hue_groups}) detected, performing ANOVA within each category")
            results['comparison_type'] = 'within_groups'
            results['test_method'] = 'ANOVA between groups'
            
            # Get the ANOVA and post-hoc test types from settings
            anova_type = app_settings.get('anova_type', "Welch's ANOVA")
            posthoc_type = app_settings.get('posthoc_type', "Tamhane's T2")
            
            # For each x-category, perform ANOVA and post-hoc tests on the groups
            all_comparisons = []
            
            for x_val in x_values:
                debug(f"Processing x-category: {x_val}")
                
                # Get data for this x-category
                df_category = df_plot[df_plot[x_col] == x_val].copy()
                
                # Skip if not enough data
                if len(df_category) < n_hue_groups + 1 or len(df_category[hue_col].unique()) < 2:
                    debug(f"Not enough data for ANOVA in category {x_val}")
                    continue
                    
                try:
                    # Run ANOVA for this category
                    anova_results = run_anova(df_category, value_col, hue_col, anova_type)
                    
                    # Get main p-value
                    if 'p-unc' in anova_results.columns:
                        main_p = anova_results['p-unc'].values[0] if len(anova_results) > 0 else 1.0
                    elif 'p' in anova_results.columns:
                        main_p = anova_results['p'].values[0] if len(anova_results) > 0 else 1.0
                    else:
                        main_p = 1.0
                        
                    all_comparisons.append((x_val, main_p))
                    debug(f"ANOVA for {x_val}: p = {main_p:.4g}")
                    
                    # Run post-hoc tests regardless of ANOVA significance
                    posthoc_matrix = run_posthoc(df_category, value_col, hue_col, posthoc_type)
                    
                    # Store all pairwise p-values
                    for i, g1 in enumerate(hue_values):
                        for j, g2 in enumerate(hue_values):
                            if i < j and g1 in posthoc_matrix.index and g2 in posthoc_matrix.columns:
                                p_val = posthoc_matrix.loc[g1, g2]
                                
                                # Store with category-specific keys
                                key = make_stat_key(x_val, g1, g2)  # Sorted key
                                results['pvals'][key] = p_val
                                results['pvals'][(x_val, g1, g2)] = p_val  # Original order
                                results['pvals'][(x_val, g2, g1)] = p_val  # Reversed order
                                
                                debug(f"Post-hoc for {x_val}: {g1} vs {g2}: p = {p_val:.4g}")
                                
                except Exception as e:
                    debug(f"Error in ANOVA for category {x_val}: {e}")
                    traceback.print_exc()
            
            # Generate a summary of the comparisons
            if all_comparisons:
                summary_parts = [f"Analysis: {anova_type} with {posthoc_type} post-hoc tests within each category"]
                for x_val, p_val in all_comparisons:
                    alpha = app_settings.get('alpha_level', 0.05) if app_settings else 0.05
                    summary_parts.append(f"{x_val}: ANOVA p = {p_val:.4g} {pval_to_annotation(p_val, alpha=alpha)}")
                results['summary'] = "\n".join(summary_parts)
            else:
                results['summary'] = "No valid ANOVA comparisons could be performed. Check your data."
    
    return results
    
    return results
