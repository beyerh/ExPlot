"""
Statistical analysis module for ExPlot

This module provides a clean implementation of all statistical tests needed for ExPlot,
separated from the main UI code for better maintainability.
"""

import numpy as np
import pandas as pd
import itertools
import scipy.stats as stats
import pingouin as pg
import scikit_posthocs as sp
import warnings

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


def pval_to_annotation(p_val, alpha_levels=None):
    """
    Convert p-value to significance annotation.
    
    Args:
        p_val (float): The p-value to convert
        alpha_levels (list): List of significance thresholds [0.05, 0.01, 0.001, 0.0001]
        
    Returns:
        str: Significance annotation (ns, *, **, ***, ****)
    """
    if alpha_levels is None:
        alpha_levels = [0.05, 0.01, 0.001, 0.0001]
    
    if p_val > alpha_levels[0]:
        return "ns"  # Not significant
    elif p_val <= alpha_levels[3]:
        return "****"  # p ≤ 0.0001
    elif p_val <= alpha_levels[2]:
        return "***"   # p ≤ 0.001
    elif p_val <= alpha_levels[1]:
        return "**"    # p ≤ 0.01
    else:
        return "*"     # p ≤ 0.05


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
            groups = sorted(df[category_col].unique())
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
            posthoc = sp.posthoc_tamhane(df, val_col=value_col, group_col=category_col)
        
        # Ensure diagonal is 1.0 (no difference with self)
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
            groups = sorted(df[category_col].unique())
            return pd.DataFrame(1.0, index=groups, columns=groups)


def run_ttest(df, value_col, group1, group2, category_col, test_type="Independent t-test", alternative="two-sided"):
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
    try:
        g1 = df[df[category_col] == group1][value_col].dropna()
        g2 = df[df[category_col] == group2][value_col].dropna()
        
        if len(g1) < 2 or len(g2) < 2:
            return 1.0, None  # Not enough data
            
        if test_type == "Independent t-test":
            equal_var = True
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=equal_var, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2}
            
        elif test_type == "Welch's t-test":
            equal_var = False
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=equal_var, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2}
            
        elif test_type == "Paired t-test":
            # For paired t-test, we need equal length samples
            min_len = min(len(g1), len(g2))
            stat, p_val = stats.ttest_rel(g1[:min_len], g2[:min_len], alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": min_len - 1}
            
        elif test_type == "Mann-Whitney U test":
            # Non-parametric test
            stat, p_val = stats.mannwhitneyu(g1, g2, alternative=alternative)
            result = {"U": stat, "p-val": p_val}
            
        else:  # Fall back to independent t-test
            stat, p_val = stats.ttest_ind(g1, g2, equal_var=True, alternative=alternative)
            result = {"t": stat, "p-val": p_val, "df": len(g1) + len(g2) - 2}
            
        return p_val, result
        
    except Exception as e:
        debug(f"Error running t-test: {e}")
        return 1.0, None  # Return 1.0 (not significant) on error


def calculate_statistics(df_plot, x_col, value_col, hue_col=None, app_settings=None):
    """
    Master statistical calculation function for ExPlot.
    
    Args:
        df_plot (DataFrame): The data
        x_col (str): Column for x-axis categories
        value_col (str): Column for values to analyze
        hue_col (str, optional): Column for grouping
        app_settings (dict): Settings from the ExPlot app
        
    Returns:
        dict: Comprehensive results
    """
    # If no settings provided, use defaults
    if app_settings is None:
        app_settings = {
            'use_stats': True,
            'alpha_level': 0.05,
            'test_type': "Independent t-test",
            'alternative': "two-sided",
            'anova_type': "One-way ANOVA",
            'posthoc_type': "Tukey's HSD"
        }
    
    # ==========================================================
    # 1. INITIALIZATION & EARLY EXITS
    # ==========================================================
    
    # Initialize results dictionary
    results = {
        # Basic information
        'x_col': x_col,
        'value_col': value_col,
        'hue_col': hue_col,
        'alpha_level': app_settings.get('alpha_level', 0.05),
        
        # Test settings
        'test_type': app_settings.get('test_type', "Independent t-test"),
        'alternative': app_settings.get('alternative', "two-sided"),
        'anova_type': app_settings.get('anova_type', "One-way ANOVA"),
        'posthoc_type': app_settings.get('posthoc_type', "Tukey's HSD"),
        
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
    
    # Get x categories
    x_values = sorted([g for g in df_plot[x_col].dropna().unique()])
    results['x_values'] = x_values
    n_x_categories = len(x_values)
    results['n_x_categories'] = n_x_categories
    debug(f"X categories: {x_values} (count: {n_x_categories})")
    
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
        
        # For ungrouped data, generate all pairwise combinations
        pairs = list(itertools.combinations(x_values, 2))
        results['pairs'] = pairs
        debug(f"Generated {len(pairs)} pairwise comparisons")
        
        # For >2 categories, determine if we use ANOVA or individual t-tests
        use_anova = n_x_categories > 2 and app_settings.get('anova_type', "One-way ANOVA") != "None"
        
        if use_anova:
            # ANOVA with post-hoc tests for multiple categories
            debug(f"Analysis: ANOVA for {n_x_categories} categories")
            results['test_method'] = 'ANOVA'
            
            try:
                # Filter data if a single group is present
                if n_hue_groups == 1:
                    single_group = hue_values[0]
                    debug(f"Filtering for single group: {single_group}")
                    df_analysis = df_plot[df_plot[hue_col] == single_group].copy()
                else:
                    df_analysis = df_plot.copy()
                
                # Run ANOVA test
                anova_type = app_settings.get('anova_type', "One-way ANOVA")
                debug(f"Running {anova_type}")
                
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
                main_p = anova_results['p-unc'].values[0] if len(anova_results) > 0 else 1.0
                results['main_anova_p'] = main_p
                debug(f"ANOVA p-value: {main_p:.4g}")
                
                # Run post-hoc tests if ANOVA is significant
                if main_p <= alpha:
                    # Run appropriate post-hoc test
                    posthoc_type = app_settings.get('posthoc_type', "Tukey's HSD")
                    debug(f"Running post-hoc test: {posthoc_type}")
                    
                    # Get post-hoc p-values matrix
                    posthoc_matrix = run_posthoc(df_analysis, value_col, x_col, posthoc_type)
                    results['posthoc_matrix'] = posthoc_matrix
                    results['posthoc_type'] = posthoc_type
                    debug(f"Post-hoc test completed with {posthoc_matrix.shape[0]} groups")
                    
                    # Store p-values for all pairwise comparisons in both formats
                    for g1 in x_values:
                        for g2 in x_values:
                            if g1 != g2 and g1 in posthoc_matrix.index and g2 in posthoc_matrix.columns:
                                # Get p-value from matrix
                                p_val = posthoc_matrix.loc[g1, g2]
                                
                                # Store with multiple key formats for robustness
                                key = make_stat_key(g1, g2)
                                results['pvals'][key] = p_val
                                results['pvals'][(g1, g2)] = p_val
                                results['pvals'][(g2, g1)] = p_val
                else:
                    debug("ANOVA not significant, skipping post-hoc tests")
                    results['summary'] = f"ANOVA not significant (p = {main_p:.4g}). No post-hoc tests performed."
                    
            except Exception as e:
                debug(f"Error in ANOVA analysis: {e}")
                # Fall back to pairwise t-tests on error
                use_anova = False
                results['summary'] = f"Error in ANOVA: {str(e)}. Falling back to pairwise t-tests."
        
        # If not using ANOVA or ANOVA failed, use pairwise t-tests
        if not use_anova:
            debug("Analysis: Individual pairwise t-tests")
            results['test_method'] = 'Pairwise t-tests'
            
            try:
                # Get test type and alternative
                test_type = app_settings.get('test_type', "Independent t-test")
                alternative = app_settings.get('alternative', "two-sided")
                debug(f"Using {test_type} with {alternative} alternative")
                
                # Filter data if a single group is present
                if n_hue_groups == 1:
                    single_group = hue_values[0]
                    debug(f"Filtering for single group: {single_group}")
                    df_analysis = df_plot[df_plot[hue_col] == single_group].copy()
                else:
                    df_analysis = df_plot.copy()
                
                # Run t-tests for all pairs
                for g1, g2 in pairs:
                    # Run t-test
                    p_val, test_result = run_ttest(df_analysis, value_col, g1, g2, x_col, 
                                                  test_type, alternative)
                    
                    # Store in both formats
                    key = make_stat_key(g1, g2)
                    results['pvals'][key] = p_val
                    results['pvals'][(g1, g2)] = p_val
                    results['pvals'][(g2, g1)] = p_val
                    results['test_results'][key] = test_result
                    
                    debug(f"T-test: {g1} vs {g2}: p = {p_val:.4g}")
            except Exception as e:
                debug(f"Error in t-test analysis: {e}")
                results['summary'] = f"Error performing t-tests: {str(e)}."
    
    # CASE B: Grouped data (multiple hue groups to compare within each x-value)
    elif n_hue_groups > 1:
        results['data_structure'] = 'grouped'
        results['comparison_type'] = 'within_groups'
        debug(f"Data structure: grouped ({n_hue_groups} groups within each x-category)")
        
        # Generate pairs for each hue group combination
        hue_pairs = list(itertools.combinations(hue_values, 2))
        debug(f"Generated {len(hue_pairs)} group pairs for comparison")
        
        # Store which test we're using
        test_type = app_settings.get('test_type', "Independent t-test")
        alternative = app_settings.get('alternative', "two-sided")
        debug(f"Using {test_type} with {alternative} alternative for group comparisons")
        
        # For each x-category, compare between hue groups
        for x_val in x_values:
            debug(f"Processing category: {x_val}")
            
            # Get data for this x-value
            df_category = df_plot[df_plot[x_col] == x_val].copy()
            
            # For each group pair within this x-value, run t-test
            for g1, g2 in hue_pairs:
                try:
                    # Run t-test
                    p_val, test_result = run_ttest(
                        df_category, value_col, g1, g2, hue_col, test_type, alternative
                    )
                    
                    # Store p-value with category-specific keys
                    key = make_stat_key(x_val, g1, g2)
                    results['pvals'][key] = p_val
                    results['pvals'][(x_val, g1, g2)] = p_val
                    results['pvals'][(x_val, g2, g1)] = p_val
                    results['test_results'][key] = test_result
                    
                    debug(f"T-test for {x_val}: {g1} vs {g2}: p = {p_val:.4g}")
                except Exception as e:
                    debug(f"Error in group t-test for {x_val}, {g1} vs {g2}: {e}")
                    # Continue with other comparisons
    
    return results
