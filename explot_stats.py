"""
Statistical analysis module for ExPlot.

Single source of truth for all statistics: ``calculate_statistics()`` runs every
test exactly once and returns a :class:`StatsResult`.  The plot annotations and
the Statistical Details window both read from that object, so the symbols on
the graph and the numbers in the details window can never disagree.

No GUI code lives here.
"""

import itertools
import math
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
import scipy.stats as stats
import scikit_posthocs as sp


# ----------------------------------------------------------------------------
# Test options (the strings are stored in preferences / project files)
# ----------------------------------------------------------------------------

TTEST_OPTIONS = [
    "Student's t-test (unpaired, equal variances)",
    "Welch's t-test (unpaired, unequal variances)",
    "Paired t-test",
    "Mann-Whitney U test (non-parametric)",
    "Wilcoxon signed-rank test (non-parametric)",
]
ALTERNATIVE_OPTIONS = ["two-sided", "less", "greater"]
ANOVA_OPTIONS = [
    "One-way ANOVA",
    "Welch's ANOVA",
    "Repeated measures ANOVA",
    "Kruskal-Wallis H test (non-parametric)",
    "Friedman test (non-parametric)",
]
POSTHOC_OPTIONS = [
    "Tukey's HSD",
    "Games-Howell",
    "Tamhane's T2",
    "Scheffe's test",
    "Dunn's test",
    "Conover's test (non-parametric)",
    "Nemenyi test (non-parametric)",
]

GROUPED_TWO_WAY = "Two-way ANOVA + Šídák multiple comparisons"
GROUPED_SEPARATE_HOLM = "Separate tests per category (Holm-Šídák across categories)"
GROUPED_SEPARATE_RAW = "Separate tests per category (uncorrected)"
GROUPED_OPTIONS = [GROUPED_SEPARATE_HOLM, GROUPED_SEPARATE_RAW, GROUPED_TWO_WAY]
NO_SUBJECT = "None (match by row order)"

PAIRED_TWO_SAMPLE = {"Paired t-test", "Wilcoxon signed-rank test (non-parametric)"}
REPEATED_OMNIBUS = {"Repeated measures ANOVA", "Friedman test (non-parametric)"}
NONPARAMETRIC_OMNIBUS = {"Kruskal-Wallis H test (non-parametric)", "Friedman test (non-parametric)"}
NONPARAMETRIC_POSTHOC = {"Dunn's test", "Conover's test (non-parametric)", "Nemenyi test (non-parametric)"}
BLOCKED_POSTHOC = {"Conover's test (non-parametric)", "Nemenyi test (non-parametric)"}

TTEST_LABELS = {
    "Student's t-test (unpaired, equal variances)": "Student's t-test (unpaired, equal variances)",
    "Welch's t-test (unpaired, unequal variances)": "Welch's t-test (unpaired, unequal variances)",
    "Paired t-test": "Paired t-test",
    "Mann-Whitney U test (non-parametric)": "Mann-Whitney U test",
    "Wilcoxon signed-rank test (non-parametric)": "Wilcoxon signed-rank test",
}
ANOVA_LABELS = {
    "One-way ANOVA": "Ordinary one-way ANOVA",
    "Welch's ANOVA": "Welch's ANOVA",
    "Repeated measures ANOVA": "Repeated measures one-way ANOVA",
    "Kruskal-Wallis H test (non-parametric)": "Kruskal-Wallis test",
    "Friedman test (non-parametric)": "Friedman test",
}

DEFAULT_SETTINGS = {
    'alpha_level': 0.05,
    'test_type': "Welch's t-test (unpaired, unequal variances)",
    'alternative': "two-sided",
    'anova_type': "Welch's ANOVA",
    'posthoc_type': "Tamhane's T2",
    'grouped_analysis': GROUPED_SEPARATE_HOLM,
    'subject_col': None,
}


class StatsError(Exception):
    """A test could not be run on the given data (reported, never hidden)."""


# ----------------------------------------------------------------------------
# Significance symbols (one canonical definition)
# ----------------------------------------------------------------------------

def get_significance_thresholds(alpha=0.05):
    """(threshold, symbol) pairs from most to least significant.

    At alpha = 0.05 these are the GraphPad Prism levels:
    **** p <= 0.0001, *** p <= 0.001, ** p <= 0.01, * p <= 0.05.
    """
    return [
        (alpha / 500, '****'),
        (alpha / 50, '***'),
        (alpha / 5, '**'),
        (alpha, '*'),
    ]


def pval_to_annotation(p_val, alpha=0.05):
    """Convert a p-value to its significance symbol (ns, *, **, ***, ****)."""
    if p_val is None or not np.isfinite(p_val):
        return "n/a"
    for threshold, symbol in get_significance_thresholds(alpha):
        if p_val <= threshold:
            return symbol
    return "ns"


def get_statannotations_format(alpha=0.05):
    """``pvalue_format`` for statannotations' ``Annotator.configure()``."""
    return {
        'text_format': 'star',
        'pvalue_thresholds': list(get_significance_thresholds(alpha)) + [(1, 'ns')],
    }


def significance_table(alpha=0.05):
    """Rows (symbol, rule) of the asterisk lookup table."""
    rows = [(symbol, f"p ≤ {threshold:.5g}") for threshold, symbol in get_significance_thresholds(alpha)]
    rows.append(("ns", f"p > {alpha:.5g}"))
    return rows


def format_significance_legend(alpha=0.05):
    lines = ["Significance levels:"]
    lines += [f"  {symbol:<5} {rule}" for symbol, rule in significance_table(alpha)]
    return "\n".join(lines) + "\n"


def make_stat_key(*args):
    """Order-independent key for a comparison: (g1, g2) or (category, g1, g2)."""
    if len(args) == 2:
        return tuple(sorted(str(a) for a in args))
    if len(args) == 3:
        return (args[0],) + tuple(sorted(str(a) for a in args[1:]))
    return args


def format_p(p):
    """Format a p-value for display (exact value, scientific notation when small)."""
    if p is None or not np.isfinite(p):
        return "n/a"
    if p == 0:
        return "<1e-300"
    if p < 1e-4:
        return f"{p:.2e}"
    return f"{p:.4f}"


def _fmt(x, digits=4):
    if x is None:
        return "—"
    try:
        if np.isnan(x):
            return "—"
        if np.isinf(x):
            return "∞" if x > 0 else "−∞"
    except TypeError:
        return str(x)
    if float(x).is_integer() and abs(x) < 1e6:
        return str(int(x))
    return f"{x:.{digits}g}"


# ----------------------------------------------------------------------------
# Result objects
# ----------------------------------------------------------------------------

NAN = float('nan')


@dataclass
class Comparison:
    category: object
    group1: object
    group2: object
    test: str
    p_value: float = NAN
    significance: str = "n/a"
    adjustment: str = "none"
    p_unadjusted: float = NAN
    statistic_name: str = ""
    statistic: float = NAN
    df: float = NAN
    difference_name: str = ""
    difference: float = NAN
    ci_low: float = NAN
    ci_high: float = NAN
    effect_size_name: str = ""
    effect_size: float = NAN
    n1: int = 0
    n2: int = 0
    error: str = ""


@dataclass
class OmnibusTest:
    category: object
    test: str
    statistic_name: str = ""
    statistic: float = NAN
    df1: float = NAN
    df2: float = NAN
    p_value: float = NAN
    significance: str = "n/a"
    n_groups: int = 0
    n_total: int = 0
    extra: dict = field(default_factory=dict)
    error: str = ""


@dataclass
class Descriptive:
    category: object
    group: object
    n: int
    mean: float
    sd: float
    sem: float
    median: float
    minimum: float
    maximum: float


@dataclass
class StatsResult:
    x_col: str
    value_col: str
    hue_col: object
    alpha: float
    settings: dict
    value_label: str = ""
    x_label: str = ""
    hue_label: str = ""
    labels: dict = field(default_factory=dict)
    structure: str = ""  # 'ungrouped' | 'grouped'
    comparisons: list = field(default_factory=list)
    omnibus: list = field(default_factory=list)
    descriptives: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    error: str = ""

    # --- lookup -----------------------------------------------------------
    def label(self, value):
        if value is None:
            return ""
        try:
            return str(self.labels.get(value, value))
        except TypeError:
            return str(value)

    @property
    def has_results(self):
        return bool(self.comparisons or self.omnibus)

    @property
    def tests_run(self):
        seen = []
        for t in [o.test for o in self.omnibus] + [c.test for c in self.comparisons]:
            if t not in seen:
                seen.append(t)
        return seen

    def get_comparison(self, group1, group2, category=None):
        for c in self.comparisons:
            if self.structure == 'grouped' and not _same(c.category, category):
                continue
            if (_same(c.group1, group1) and _same(c.group2, group2)) or \
               (_same(c.group1, group2) and _same(c.group2, group1)):
                return c
        return None

    def get_p(self, group1, group2, category=None):
        c = self.get_comparison(group1, group2, category)
        return c.p_value if c is not None else None

    # --- tables -----------------------------------------------------------
    def to_tables(self):
        """All results as labelled DataFrames (full precision), for export."""
        grouped = self.structure == 'grouped'
        cat_col = f"Category ({self.x_label or self.x_col})"
        summary = pd.DataFrame([
            ("Values", self.value_label or self.value_col),
            ("Compared", self._compared_text()),
            ("Tests run", "; ".join(self.tests_run) or "none"),
            ("Alpha", self.alpha),
            ("Alternative hypothesis", self.settings.get('alternative')),
            ("Selected t-test", self.settings.get('test_type')),
            ("Selected ANOVA", self.settings.get('anova_type')),
            ("Selected post-hoc", self.settings.get('posthoc_type')),
            ("Grouped data analysis", self.settings.get('grouped_analysis') if grouped else "not applicable"),
            ("Subject column", self.settings.get('subject_col') or NO_SUBJECT),
        ] + [("Note", n) for n in self.notes] + [("Warning", w) for w in self.warnings]
          + ([("Error", self.error)] if self.error else []),
            columns=["Item", "Value"])

        omni = pd.DataFrame([{
            **({cat_col: self._cat(o.category)} if grouped else {}),
            "Test": o.test, "Statistic": o.statistic_name, "Value": o.statistic,
            "df1": o.df1, "df2": o.df2, "p": o.p_value, "Summary": o.significance,
            "Groups": o.n_groups, "N": o.n_total,
            **o.extra,
            "Error": o.error,
        } for o in self.omnibus])

        comps = pd.DataFrame([{
            **({cat_col: self.label(c.category)} if grouped else {}),
            "Group 1": self.label(c.group1), "Group 2": self.label(c.group2),
            "Test": c.test, "n1": c.n1, "n2": c.n2,
            "Statistic": c.statistic_name, "Value": c.statistic, "df": c.df,
            "Difference": c.difference_name, "Difference value": c.difference,
            "95% CI low": c.ci_low, "95% CI high": c.ci_high,
            "Effect size": c.effect_size_name, "Effect size value": c.effect_size,
            "p": c.p_value, "p (unadjusted)": c.p_unadjusted, "Multiplicity adjustment": c.adjustment,
            "Summary (on graph)": c.significance, "Error": c.error,
        } for c in self.comparisons])

        desc = pd.DataFrame([{
            **({cat_col: self.label(d.category)} if grouped else {}),
            "Group": self.label(d.group), "n": d.n, "Mean": d.mean, "SD": d.sd,
            "SEM": d.sem, "Median": d.median, "Min": d.minimum, "Max": d.maximum,
        } for d in self.descriptives])

        levels = pd.DataFrame(significance_table(self.alpha), columns=["Symbol", "Meaning"])
        return {
            "Summary": summary,
            "Overall tests (omnibus)": omni,
            "Pairwise comparisons": comps,
            "Descriptive statistics": desc,
            "Significance levels": levels,
        }

    def _cat(self, value):
        return "(all)" if value is None else self.label(value)

    def _compared_text(self):
        x = self.x_label or self.x_col
        if self.structure == 'grouped':
            return f"groups of '{self.hue_label or self.hue_col}' within each category of '{x}'"
        return f"categories of '{x}'"

    # --- text report ------------------------------------------------------
    def report(self):
        """Plain-text report, rendered by the Statistical Details window."""
        grouped = self.structure == 'grouped'
        out = ["STATISTICAL DETAILS", "=" * 19, ""]

        out += ["Data",
                f"  Values:       {self.value_label or self.value_col}",
                f"  Compared:     {self._compared_text()}",
                f"  Alpha:        {self.alpha:g}",
                f"  Alternative:  {self.settings.get('alternative')}"
                + ("  (applies to two-group tests; ANOVA and post-hoc tests are two-sided)"
                   if self.settings.get('alternative') != 'two-sided' else ""),
                ""]

        out.append("Tests actually run")
        out += [f"  - {t}" for t in self.tests_run] or ["  (none)"]
        out += ["", "Selected in settings",
                f"  t-test:   {self.settings.get('test_type')}",
                f"  ANOVA:    {self.settings.get('anova_type')}",
                f"  Post-hoc: {self.settings.get('posthoc_type')}"]
        if grouped:
            out.append(f"  Grouped:  {self.settings.get('grouped_analysis')}")
        out += [f"  Subject:  {self.settings.get('subject_col') or NO_SUBJECT}", ""]

        if self.error:
            out += ["ERROR", f"  {self.error}", ""]

        out.append("Significance levels (asterisk lookup, used on the graph)")
        out += [f"  {s:<5} {rule}" for s, rule in significance_table(self.alpha)]
        out.append("")

        if self.omnibus:
            out.append("Overall tests (omnibus: is there any difference among the groups at all?)")
            headers = (["Category"] if grouped else []) + ["Test", "Statistic", "df", "p", "Summary", "Groups", "N"]
            rows = []
            for o in self.omnibus:
                df_txt = _fmt(o.df1) if not np.isfinite(o.df2) else f"{_fmt(o.df1)}, {_fmt(o.df2)}"
                stat_txt = f"{o.statistic_name} = {_fmt(o.statistic)}" if o.statistic_name and np.isfinite(o.statistic) else "—"
                row = ([self._cat(o.category)] if grouped else []) + [
                    o.test, stat_txt, df_txt if np.isfinite(o.df1) else "—",
                    format_p(o.p_value), o.significance, str(o.n_groups), str(o.n_total)]
                rows.append(row)
            out += _table(headers, rows)
            for o in self.omnibus:
                where = f" [{self._cat(o.category)}]" if grouped else ""
                if o.extra:
                    out.append("  " + where.strip() + (" " if where else "") + ", ".join(
                        f"{k}: {format_p(v) if k.startswith('p') else _fmt(v)}" for k, v in o.extra.items()))
                if o.error:
                    out.append(f"  Not computed{where}: {o.error}")
            out.append("")

        if self.comparisons:
            out.append("Pairwise comparisons (these p-values and symbols are drawn on the graph)")
            show_raw = any(np.isfinite(c.p_unadjusted) for c in self.comparisons)
            headers = (["Category"] if grouped else []) + [
                "Comparison", "Test", "n", "Statistic", "df", "Difference", "95% CI",
                "Effect size"] + (["p (unadj.)"] if show_raw else []) + ["p", "Adjusted by", "Summary"]
            rows = []
            for c in self.comparisons:
                ci = f"{_fmt(c.ci_low)} to {_fmt(c.ci_high)}" if not np.isnan(c.ci_low) else "—"
                rows.append(([self.label(c.category)] if grouped else []) + [
                    f"{self.label(c.group1)} vs {self.label(c.group2)}", c.test,
                    f"{c.n1}, {c.n2}",
                    f"{c.statistic_name} = {_fmt(c.statistic)}" if c.statistic_name and np.isfinite(c.statistic) else "—",
                    _fmt(c.df),
                    f"{c.difference_name} = {_fmt(c.difference)}" if c.difference_name and np.isfinite(c.difference) else "—",
                    ci,
                    f"{c.effect_size_name} = {_fmt(c.effect_size, 3)}" if c.effect_size_name and np.isfinite(c.effect_size) else "—"]
                    + ([format_p(c.p_unadjusted)] if show_raw else [])
                    + [format_p(c.p_value), c.adjustment, c.significance])
            out += _table(headers, rows)
            errors = [c for c in self.comparisons if c.error]
            for c in errors:
                where = f"{self.label(c.category)}: " if grouped else ""
                out.append(f"  Not computed – {where}{self.label(c.group1)} vs {self.label(c.group2)}: {c.error}")
            if any(c.difference_name for c in self.comparisons):
                out.append("  Difference = Group 1 − Group 2.")
            out.append("")

        if self.descriptives:
            out.append("Descriptive statistics")
            headers = (["Category"] if grouped else []) + ["Group", "n", "Mean", "SD", "SEM", "Median", "Min", "Max"]
            rows = [([self.label(d.category)] if grouped else []) + [
                self.label(d.group), str(d.n), _fmt(d.mean), _fmt(d.sd), _fmt(d.sem),
                _fmt(d.median), _fmt(d.minimum), _fmt(d.maximum)] for d in self.descriptives]
            out += _table(headers, rows)
            out.append("")

        if self.warnings:
            out.append("Warnings")
            out += [f"  ! {w}" for w in self.warnings]
            out.append("")
        if self.notes:
            out.append("Notes")
            out += [f"  - {n}" for n in self.notes]
            out.append("")
        return "\n".join(out)


def _table(headers, rows, indent="  "):
    widths = [max(len(str(h)), *(len(str(r[i])) for r in rows)) if rows else len(str(h))
              for i, h in enumerate(headers)]
    line = lambda cells: indent + " | ".join(f"{str(c):<{w}}" for c, w in zip(cells, widths))
    return [line(headers), indent + "-+-".join("-" * w for w in widths)] + [line(r) for r in rows]


def _same(a, b):
    if a is None or b is None:
        return a is None and b is None
    try:
        if a == b:
            return True
    except Exception:
        pass
    return str(a) == str(b)


# ----------------------------------------------------------------------------
# Data helpers
# ----------------------------------------------------------------------------

def _level_order(series, order=None):
    """Order of levels as drawn by seaborn (explicit order > categorical > sorted numeric > appearance)."""
    s = series.dropna()
    if isinstance(s.dtype, pd.CategoricalDtype):
        present = set(s.unique().tolist())
        levels = [c for c in s.cat.categories if c in present]
    else:
        levels = list(pd.unique(s))
        if pd.api.types.is_numeric_dtype(s):
            levels = sorted(levels)
    if order:
        ordered = [v for v in order if any(_same(v, lv) for lv in levels)]
        ordered = [next(lv for lv in levels if _same(v, lv)) for v in ordered]
        levels = ordered + [lv for lv in levels if not any(_same(lv, o) for o in ordered)]
    return levels


def _clean(a):
    a = np.asarray(a, dtype=float)
    return a[np.isfinite(a)]


def _paired_matrix(raw, names, subjects=None):
    """Values of all groups as a (subjects x groups) matrix for paired / repeated-measures tests.

    With ``subjects`` (one array of IDs per group) values are matched by subject ID;
    otherwise by their position within each group (row order in the data).
    Subjects (rows) with a missing value in any group are dropped.
    """
    if subjects is not None:
        cols = []
        for a, ids, name in zip(raw, subjects, names):
            ser = pd.Series(np.asarray(a, dtype=float), index=pd.Index(ids)).dropna()
            ser = ser[ser.index.notna()]
            dup = ser.index[ser.index.duplicated()]
            if len(dup):
                raise StatsError(f"subject '{dup[0]}' has more than one value in group '{name}'")
            cols.append(ser)
        m = pd.concat(cols, axis=1, join='inner')
        if m.empty:
            raise StatsError("no subject has a value in every group")
        n_all = len(set().union(*(set(c.index) for c in cols)))
        return m.to_numpy(dtype=float), n_all - len(m)
    lengths = [len(a) for a in raw]
    if len(set(lengths)) != 1:
        detail = ", ".join(f"{n}: {l}" for n, l in zip(names, lengths))
        raise StatsError(f"paired/repeated-measures tests need the same number of rows in every group ({detail}); "
                         "set a subject column to match values by ID")
    m = np.column_stack([np.asarray(a, dtype=float) for a in raw])
    keep = np.isfinite(m).all(axis=1)
    return m[keep], int((~keep).sum())


def _describe(category, group, values):
    v = _clean(values)
    n = len(v)
    sd = float(np.std(v, ddof=1)) if n > 1 else NAN
    return Descriptive(
        category=category, group=group, n=n,
        mean=float(np.mean(v)) if n else NAN, sd=sd,
        sem=sd / math.sqrt(n) if n > 1 else NAN,
        median=float(np.median(v)) if n else NAN,
        minimum=float(np.min(v)) if n else NAN,
        maximum=float(np.max(v)) if n else NAN,
    )


# ----------------------------------------------------------------------------
# Two-group tests
# ----------------------------------------------------------------------------

def _two_sample(a_raw, b_raw, test_type, alternative, category, g1, g2, names, subjects=None):
    if test_type not in TTEST_LABELS:
        return Comparison(category, g1, g2, test=str(test_type), error=f"unknown test '{test_type}'")
    c = Comparison(category, g1, g2, test=TTEST_LABELS[test_type])
    try:
        if test_type in PAIRED_TWO_SAMPLE:
            m, _ = _paired_matrix([a_raw, b_raw], names, subjects)
            a, b = m[:, 0], m[:, 1]
        else:
            a, b = _clean(a_raw), _clean(b_raw)
        c.n1, c.n2 = len(a), len(b)
        if min(c.n1, c.n2) < 2:
            raise StatsError(f"needs at least 2 values per group (n = {c.n1}, {c.n2})")

        if test_type in ("Student's t-test (unpaired, equal variances)", "Welch's t-test (unpaired, unequal variances)"):
            r = stats.ttest_ind(a, b, equal_var=test_type.startswith("Student"), alternative=alternative)
            ci = r.confidence_interval(0.95)
            c.statistic_name, c.df = "t", float(r.df)
            c.difference_name, c.difference = "Mean diff", float(np.mean(a) - np.mean(b))
            c.ci_low, c.ci_high = float(ci.low), float(ci.high)
            pooled = math.sqrt(((c.n1 - 1) * np.var(a, ddof=1) + (c.n2 - 1) * np.var(b, ddof=1)) / (c.n1 + c.n2 - 2))
            c.effect_size_name, c.effect_size = "Cohen's d", c.difference / pooled if pooled > 0 else NAN
        elif test_type == "Paired t-test":
            r = stats.ttest_rel(a, b, alternative=alternative)
            ci = r.confidence_interval(0.95)
            d = a - b
            c.statistic_name, c.df = "t", float(r.df)
            c.difference_name, c.difference = "Mean of differences", float(np.mean(d))
            c.ci_low, c.ci_high = float(ci.low), float(ci.high)
            sd = np.std(d, ddof=1)
            c.effect_size_name, c.effect_size = "Cohen's dz", float(np.mean(d) / sd) if sd > 0 else NAN
        elif test_type == "Mann-Whitney U test (non-parametric)":
            r = stats.mannwhitneyu(a, b, alternative=alternative, method='auto')
            c.statistic_name = "U"
            c.difference_name, c.difference = "Median diff", float(np.median(a) - np.median(b))
            c.effect_size_name, c.effect_size = "Rank-biserial r", float(2 * r.statistic / (c.n1 * c.n2) - 1)
        else:  # Wilcoxon signed-rank
            r = stats.wilcoxon(a, b, alternative=alternative)
            c.statistic_name = "W"
            c.difference_name, c.difference = "Median of differences", float(np.median(a - b))

        c.statistic, c.p_value = float(r.statistic), float(r.pvalue)
        if not np.isfinite(c.p_value):
            raise StatsError("the test returned no p-value (e.g. zero variance or all differences zero)")
    except Exception as e:
        c.p_value = NAN
        c.error = str(e)
    return c


# ----------------------------------------------------------------------------
# Omnibus tests (k > 2 groups)
# ----------------------------------------------------------------------------

def _welch_anova(samples):
    k = len(samples)
    n = np.array([len(s) for s in samples], dtype=float)
    means = np.array([np.mean(s) for s in samples])
    var = np.array([np.var(s, ddof=1) for s in samples])
    if np.any(var <= 0):
        raise StatsError("Welch's ANOVA needs non-zero variance in every group")
    w = n / var
    mw = np.sum(w * means) / np.sum(w)
    a = np.sum(w * (means - mw) ** 2) / (k - 1)
    tmp = np.sum((1 - w / np.sum(w)) ** 2 / (n - 1))
    f = a / (1 + 2 * (k - 2) / (k ** 2 - 1) * tmp)
    df1, df2 = k - 1, (k ** 2 - 1) / (3 * tmp)
    return f, df1, df2, stats.f.sf(f, df1, df2)


def _rm_anova(m):
    n, k = m.shape
    gm = m.mean()
    ss_cond = n * np.sum((m.mean(axis=0) - gm) ** 2)
    ss_subj = k * np.sum((m.mean(axis=1) - gm) ** 2)
    ss_err = np.sum((m - gm) ** 2) - ss_cond - ss_subj
    df1, df2 = k - 1, (k - 1) * (n - 1)
    if ss_err <= 0:
        raise StatsError("repeated measures ANOVA: no residual variance")
    f = (ss_cond / df1) / (ss_err / df2)
    # Greenhouse-Geisser epsilon from the double-centred covariance matrix
    s = np.cov(m, rowvar=False)
    sc = s - s.mean(axis=0, keepdims=True) - s.mean(axis=1, keepdims=True) + s.mean()
    denom = (k - 1) * np.trace(sc @ sc)
    eps = float(np.trace(sc) ** 2 / denom) if denom > 0 else 1.0
    eps = min(max(eps, 1.0 / (k - 1)), 1.0)
    return f, df1, df2, stats.f.sf(f, df1, df2), eps, stats.f.sf(f, eps * df1, eps * df2)


def _omnibus(anova_type, samples, matrix, category):
    o = OmnibusTest(category, test=ANOVA_LABELS.get(anova_type, str(anova_type)),
                    n_groups=len(samples), n_total=int(sum(len(s) for s in samples)))
    try:
        k = len(samples)
        if anova_type == "One-way ANOVA":
            r = stats.f_oneway(*samples)
            o.statistic_name, o.statistic, o.p_value = "F", float(r.statistic), float(r.pvalue)
            o.df1, o.df2 = k - 1, o.n_total - k
        elif anova_type == "Welch's ANOVA":
            o.statistic_name = "F"
            o.statistic, o.df1, o.df2, o.p_value = map(float, _welch_anova(samples))
        elif anova_type == "Kruskal-Wallis H test (non-parametric)":
            r = stats.kruskal(*samples)
            o.statistic_name, o.statistic, o.p_value, o.df1 = "H", float(r.statistic), float(r.pvalue), k - 1
        elif anova_type == "Repeated measures ANOVA":
            f, df1, df2, p, eps, p_gg = _rm_anova(matrix)
            o.statistic_name, o.statistic, o.df1, o.df2, o.p_value = "F", float(f), df1, df2, float(p)
            o.extra = {"Subjects": matrix.shape[0], "Greenhouse-Geisser epsilon": eps, "p (Greenhouse-Geisser)": float(p_gg)}
        elif anova_type == "Friedman test (non-parametric)":
            r = stats.friedmanchisquare(*[matrix[:, i] for i in range(k)])
            o.statistic_name, o.statistic, o.p_value, o.df1 = "Chi²", float(r.statistic), float(r.pvalue), k - 1
            o.extra = {"Subjects": matrix.shape[0]}
        else:
            raise StatsError(f"unknown ANOVA type '{anova_type}'")
        if not np.isfinite(o.p_value):
            raise StatsError("the test returned no p-value (e.g. all values identical)")
    except Exception as e:
        o.p_value, o.error = NAN, str(e)
    return o


# ----------------------------------------------------------------------------
# Post-hoc tests (k > 2 groups)
# ----------------------------------------------------------------------------

def _games_howell(samples):
    k = len(samples)
    out = {}
    for i, j in itertools.combinations(range(k), 2):
        a, b = samples[i], samples[j]
        va, vb = np.var(a, ddof=1) / len(a), np.var(b, ddof=1) / len(b)
        se = math.sqrt(va + vb)
        if se <= 0:
            raise StatsError("Games-Howell needs non-zero variance in the groups")
        diff = float(np.mean(a) - np.mean(b))
        df = (va + vb) ** 2 / (va ** 2 / (len(a) - 1) + vb ** 2 / (len(b) - 1))
        t = diff / se
        qcrit = stats.studentized_range.ppf(0.95, k, df)
        out[(i, j)] = dict(p_value=float(stats.studentized_range.sf(abs(t) * math.sqrt(2), k, df)),
                           statistic_name="t", statistic=float(t), df=float(df),
                           difference_name="Mean diff", difference=diff,
                           ci_low=diff - qcrit / math.sqrt(2) * se, ci_high=diff + qcrit / math.sqrt(2) * se)
    return out


def _posthoc(posthoc_type, samples, matrix):
    """Return (label, adjustment, {(i, j): Comparison field values}) for all i < j."""
    k = len(samples)
    pairs = list(itertools.combinations(range(k), 2))
    if posthoc_type == "Tukey's HSD":
        r = stats.tukey_hsd(*samples)
        ci = r.confidence_interval(0.95)
        return "Tukey's HSD", "Tukey", {
            (i, j): dict(p_value=float(r.pvalue[i, j]), difference_name="Mean diff",
                         difference=float(r.statistic[i, j]),
                         ci_low=float(ci.low[i, j]), ci_high=float(ci.high[i, j]))
            for i, j in pairs}
    if posthoc_type == "Games-Howell":
        return "Games-Howell", "Games-Howell", _games_howell(samples)

    long = pd.DataFrame({'v': np.concatenate(samples),
                         'g': np.repeat(np.arange(k), [len(s) for s in samples])})
    if posthoc_type == "Tamhane's T2":
        label, adj, P = "Tamhane's T2", "Tamhane", sp.posthoc_tamhane(long, val_col='v', group_col='g')
    elif posthoc_type == "Scheffe's test":
        label, adj, P = "Scheffé's test", "Scheffé", sp.posthoc_scheffe(long, val_col='v', group_col='g')
    elif posthoc_type == "Dunn's test":
        label, adj = "Dunn's test", "Bonferroni"
        P = sp.posthoc_dunn(long, val_col='v', group_col='g', p_adjust='bonferroni')
    elif posthoc_type == "Conover's test (non-parametric)":
        if matrix is not None:
            label, adj, P = "Conover's test for Friedman (blocked)", "Holm", sp.posthoc_conover_friedman(matrix, p_adjust='holm')
        else:
            label, adj, P = "Conover's test", "Holm", sp.posthoc_conover(long, val_col='v', group_col='g', p_adjust='holm')
    elif posthoc_type == "Nemenyi test (non-parametric)":
        if matrix is not None:
            label, adj, P = "Nemenyi test for Friedman (blocked)", "Nemenyi", sp.posthoc_nemenyi_friedman(matrix)
        else:
            label, adj, P = "Nemenyi test", "Nemenyi", sp.posthoc_nemenyi(long, val_col='v', group_col='g')
    else:
        raise StatsError(f"unknown post-hoc test '{posthoc_type}'")

    P.index, P.columns = range(k), range(k)
    nonpar = posthoc_type in NONPARAMETRIC_POSTHOC
    out = {}
    for i, j in pairs:
        if nonpar:
            diff_name, diff = "Median diff", float(np.median(samples[i]) - np.median(samples[j]))
        else:
            diff_name, diff = "Mean diff", float(np.mean(samples[i]) - np.mean(samples[j]))
        out[(i, j)] = dict(p_value=float(P.loc[i, j]), difference_name=diff_name, difference=diff)
    return label, adj, out


def _holm_sidak(p_values):
    """Holm-Šídák step-down adjusted p-values (same order as the input)."""
    p = np.asarray(p_values, dtype=float)
    m = len(p)
    adj = np.empty(m)
    running = 0.0
    for rank, idx in enumerate(np.argsort(p)):
        running = max(running, -math.expm1((m - rank) * math.log1p(-p[idx])) if p[idx] < 1 else 1.0)
        adj[idx] = min(running, 1.0)
    return adj


# ----------------------------------------------------------------------------
# Analysis of one set of groups
# ----------------------------------------------------------------------------

def _analyze(res, raw, category, s, subj=None):
    """Compare the groups in ``raw`` ({group: values incl. NaN}) and append results to ``res``.

    ``subj`` ({group: subject IDs}) is used to match values for paired / repeated-measures tests.
    """
    groups = [g for g, v in raw.items() if len(_clean(v)) > 0]
    for g in groups:
        res.descriptives.append(_describe(category, g, raw[g]))
    where = f"Category '{res.label(category)}': " if category is not None else ""
    if len(groups) < 2:
        res.notes.append(f"{where}fewer than two groups with data – no test performed.")
        return
    names = [res.label(g) for g in groups]
    ids = [subj[g] for g in groups] if subj else None
    alpha = res.alpha

    if len(groups) == 2:
        c = _two_sample(raw[groups[0]], raw[groups[1]], s['test_type'], s['alternative'],
                        category, groups[0], groups[1], names, ids)
        c.significance = pval_to_annotation(c.p_value, alpha)
        res.comparisons.append(c)
        return

    anova_type, posthoc_type = s['anova_type'], s['posthoc_type']
    matrix = None
    try:
        if anova_type in REPEATED_OMNIBUS:
            matrix, _ = _paired_matrix([raw[g] for g in groups], names, ids)
            samples = [matrix[:, i] for i in range(len(groups))]
        else:
            samples = [_clean(raw[g]) for g in groups]
        too_small = [f"{n} (n={len(x)})" for n, x in zip(names, samples) if len(x) < 2]
        if too_small:
            raise StatsError("needs at least 2 values per group: " + ", ".join(too_small))
    except StatsError as e:
        res.omnibus.append(OmnibusTest(category, ANOVA_LABELS.get(anova_type, anova_type),
                                       n_groups=len(groups), error=str(e)))
        for g1, g2 in itertools.combinations(groups, 2):
            res.comparisons.append(Comparison(category, g1, g2, test=str(posthoc_type), error=str(e)))
        return

    o = _omnibus(anova_type, samples, matrix, category)
    o.significance = pval_to_annotation(o.p_value, alpha)
    res.omnibus.append(o)

    try:
        label, adj, ph = _posthoc(posthoc_type, samples, matrix if posthoc_type in BLOCKED_POSTHOC else None)
        err = ""
    except Exception as e:
        label, adj, ph, err = str(posthoc_type), "none", {}, str(e)
    for i, j in itertools.combinations(range(len(groups)), 2):
        fields = ph.get((i, j), {})
        c = Comparison(category, groups[i], groups[j], test=label, adjustment=adj,
                       n1=len(samples[i]), n2=len(samples[j]), **fields)
        if not err and not np.isfinite(c.p_value):
            err = "post-hoc test returned no p-value"
        c.error = err if not np.isfinite(c.p_value) else ""
        c.significance = pval_to_annotation(c.p_value, alpha)
        res.comparisons.append(c)
    if np.isfinite(o.p_value) and o.p_value > alpha:
        res.notes.append(f"{where}the omnibus test is not significant (p = {format_p(o.p_value)}); "
                         "interpret the post-hoc comparisons with caution.")


def _two_way_anova(res, data, x_col, hue_col, value_col, x_levels, hue_levels):
    """Two-way ANOVA (Type III SS) + Šídák comparisons of groups within each category (pooled error)."""
    d = data[np.isfinite(data[value_col].to_numpy(dtype=float))]
    cells = {}
    for x in x_levels:
        for h in hue_levels:
            v = d.loc[(d[x_col] == x) & (d[hue_col] == h), value_col].to_numpy(dtype=float)
            cells[(x, h)] = v
            if len(v):
                res.descriptives.append(_describe(x, h, v))

    missing = [f"{res.label(x)} / {res.label(h)}" for (x, h), v in cells.items() if len(v) == 0]
    n_total = len(d)
    dfr = n_total - len(x_levels) * len(hue_levels)
    error = ""
    if missing:
        error = ("two-way ANOVA needs data in every category × group cell (empty: " + ", ".join(missing)
                 + "). Choose 'Separate tests per category' in the statistics settings.")
    elif dfr <= 0:
        error = "two-way ANOVA needs replicates (more than one value per cell)."
    if error:
        res.omnibus.append(OmnibusTest(None, "Two-way ANOVA", error=error))
        return

    import statsmodels.formula.api as smf
    from statsmodels.stats.anova import anova_lm

    xcode = {x: i for i, x in enumerate(x_levels)}
    hcode = {h: i for i, h in enumerate(hue_levels)}
    model_df = pd.DataFrame({'y': d[value_col].to_numpy(dtype=float),
                             'a': [xcode[v] for v in d[x_col]],
                             'b': [hcode[v] for v in d[hue_col]]})
    model = smf.ols('y ~ C(a, Sum) * C(b, Sum)', data=model_df).fit()
    table = anova_lm(model, typ=3)
    x_name, h_name = res.x_label or x_col, res.hue_label or hue_col
    for term, name, n_levels in [('C(a, Sum):C(b, Sum)', f"Interaction ({x_name} × {h_name})", 0),
                                 ('C(a, Sum)', f"{x_name} (categories)", len(x_levels)),
                                 ('C(b, Sum)', f"{h_name} (groups)", len(hue_levels))]:
        row = table.loc[term]
        o = OmnibusTest(None, f"Two-way ANOVA – {name}", statistic_name="F", statistic=float(row['F']),
                        df1=float(row['df']), df2=float(model.df_resid), p_value=float(row['PR(>F)']),
                        n_groups=n_levels or len(x_levels) * len(hue_levels), n_total=n_total)
        o.significance = pval_to_annotation(o.p_value, res.alpha)
        res.omnibus.append(o)

    mse, dfr = float(model.mse_resid), float(model.df_resid)
    pairs = [(x, h1, h2) for x in x_levels for h1, h2 in itertools.combinations(hue_levels, 2)]
    m = len(pairs)
    tcrit = stats.t.ppf(1 - (-math.expm1(math.log1p(-0.05) / m)) / 2, dfr)
    for x, h1, h2 in pairs:
        a, b = cells[(x, h1)], cells[(x, h2)]
        diff = float(np.mean(a) - np.mean(b))
        se = math.sqrt(mse * (1 / len(a) + 1 / len(b)))
        t = diff / se
        p_raw = float(2 * stats.t.sf(abs(t), dfr))
        c = Comparison(x, h1, h2, test="Šídák's multiple comparisons (two-way ANOVA)",
                       adjustment=f"Šídák (m={m})", p_unadjusted=p_raw,
                       p_value=float(-math.expm1(m * math.log1p(-p_raw))) if p_raw < 1 else 1.0,
                       statistic_name="t", statistic=float(t), df=dfr,
                       difference_name="Mean diff", difference=diff,
                       ci_low=diff - tcrit * se, ci_high=diff + tcrit * se, n1=len(a), n2=len(b))
        c.significance = pval_to_annotation(c.p_value, res.alpha)
        res.comparisons.append(c)

    bf = stats.levene(*[v for v in cells.values() if len(v) > 1], center='median')
    if np.isfinite(bf.pvalue) and bf.pvalue < 0.05:
        sds = [np.std(v, ddof=1) for v in cells.values() if len(v) > 1]
        res.warnings.append(f"The groups have clearly different variances (Brown-Forsythe test p = {format_p(bf.pvalue)}; "
                            f"SD ranges from {_fmt(min(sds))} to {_fmt(max(sds))}). The two-way ANOVA assumes equal "
                            "variances and its pooled error is unreliable here. Use 'Separate tests per category' "
                            "with Welch's t-test / Welch's ANOVA instead.")
    res.notes.append(f"Two-way ANOVA with Type III sums of squares (factors: '{x_name}' and '{h_name}'). "
                     "It assumes independent values, normally distributed residuals and equal variances.")
    res.notes.append(f"Groups are compared within each category using the pooled residual variance of the "
                     f"two-way ANOVA (MS = {_fmt(mse)}, df = {_fmt(dfr)}). p-values and 95% CIs are Šídák-adjusted "
                     f"for all {m} comparisons in this graph (one family). All comparisons are two-sided.")


def _add_design_notes(res, s, n_blocks):
    anova_type, posthoc_type, test_type = s['anova_type'], s['posthoc_type'], s['test_type']
    uses_two = any(c.test in TTEST_LABELS.values() for c in res.comparisons)
    uses_multi = bool(res.omnibus)
    subject = s.get('subject_col')

    if (uses_two and test_type in PAIRED_TWO_SAMPLE) or (uses_multi and anova_type in REPEATED_OMNIBUS):
        if subject:
            res.notes.append(f"Paired / repeated-measures design: values are matched by subject ID "
                             f"(column '{subject}'). Subjects without a value in every group are excluded.")
        else:
            res.notes.append("Paired / repeated-measures design: no subject column is set, so values are matched "
                             "by their order within each group (the n-th value of one group is paired with the "
                             "n-th value of the other groups). Rows with a missing value in any group are "
                             "excluded. Set a subject column in the statistics settings to match by ID.")
    if uses_multi:
        if anova_type in REPEATED_OMNIBUS and posthoc_type not in BLOCKED_POSTHOC:
            res.warnings.append(f"{posthoc_type} treats the groups as independent and ignores the pairing "
                                f"used by the {ANOVA_LABELS[anova_type]}.")
        if anova_type in NONPARAMETRIC_OMNIBUS and posthoc_type not in NONPARAMETRIC_POSTHOC:
            res.warnings.append(f"A parametric post-hoc test ({posthoc_type}) follows a non-parametric omnibus "
                                f"test ({ANOVA_LABELS[anova_type]}).")
        if anova_type not in NONPARAMETRIC_OMNIBUS and posthoc_type in NONPARAMETRIC_POSTHOC:
            res.warnings.append(f"A non-parametric post-hoc test ({posthoc_type}) follows a parametric omnibus "
                                f"test ({ANOVA_LABELS[anova_type]}).")
        res.notes.append("Post-hoc p-values are adjusted for multiple comparisons by the post-hoc method "
                         "(see 'Adjusted by').")
    if uses_two and s['alternative'] != 'two-sided':
        res.notes.append(f"One-sided test ('{s['alternative']}'): the alternative hypothesis is "
                         f"Group 1 {'<' if s['alternative'] == 'less' else '>'} Group 2.")


def _correct_across_categories(res, n_blocks):
    """Separate-tests mode: apply the selected correction across the per-category two-group tests."""
    if n_blocks < 2:
        return
    two_group = [c for c in res.comparisons if c.test in TTEST_LABELS.values() and np.isfinite(c.p_value)]
    has_multi = bool(res.omnibus)
    if res.settings['grouped_analysis'] == GROUPED_SEPARATE_HOLM and len(two_group) > 1:
        m = len(two_group)
        for c, p_adj in zip(two_group, _holm_sidak([c.p_value for c in two_group])):
            c.p_unadjusted, c.p_value = c.p_value, float(p_adj)
            c.adjustment = f"Holm-Šídák (m={m})"
            c.significance = pval_to_annotation(c.p_value, res.alpha)
        res.notes.append(f"A separate test is run within each of the {n_blocks} categories. The two-group "
                         f"p-values are Holm-Šídák adjusted across the {m} categories (one family per graph).")
        if has_multi:
            res.notes.append("Categories with more than two groups use post-hoc p-values adjusted within "
                             "that category only.")
    elif has_multi:
        res.notes.append(f"A separate ANOVA + post-hoc test is run within each of the {n_blocks} categories. "
                         "Post-hoc p-values are adjusted for the comparisons within each category, not across "
                         "categories (the Holm-Šídák option applies to two-group tests only).")
    else:
        res.notes.append(f"A separate test is run within each of the {n_blocks} categories; p-values are "
                         "not adjusted across categories.")


# ----------------------------------------------------------------------------
# Public entry point
# ----------------------------------------------------------------------------

def calculate_statistics(df, x_col, value_col, hue_col=None, settings=None,
                         x_order=None, hue_order=None, labels=None, value_label=None,
                         x_label=None, hue_label=None):
    """Run all statistics for one plot.

    Ungrouped data (no ``hue_col`` or a single group): compares the x categories.
    Two groups -> the selected two-group test; more -> selected ANOVA + post-hoc.

    Grouped data (two or more groups), depending on ``settings['grouped_analysis']``:
    the ungrouped logic applied separately within each category, Holm-Šídák
    corrected across categories (default) or uncorrected, or a two-way ANOVA +
    Šídák comparisons within each category.

    Args:
        df: data exactly as plotted.
        x_col, value_col, hue_col: column names.
        settings: dict with alpha_level, test_type, alternative, anova_type, posthoc_type,
            grouped_analysis and subject_col (column with subject IDs for paired tests).
        x_order, hue_order: optional level order (defaults to seaborn's order).
        labels: optional {raw value: display label} for x categories / groups.
        value_label, x_label, hue_label: display names of the columns.

    Returns:
        StatsResult
    """
    s = {**DEFAULT_SETTINGS, **{k: v for k, v in (settings or {}).items() if v not in (None, "")}}
    if s.get('subject_col') in (NO_SUBJECT, "None"):
        s['subject_col'] = None
    if s['grouped_analysis'] not in GROUPED_OPTIONS:
        s['grouped_analysis'] = GROUPED_SEPARATE_HOLM
    try:
        alpha = float(s['alpha_level'])
    except (TypeError, ValueError):
        alpha = 0.05
    res = StatsResult(x_col=x_col, value_col=value_col, hue_col=hue_col, alpha=alpha, settings=s,
                      value_label=value_label or str(value_col), x_label=x_label or str(x_col),
                      hue_label=hue_label or (str(hue_col) if hue_col else ""), labels=dict(labels or {}))

    if df is None or x_col not in df.columns or value_col not in df.columns:
        res.error = "The data needed for statistics is not available."
        return res
    if hue_col is not None and hue_col not in df.columns:
        hue_col = res.hue_col = None
    subject_col = s['subject_col']
    if subject_col and subject_col not in df.columns:
        res.error = f"The subject column '{subject_col}' is not in the plotted data."
        return res
    if subject_col and subject_col in (x_col, hue_col, value_col):
        res.error = f"The subject column '{subject_col}' must differ from the X, group and value columns."
        return res

    cols = [x_col, value_col] + ([hue_col] if hue_col else []) + ([subject_col] if subject_col else [])
    data = df.loc[:, list(dict.fromkeys(cols))].copy()
    data[value_col] = pd.to_numeric(data[value_col], errors='coerce')
    data = data[data[x_col].notna()]
    if hue_col:
        data = data[data[hue_col].notna()]

    x_levels = _level_order(data[x_col], x_order)
    hue_levels = _level_order(data[hue_col], hue_order) if hue_col else []

    def split(sub, col, levels):
        raw = {g: sub.loc[sub[col] == g, value_col].to_numpy(dtype=float) for g in levels}
        subj = {g: sub.loc[sub[col] == g, subject_col].to_numpy() for g in levels} if subject_col else None
        return raw, subj

    if len(hue_levels) > 1:
        res.structure = 'grouped'
        if s['grouped_analysis'] == GROUPED_TWO_WAY and len(x_levels) > 1:
            _two_way_anova(res, data, x_col, hue_col, value_col, x_levels, hue_levels)
            chosen = [t for t in (s['test_type'], s['anova_type'])
                      if t in PAIRED_TWO_SAMPLE | REPEATED_OMNIBUS or "non-parametric" in t]
            if chosen or subject_col:
                res.warnings.append("Grouped data is analysed with an ordinary two-way ANOVA, which treats all "
                                    "values as independent and is parametric. The selected "
                                    + (" / ".join(chosen) if chosen else "subject column")
                                    + " is not used for grouped data. Choose 'Separate tests per category' "
                                    "to use it (repeated-measures two-way ANOVA is not available).")
        else:
            if s['grouped_analysis'] == GROUPED_TWO_WAY:
                res.notes.append("Only one category: the groups are compared with the selected t-test / ANOVA "
                                 "instead of a two-way ANOVA.")
            for x in x_levels:
                raw, subj = split(data[data[x_col] == x], hue_col, hue_levels)
                _analyze(res, raw, x, s, subj)
            _add_design_notes(res, s, len(x_levels))
            _correct_across_categories(res, len(x_levels))
    else:
        res.structure = 'ungrouped'
        raw, subj = split(data, x_col, x_levels)
        _analyze(res, raw, None, s, subj)
        _add_design_notes(res, s, 1)

    if not res.has_results and not res.error:
        res.error = "No statistical test could be performed (need at least two groups with data)."
    return res


def export_statistics(result, path):
    """Write all result tables to an .xlsx workbook (one sheet per table) or a .csv file."""
    tables = result.to_tables()
    if str(path).lower().endswith(".csv"):
        with open(path, "w", newline="", encoding="utf-8") as f:
            for name, table in tables.items():
                f.write(f"# {name}\n")
                table.to_csv(f, index=False)
                f.write("\n")
    else:
        with pd.ExcelWriter(path) as writer:
            for name, table in tables.items():
                table.to_excel(writer, sheet_name=name[:31], index=False)
