"""
Validation tests for explot_stats.

Reference values were produced with base R 4.6.1 (script at the bottom of this
file).  Post-hoc tests that base R lacks are checked against their textbook
definitions.

Run:  python -m pytest tests/   or   python tests/test_explot_stats.py
"""

import itertools
import math
import os
import sys
import tempfile

import numpy as np
import pandas as pd
import scipy.stats as st

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import explot_stats as es  # noqa: E402

A = [9.1, 10.2, 11.5, 8.7, 12.0, 10.8, 9.9, 11.1]
B = [12.3, 13.1, 11.8, 14.2, 12.9, 13.5, 12.05, 14.8]
C = [10.5, 16.2, 9.0, 18.4, 12.25, 15.1, 11.0, 17.3]


def long_df(**groups):
    return pd.DataFrame([(g, v) for g, vals in groups.items() for v in vals], columns=["cat", "val"])


def run(df, hue=None, **settings):
    return es.calculate_statistics(df, "cat", "val", hue, settings=settings)


def close(a, b, rel=1e-6, abs_=1e-10):
    assert math.isclose(a, b, rel_tol=rel, abs_tol=abs_), f"{a} != {b}"


# ---------------------------------------------------------------------------
# Two-group tests vs R
# ---------------------------------------------------------------------------

def test_student_t_matches_r():
    c = run(long_df(A=A, B=B), test_type="Student's t-test (unpaired, equal variances)").comparisons[0]
    close(c.statistic, -4.84107781883); close(c.df, 14); close(c.p_value, 0.000261598696974)
    close(c.ci_low, -3.85111055845); close(c.ci_high, -1.48638944155)


def test_welch_t_matches_r_including_df():
    c = run(long_df(A=A, B=B), test_type="Welch's t-test (unpaired, unequal variances)").comparisons[0]
    close(c.statistic, -4.84107781883); close(c.df, 13.8834838456); close(c.p_value, 0.000267799740857)
    close(c.ci_low, -3.85204216484); close(c.ci_high, -1.48545783516)


def test_one_sided_matches_r():
    c = run(long_df(A=A, B=B), alternative="less").comparisons[0]
    close(c.p_value, 0.000133899870428)


def test_paired_t_matches_r():
    c = run(long_df(A=A, B=B), test_type="Paired t-test").comparisons[0]
    close(c.statistic, -4.65266560956); close(c.df, 7); close(c.p_value, 0.00233454424476)
    close(c.ci_low, -4.02508881757); close(c.ci_high, -1.31241118243)


def test_mann_whitney_matches_r():
    c = run(long_df(A=A, B=B), test_type="Mann-Whitney U test (non-parametric)").comparisons[0]
    close(c.statistic, 1); close(c.p_value, 0.0003108003108)
    c = run(long_df(A=A, C=C), test_type="Mann-Whitney U test (non-parametric)").comparisons[0]
    close(c.statistic, 14); close(c.p_value, 0.0649572649573)


def test_wilcoxon_matches_r():
    c = run(long_df(A=A, C=C), test_type="Wilcoxon signed-rank test (non-parametric)").comparisons[0]
    close(c.p_value, 0.0546875)


# ---------------------------------------------------------------------------
# Omnibus tests vs R
# ---------------------------------------------------------------------------

def test_one_way_anova_matches_r():
    o = run(long_df(A=A, B=B, C=C), anova_type="One-way ANOVA", posthoc_type="Tukey's HSD").omnibus[0]
    close(o.statistic, 5.0675226873); assert (o.df1, o.df2) == (2, 21); close(o.p_value, 0.0160023149397)


def test_welch_anova_matches_r():
    o = run(long_df(A=A, B=B, C=C), anova_type="Welch's ANOVA").omnibus[0]
    close(o.statistic, 12.1580380805); close(o.df2, 12.7865579655); close(o.p_value, 0.00110190079743)


def test_kruskal_matches_r():
    o = run(long_df(A=A, B=B, C=C), anova_type="Kruskal-Wallis H test (non-parametric)",
            posthoc_type="Dunn's test").omnibus[0]
    close(o.statistic, 9.105); close(o.p_value, 0.0105408193679)


def test_friedman_matches_r():
    o = run(long_df(A=A, B=B, C=C), anova_type="Friedman test (non-parametric)",
            posthoc_type="Nemenyi test (non-parametric)").omnibus[0]
    close(o.statistic, 9.25); close(o.p_value, 0.00980365503582)


def test_rm_anova_matches_r():
    o = run(long_df(A=A, B=B, C=C), anova_type="Repeated measures ANOVA").omnibus[0]
    close(o.statistic, 5.97924627203); assert (o.df1, o.df2) == (2, 14); close(o.p_value, 0.0132721099089)


def test_rm_anova_greenhouse_geisser_matches_pingouin():
    try:
        import pingouin as pg
    except ImportError:
        return
    df = pd.DataFrame({"val": A + B + C, "cat": np.repeat(list("ABC"), 8), "s": np.tile(range(8), 3)})
    ref = pg.rm_anova(df, dv="val", within="cat", subject="s", correction=True)
    o = run(long_df(A=A, B=B, C=C), anova_type="Repeated measures ANOVA").omnibus[0]
    close(o.extra["Greenhouse-Geisser epsilon"], float(ref["eps"].iloc[0]), rel=1e-6)
    gg_col = [c for c in ref.columns if "GG" in c and c.startswith("p")][0]
    close(o.extra["p (Greenhouse-Geisser)"], float(ref[gg_col].iloc[0]), rel=1e-6)


# ---------------------------------------------------------------------------
# Post-hoc tests
# ---------------------------------------------------------------------------

def test_tukey_matches_r():
    res = run(long_df(A=A, B=B, C=C), anova_type="One-way ANOVA", posthoc_type="Tukey's HSD")
    # R reports B-A, C-A, C-B; ExPlot reports Group1 - Group2 (A-B, A-C, B-C)
    ref = {("A", "B"): (2.66875, -0.108757601906, 5.44625760191, 0.0611099580619),
           ("A", "C"): (3.30625, 0.528742398094, 6.08375760191, 0.0179363313579),
           ("B", "C"): (0.6375, -2.14000760191, 3.41500760191, 0.8329462055)}
    for (g1, g2), (diff, lwr, upr, p) in ref.items():
        c = res.get_comparison(g1, g2)
        # studentized-range quantiles differ between R and scipy at ~1e-6
        close(c.difference, -diff); close(c.ci_low, -upr, rel=1e-4); close(c.ci_high, -lwr, rel=1e-4)
        close(c.p_value, p, rel=1e-4)


def test_tamhane_is_sidak_adjusted_welch():
    res = run(long_df(A=A, B=B, C=C), anova_type="Welch's ANOVA", posthoc_type="Tamhane's T2")
    m = 3
    for (g1, a), (g2, b) in itertools.combinations({"A": A, "B": B, "C": C}.items(), 2):
        p = st.ttest_ind(a, b, equal_var=False).pvalue
        close(res.get_p(g1, g2), min(1.0, 1 - (1 - p) ** m), rel=1e-6)


def test_scheffe_matches_definition():
    groups = {"A": A, "B": B, "C": C}
    res = run(long_df(**groups), anova_type="One-way ANOVA", posthoc_type="Scheffe's test")
    k, n_tot = 3, 24
    mse = sum(np.sum((np.array(v) - np.mean(v)) ** 2) for v in groups.values()) / (n_tot - k)
    for (g1, a), (g2, b) in itertools.combinations(groups.items(), 2):
        f = (np.mean(a) - np.mean(b)) ** 2 / (mse * (1 / len(a) + 1 / len(b)) * (k - 1))
        close(res.get_p(g1, g2), st.f.sf(f, k - 1, n_tot - k), rel=1e-6)


def test_dunn_is_bonferroni_adjusted():
    groups = {"A": A, "B": B, "C": C}
    res = run(long_df(**groups), anova_type="Kruskal-Wallis H test (non-parametric)", posthoc_type="Dunn's test")
    v = np.concatenate(list(groups.values()))
    ranks = st.rankdata(v)
    n_tot = len(v)
    idx = {g: slice(i * 8, (i + 1) * 8) for i, g in enumerate(groups)}
    for g1, g2 in itertools.combinations(groups, 2):
        z = (ranks[idx[g1]].mean() - ranks[idx[g2]].mean()) / math.sqrt(n_tot * (n_tot + 1) / 12 * (2 / 8))
        close(res.get_p(g1, g2), min(1.0, 2 * st.norm.sf(abs(z)) * 3), rel=1e-6)
        assert res.get_comparison(g1, g2).adjustment == "Bonferroni"


def test_blocked_posthoc_used_after_friedman():
    res = run(long_df(A=A, B=B, C=C), anova_type="Friedman test (non-parametric)",
              posthoc_type="Nemenyi test (non-parametric)")
    assert all("Friedman" in c.test for c in res.comparisons)
    res = run(long_df(A=A, B=B, C=C), anova_type="Kruskal-Wallis H test (non-parametric)",
              posthoc_type="Nemenyi test (non-parametric)")
    assert all("Friedman" not in c.test for c in res.comparisons)


# ---------------------------------------------------------------------------
# Consistency, errors and ordering
# ---------------------------------------------------------------------------

def test_prism_significance_levels():
    f = es.pval_to_annotation
    assert [f(p) for p in (0.0001, 0.00011, 0.001, 0.0011, 0.01, 0.011, 0.05, 0.0501)] == \
        ["****", "***", "***", "**", "**", "*", "*", "ns"]
    assert f(float("nan")) == "n/a"


def test_annotation_symbol_equals_details_symbol():
    res = run(long_df(A=A, B=B, C=C), anova_type="One-way ANOVA", posthoc_type="Tukey's HSD")
    text = res.report()
    for c in res.comparisons:
        assert c.significance == es.pval_to_annotation(c.p_value, res.alpha)
        assert es.format_p(c.p_value) in text
        assert res.get_p(c.group2, c.group1) == c.p_value  # order-independent lookup
    table = res.to_tables()["Pairwise comparisons"]
    assert list(table["Summary (on graph)"]) == [c.significance for c in res.comparisons]


def test_no_silent_fallback_for_unpaired_sizes():
    res = run(long_df(A=A, B=B[:6]), test_type="Paired t-test")
    c = res.comparisons[0]
    assert math.isnan(c.p_value) and "same number" in c.error and c.significance == "n/a"
    assert c.test == "Paired t-test"


def test_pairing_drops_incomplete_pairs_only():
    b = list(B)
    b[2] = float("nan")
    c = run(long_df(A=A, B=b), test_type="Paired t-test").comparisons[0]
    keep = [i for i in range(8) if i != 2]
    ref = st.ttest_rel([A[i] for i in keep], [B[i] for i in keep])
    assert (c.n1, c.n2) == (7, 7)
    close(c.p_value, ref.pvalue)


def test_unknown_test_is_an_error_not_a_fallback():
    c = run(long_df(A=A, B=B), test_type="Something else").comparisons[0]
    assert math.isnan(c.p_value) and c.error


def test_grouped_data_per_category_in_hue_order():
    df = pd.DataFrame({
        "cat": ["X"] * 16 + ["Y"] * 16,
        "grp": (["Treated"] * 8 + ["Control"] * 8) * 2,
        "val": A + B + B + C,
    })
    res = es.calculate_statistics(df, "cat", "val", "grp", settings={"grouped_analysis": es.GROUPED_SEPARATE_RAW})
    assert res.structure == "grouped"
    assert [(c.category, c.group1, c.group2) for c in res.comparisons] == \
        [("X", "Treated", "Control"), ("Y", "Treated", "Control")]
    close(res.get_p("Control", "Treated", "X"), st.ttest_ind(A, B, equal_var=False).pvalue)
    assert res.get_p("Control", "Treated", "Z") is None


# ---------------------------------------------------------------------------
# Grouped data: two-way ANOVA, correction across categories
# ---------------------------------------------------------------------------

TWO_WAY = pd.DataFrame({
    "cat": ["X"] * 9 + ["Y"] * 10 + ["Z"] * 8,
    "grp": ["Ctrl"] * 4 + ["Trt"] * 5 + ["Ctrl"] * 5 + ["Trt"] * 5 + ["Ctrl"] * 4 + ["Trt"] * 4,
    "val": [5.1, 6.2, 5.8, 4.9, 7.4, 8.1, 6.9, 7.7, 8.4, 6.0, 5.5, 6.8, 7.1, 6.3, 9.2, 10.1, 8.7, 9.9, 10.4,
            4.2, 5.0, 4.8, 5.5, 5.1, 4.6, 5.9, 5.3],
})


TW = {"grouped_analysis": es.GROUPED_TWO_WAY}


def test_default_grouped_is_welch_per_category_holm_sidak():
    from statsmodels.stats.multitest import multipletests
    res = es.calculate_statistics(TWO_WAY, "cat", "val", "grp")
    raw = [st.ttest_ind(TWO_WAY.val[(TWO_WAY.cat == c) & (TWO_WAY.grp == "Ctrl")],
                        TWO_WAY.val[(TWO_WAY.cat == c) & (TWO_WAY.grp == "Trt")], equal_var=False).pvalue
           for c in "XYZ"]
    assert all(c.test.startswith("Welch") for c in res.comparisons) and not res.omnibus
    for c, p, padj in zip(res.comparisons, raw, multipletests(raw, method="holm-sidak")[1]):
        close(c.p_unadjusted, p); close(c.p_value, padj)


def test_two_way_warns_on_unequal_variances():
    df = TWO_WAY.copy()
    trt = df.grp == "Trt"
    df.loc[trt, "val"] = df.loc[trt, "val"] * 40
    res = es.calculate_statistics(df, "cat", "val", "grp", settings=TW)
    assert any("Brown-Forsythe" in w for w in res.warnings)
    assert not any("Brown-Forsythe" in w for w in es.calculate_statistics(TWO_WAY, "cat", "val", "grp", settings=TW).warnings)


def test_two_way_anova_type3_matches_r():
    res = es.calculate_statistics(TWO_WAY, "cat", "val", "grp", settings=TW)
    inter, cat, grp = res.omnibus
    close(inter.statistic, 13.3438965573); close(inter.p_value, 0.000181981288304); assert inter.df1 == 2
    close(cat.statistic, 52.4049916972); close(cat.p_value, 6.85952418896e-09)
    close(grp.statistic, 69.0805802581); close(grp.p_value, 4.44015042147e-08)
    assert inter.df2 == 21


def test_two_way_sidak_comparisons_match_r():
    res = es.calculate_statistics(TWO_WAY, "cat", "val", "grp", settings=TW)
    ref = {"X": (-5.39539438173, 2.37088416502e-05, 7.11248386366e-05, -3.25754754137, -1.14245245863),
           "Y": (-8.63604426064, 2.36453518914e-08, 7.09360540307e-08, -4.3170653839, -2.3229346161),
           "Z": (-0.814310085328, 0.424605676346, 0.809499236942, -1.46475298821, 0.764752988211)}
    for cat, (t, p, padj, lo, hi) in ref.items():
        c = res.get_comparison("Ctrl", "Trt", cat)
        close(c.statistic, t); close(c.p_unadjusted, p); close(c.p_value, padj)
        close(c.ci_low, lo); close(c.ci_high, hi); assert c.df == 21
        assert c.significance == es.pval_to_annotation(padj)


def test_two_way_requires_every_cell():
    df = TWO_WAY[~((TWO_WAY.cat == "Z") & (TWO_WAY.grp == "Trt"))]
    res = es.calculate_statistics(df, "cat", "val", "grp", settings=TW)
    assert not res.comparisons and "every category" in res.omnibus[0].error


def test_two_way_warns_when_paired_test_selected():
    res = es.calculate_statistics(TWO_WAY, "cat", "val", "grp", settings={**TW, "test_type": "Paired t-test"})
    assert any("Paired t-test" in w for w in res.warnings)


def test_holm_sidak_across_categories_matches_statsmodels():
    from statsmodels.stats.multitest import multipletests
    res = es.calculate_statistics(TWO_WAY, "cat", "val", "grp",
                                  settings={"grouped_analysis": es.GROUPED_SEPARATE_HOLM})
    raw = [c.p_unadjusted for c in res.comparisons]
    ref = multipletests(raw, method="holm-sidak")[1]
    for c, r in zip(res.comparisons, ref):
        close(c.p_value, r)
        assert c.adjustment.startswith("Holm-Šídák")


# ---------------------------------------------------------------------------
# Games-Howell and subject matching
# ---------------------------------------------------------------------------

def test_games_howell_matches_pingouin():
    try:
        import pingouin as pg
    except ImportError:
        return
    df = long_df(A=A, B=B, C=C)
    ref = pg.pairwise_gameshowell(data=df, dv="val", between="cat")
    res = run(df, anova_type="Welch's ANOVA", posthoc_type="Games-Howell")
    pcol = "pval" if "pval" in ref.columns else "p-val"
    for _, row in ref.iterrows():
        c = res.get_comparison(row["A"], row["B"])
        close(c.p_value, float(row[pcol]), rel=1e-4)
        close(abs(c.statistic), abs(float(row["T"])), rel=1e-6)
        close(c.df, float(row["df"]), rel=1e-6)


def test_subject_column_matches_by_id_not_row_order():
    ids = list(range(8))
    df = pd.DataFrame({"cat": ["A"] * 8 + ["B"] * 8, "val": A + B, "subj": ids + ids})
    shuffled = pd.concat([df.iloc[:8], df.iloc[8:].sample(frac=1, random_state=1)])
    ref = st.ttest_rel(A, B).pvalue
    c = es.calculate_statistics(shuffled, "cat", "val", settings={"test_type": "Paired t-test",
                                                                   "subject_col": "subj"}).comparisons[0]
    close(c.p_value, ref)
    c = es.calculate_statistics(shuffled, "cat", "val", settings={"test_type": "Paired t-test"}).comparisons[0]
    assert not math.isclose(c.p_value, ref)  # row order gives a different (wrong) pairing


def test_subject_column_rm_anova_and_errors():
    ids = list(range(8))
    df = pd.DataFrame({"cat": list("A" * 8 + "B" * 8 + "C" * 8), "val": A + B + C, "subj": ids * 3})
    df = df.sample(frac=1, random_state=2)
    o = es.calculate_statistics(df, "cat", "val", settings={"anova_type": "Repeated measures ANOVA",
                                                            "subject_col": "subj"}).omnibus[0]
    close(o.statistic, 5.97924627203)
    df2 = pd.DataFrame({"cat": ["A"] * 3 + ["B"] * 3, "val": [1, 2, 3, 2, 3, 5.0], "subj": [1, 1, 2, 1, 2, 3]})
    c = es.calculate_statistics(df2, "cat", "val", settings={"test_type": "Paired t-test",
                                                             "subject_col": "subj"}).comparisons[0]
    assert "more than one value" in c.error
    res = es.calculate_statistics(df2, "cat", "val", settings={"subject_col": "missing"})
    assert "not in the plotted data" in res.error


def test_labels_and_x_order():
    df = long_df(**{"0": A, "1": B}).assign(cat=lambda d: d["cat"].astype(int))
    res = es.calculate_statistics(df, "cat", "val", labels={0: "Ctrl", 1: "Drug"})
    assert "Ctrl vs Drug" in res.report()


def test_export_roundtrip():
    res = run(long_df(A=A, B=B, C=C), anova_type="One-way ANOVA", posthoc_type="Tukey's HSD")
    with tempfile.TemporaryDirectory() as d:
        es.export_statistics(res, os.path.join(d, "s.xlsx"))
        sheets = pd.read_excel(os.path.join(d, "s.xlsx"), sheet_name=None)
        assert set(sheets) == set(res.to_tables())
        es.export_statistics(res, os.path.join(d, "s.csv"))
        assert "Pairwise comparisons" in open(os.path.join(d, "s.csv"), encoding="utf-8").read()


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"PASS {name}")
        except Exception as e:  # noqa: BLE001
            failed += 1
            print(f"FAIL {name}: {e!r}")
    print(f"\n{len(tests) - failed}/{len(tests)} passed")
    sys.exit(1 if failed else 0)


# R reference script (base R 4.6.1):
#   A <- c(9.1, 10.2, 11.5, 8.7, 12.0, 10.8, 9.9, 11.1)
#   B <- c(12.3, 13.1, 11.8, 14.2, 12.9, 13.5, 12.05, 14.8)
#   C <- c(10.5, 16.2, 9.0, 18.4, 12.25, 15.1, 11.0, 17.3)
#   t.test(A, B, var.equal=TRUE); t.test(A, B); t.test(A, B, alternative="less")
#   t.test(A, B, paired=TRUE); wilcox.test(A, B); wilcox.test(A, C)
#   wilcox.test(A, C, paired=TRUE)
#   v <- c(A, B, C); g <- factor(rep(c("A","B","C"), each=8)); s <- factor(rep(1:8, 3))
#   summary(aov(v ~ g)); oneway.test(v ~ g, var.equal=FALSE); TukeyHSD(aov(v ~ g))
#   kruskal.test(v ~ g); friedman.test(cbind(A, B, C)); summary(aov(v ~ g + Error(s/g)))
