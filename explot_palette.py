"""
Palette and color helper module for ExPlot.

Provides default color/palette definitions, hex conversion utilities,
JSON persistence for custom colors/palettes, and palette resolution logic.
"""

import os
import json
import seaborn as sns


# ── Default single colors ──────────────────────────────────────────────────

DEFAULT_COLORS = {
    "Black": "#000000",
    "Red": "#E74C3C",
    "Blue": "#3498DB",
    "Blueish": "#2B2FFF",
    "Green": "#27AE60",
    "Orange": "#E67E22",
    "Purple": "#8E44AD",
    "Gray": "#7F8C8D",
    "Cyan": "#00838F"
}


# ── Default palettes ───────────────────────────────────────────────────────

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


# ── Utility functions ──────────────────────────────────────────────────────

def to_hex(color):
    """Convert a color value to hex string.

    Accepts hex strings, RGB tuples/lists (0-255 ints or 0-1 floats).
    """
    if isinstance(color, str) and color.startswith('#'):
        return color
    if isinstance(color, (tuple, list)) and len(color) == 3:
        # If already 0-1 floats, convert to 0-255
        if all(isinstance(x, float) and 0 <= x <= 1 for x in color):
            color = tuple(int(x * 255) for x in color)
        return '#{:02x}{:02x}{:02x}'.format(*color)
    return str(color)


def load_colors_palettes(colors_file, palettes_file):
    """Load custom colors and palettes from JSON files.

    Returns (custom_colors: dict, custom_palettes: dict).
    Falls back to defaults when files are missing.
    """
    if os.path.exists(colors_file):
        with open(colors_file, 'r') as f:
            custom_colors = json.load(f)
    else:
        custom_colors = DEFAULT_COLORS.copy()

    if os.path.exists(palettes_file):
        with open(palettes_file, 'r') as f:
            loaded = json.load(f)
            custom_palettes = {k: [to_hex(c) for c in v] for k, v in loaded.items()}
    else:
        custom_palettes = {k: [to_hex(c) for c in v] for k, v in DEFAULT_PALETTES.items()}

    return custom_colors, custom_palettes


def save_colors_palettes(custom_colors, custom_palettes, colors_file, palettes_file):
    """Persist custom colors and palettes to JSON files."""
    with open(colors_file, 'w') as f:
        json.dump(custom_colors, f, indent=2)
    # Always save palettes as hex
    palettes_hex = {k: [to_hex(c) for c in v] for k, v in custom_palettes.items()}
    with open(palettes_file, 'w') as f:
        json.dump(palettes_hex, f, indent=2)


def resolve_palette(custom_palettes, palette_name, n_colors):
    """Resolve a palette name to a list of *n_colors* hex values.

    If the stored palette has fewer colors than needed it is repeated
    (tiled) to reach the required length.
    """
    palette = custom_palettes.get(palette_name, None)
    if palette is None:
        # Fallback to Set2
        palette = DEFAULT_PALETTES.get("Set2", ["#333333"])
    if isinstance(palette, list):
        if len(palette) < n_colors:
            palette = (palette * ((n_colors // len(palette)) + 1))[:n_colors]
        else:
            palette = palette[:n_colors]
    else:
        palette = [palette] * n_colors
    return palette


# ── Default marker sequence for XY plots ──────────────────────────────────

DEFAULT_MARKER_SEQUENCE = ["o", "s", "^", "D", "v", "P", "X", "p", "*", "+", "x"]


def resolve_markers(mode, single_marker, n_groups):
    """Return a list of *n_groups* marker symbols.

    Parameters
    ----------
    mode : str
        'Single' – every group uses *single_marker*.
        'Varied' – cycle through DEFAULT_MARKER_SEQUENCE.
    single_marker : str
        The marker used when *mode* is 'Single'.
    n_groups : int
        Number of groups/categories to assign markers to.
    """
    if mode == "Varied" and n_groups > 1:
        seq = DEFAULT_MARKER_SEQUENCE
        return [(seq[i % len(seq)]) for i in range(n_groups)]
    return [single_marker] * n_groups


def resolve_single_color(custom_colors, color_name):
    """Resolve a named single color to its hex value."""
    return custom_colors.get(color_name, '#000000')


def get_outline_color(setting, color=None):
    """Determine outline color based on a setting string.

    Parameters
    ----------
    setting : str
        One of 'as_set', 'black', 'gray', 'white'.
    color : str or None
        The palette/bar color; used when *setting* is 'as_set'.
    """
    if setting == "as_set" and color:
        return color
    elif setting == "black":
        return "black"
    elif setting == "gray":
        return "gray"
    elif setting == "white":
        return "white"
    else:
        return "black"
