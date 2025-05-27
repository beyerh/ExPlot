"""
Modern theming for ExPlot using ttkbootstrap.

This module provides a modern, cross-platform theming solution using ttkbootstrap.
It supports multiple themes, automatic dark/light mode detection, and consistent
styling across different platforms.
"""
import os
import sys
import platform
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.style import Style
from ttkbootstrap.constants import *
from ttkbootstrap.themes.standard import STANDARD_THEMES

def _get_nord_colors(variant='nord'):
    """Return color palette for different Nord theme variants.
    
    Args:
        variant: 'nord' (darker) or 'nordic' (lighter dark theme)
    """
    # Official Nord color palette: https://www.nordtheme.com/docs/colors-and-palettes
    nord_colors = {
        # Polar Night (dark background colors)
        'nord0': '#2E3440',  # nord0 - Polar Night (darkest)
        'nord1': '#3B4252',  # nord1 - Polar Night (darker)
        'nord2': '#434C5E',  # nord2 - Polar Night (lighter)
        'nord3': '#4C566A',  # nord3 - Polar Night (lightest)
        
        # Snow Storm (light text colors)
        'nord4': '#D8DEE9',  # nord4 - Snow Storm (darkest)
        'nord5': '#E5E9F0',  # nord5 - Snow Storm (lighter)
        'nord6': '#ECEFF4',  # nord6 - Snow Storm (lightest)
        
        # Frost (accent colors)
        'nord7': '#8FBCBB',  # nord7 - Frost (turquoise)
        'nord8': '#88C0D0',  # nord8 - Frost (light blue)
        'nord9': '#81A1C1',  # nord9 - Frost (blue)
        'nord10': '#5E81AC',  # nord10 - Frost (dark blue)
        
        # Aurora (accent colors)
        'nord11': '#BF616A',  # nord11 - Aurora (red)
        'nord12': '#D08770',  # nord12 - Aurora (orange)
        'nord13': '#EBCB8B',  # nord13 - Aurora (yellow)
        'nord14': '#A3BE8C',  # nord14 - Aurora (green)
        'nord15': '#B48EAD',  # nord15 - Aurora (purple)
    }
    
    if variant.lower() == 'nordic':
        # Nordic variant - lighter dark theme
        return {
            **nord_colors,
            'nord0': '#3B4252',  # Lighter than standard Nord
            'nord1': '#434C5E',
            'nord2': '#4C566A',
            'nord3': '#4C5A6E',  # Slightly lighter than standard
        }
    else:  # Default Nord (standard dark theme)
        return nord_colors

def _setup_nord_theme(style, variant='nord'):
    """Configure the Nord theme colors and styles.
    
    Args:
        style: The ttkbootstrap Style object
        variant: 'nord' (darker) or 'nordic' (lighter dark theme)
    """
    # Get the Nord color palette
    nord_colors = _get_nord_colors(variant)
    is_nordic = variant.lower() == 'nordic'
    
    # Theme name should be lowercase for consistency
    theme_name = variant.lower()
    
    # Check if theme already exists
    available_themes = [t.lower() for t in style.theme_names()]
    
    if theme_name not in available_themes:
        # Create the theme based on darkly (both are dark themes)
        style.theme_create(theme_name, 'darkly')
    
    # Use the theme
    style.theme_use(theme_name)
    
    # Common color mappings for both themes
    style.colors.primary = nord_colors['nord8']  # Light blue
    style.colors.secondary = nord_colors['nord7']  # Turquoise
    style.colors.success = nord_colors['nord14']  # Green
    style.colors.info = nord_colors['nord9']  # Blue
    style.colors.warning = nord_colors['nord13']  # Yellow
    style.colors.danger = nord_colors['nord11']  # Red
    style.colors.light = nord_colors['nord6']  # Light text
    style.colors.dark = nord_colors['nord0']  # Dark background
    
    # Background and foreground colors
    style.colors.bg = nord_colors['nord0']  # Dark background
    style.colors.fg = nord_colors['nord4']  # Light text
    style.colors.selectbg = nord_colors['nord10']  # Dark blue for selection
    style.colors.selectfg = nord_colors['nord6']  # Light text for selection
    style.colors.border = nord_colors['nord3']  # Lighter border
    style.colors.inputfg = nord_colors['nord4']  # Light text for input
    style.colors.inputbg = nord_colors['nord1']  # Darker background for input
    style.colors.active = nord_colors['nord10']  # Active element color
    
    # Configure base styles for all widgets
    style.configure('.',
        background=style.colors.bg,
        foreground=style.colors.fg,
        troughcolor=style.colors.border,
        selectbackground=style.colors.selectbg,
        selectforeground=style.colors.selectfg,
        fieldbackground=style.colors.inputbg,
        insertcolor=style.colors.fg,
        highlightthickness=0,
        highlightbackground=style.colors.bg,
        highlightcolor=style.colors.primary,
        insertwidth=1,
        borderwidth=1,
        focuscolor=style.colors.primary,
        focuswidth=1,
        relief='flat',
    )
    
    # Explicitly set background for all ttk widgets
    style.configure('TFrame', background=style.colors.bg)
    style.configure('TLabelframe', background=style.colors.bg)
    style.configure('TLabelframe.Label', background=style.colors.bg, foreground=style.colors.fg)
    style.configure('TNotebook', background=style.colors.bg, borderwidth=0)
    style.configure('TNotebook.Tab', 
        background=style.colors.border, 
        foreground=style.colors.fg,
        padding=[10, 2],
        borderwidth=0
    )
    style.map('TNotebook.Tab',
        background=[('selected', style.colors.primary)],
        foreground=[('selected', style.colors.selectfg)],
    )
    style.configure('TLabel', background=style.colors.bg, foreground=style.colors.fg)
    style.configure('TButton', 
        background=style.colors.bg, 
        foreground=style.colors.fg,
        relief='flat',
        borderwidth=0,
    )
    style.map('TButton',
        background=[('active', style.colors.primary), ('!disabled', style.colors.bg)],
        foreground=[('active', style.colors.selectfg), ('!disabled', style.colors.fg)],
    )
    
    # Configure button styles
    style.configure('TButton',
        padding=6,
        relief='flat',
        font=('Segoe UI', 10) if platform.system() == 'Windows' else 
            ('SF Pro Text', 12) if platform.system() == 'Darwin' else 
            ('DejaVu Sans', 10),
        background=style.colors.primary,
        foreground=style.colors.selectfg,
        borderwidth=1,
        focusthickness=3,
        focuscolor=style.colors.selectbg
    )
    
    # Button states
    style.map('TButton',
        background=[
            ('active', style.colors.selectbg),
            ('!disabled', style.colors.primary),
            ('pressed', style.colors.selectbg)
        ],
        foreground=[
            ('active', style.colors.selectfg),
            ('!disabled', style.colors.selectfg),
            ('pressed', style.colors.selectfg)
        ],
        relief=[
            ('pressed', 'sunken'),
            ('!pressed', 'flat')
        ]
    )
    
    # Configure checkbutton and radiobutton labels to match theme colors
    style.configure('TCheckbutton',
        background=style.colors.bg,
        foreground=style.colors.fg,
        font=('Segoe UI', 10) if platform.system() == 'Windows' else 
            ('SF Pro Text', 12) if platform.system() == 'Darwin' else 
            ('DejaVu Sans', 10)
    )
    
    style.configure('TRadiobutton',
        background=style.colors.bg,
        foreground=style.colors.fg,
        font=('Segoe UI', 10) if platform.system() == 'Windows' else 
            ('SF Pro Text', 12) if platform.system() == 'Darwin' else 
            ('DejaVu Sans', 10)
    )
    
    # Map the active states for labels only
    style.map('TCheckbutton',
        background=[('active', style.colors.bg), ('!disabled', style.colors.bg)],
        foreground=[('active', style.colors.fg), ('!disabled', style.colors.fg)]
    )
    
    style.map('TRadiobutton',
        background=[('active', style.colors.bg), ('!disabled', style.colors.bg)],
        foreground=[('active', style.colors.fg), ('!disabled', style.colors.fg)]
    )
    
    # Configure notebook tabs
    style.configure('TNotebook',
        background=style.colors.bg,
        borderwidth=0,
    )
    style.configure('TNotebook.Tab',
        padding=[12, 4],
        background=style.colors.border,
        foreground=style.colors.fg,
        borderwidth=0,
        font=('Segoe UI', 9) if platform.system() == 'Windows' else 
            ('SF Pro Text', 11) if platform.system() == 'Darwin' else 
            ('DejaVu Sans', 9),
    )
    style.map('TNotebook.Tab',
        background=[('selected', style.colors.primary)],
        foreground=[('selected', style.colors.selectfg)],
    )
    
    # Configure treeview
    style.configure('Treeview',
        fieldbackground=style.colors.bg,
        background=style.colors.border,
        foreground=style.colors.fg,
        rowheight=25,
        borderwidth=0,
        relief='flat',
    )
    style.configure('Treeview.Heading',
        background=style.colors.bg,
        foreground=style.colors.fg,
        relief='flat',
        padding=4,
    )
    style.map('Treeview',
        background=[('selected', style.colors.primary)],
        foreground=[('selected', style.colors.selectfg)],
    )
    
    # Skip scrollbar configuration - let ttkbootstrap handle it
    # The scrollbars will use the default ttkbootstrap styling
    # which should match the overall theme
    
    # Configure entry and combobox
    style.configure('TEntry',
        fieldbackground=style.colors.border,
        foreground=style.colors.fg,
        insertcolor=style.colors.fg,
        borderwidth=1,
        relief='solid',
        padding=4,
    )
    style.configure('TCombobox',
        fieldbackground=style.colors.border,
        background=style.colors.border,
        foreground=style.colors.fg,
        arrowcolor=style.colors.fg,
        arrowsize=12,
        padding=4,
    )
    
    return style

def _setup_nordic_theme(style):
    """Configure the Nordic theme (lighter version of Nord)."""
    return _setup_nord_theme(style, variant='nordic')

def setup_theme(root, dark_mode=None, theme_name=None):
    """
    Set up ttkbootstrap theme for the application.
    
    Args:
        root: The root Tk instance
        dark_mode: Optional bool to force dark/light mode. If None, auto-detect.
        theme_name: Optional theme name to use. If None, use default based on dark_mode.
                 
    Returns:
        ttkbootstrap.Style: The configured style object
    """
    # Available themes in ttkbootstrap
    LIGHT_THEMES = [
        'cosmo', 'flatly', 'journal', 'litera', 'lumen', 
        'minty', 'pulse', 'sandstone', 'united', 'yeti',
        'morph', 'simplex', 'cerculean'
    ]
    
    DARK_THEMES = [
        'solar', 'superhero', 'darkly', 'cyborg', 'vapor',
        'sharish', 'hacker', 'nord', 'nordic'
    ]
    
    # Default theme selection
    DEFAULT_LIGHT_THEME = 'cosmo'
    DEFAULT_DARK_THEME = 'nord'  # Changed default dark theme to nord
    
    # Determine if we should use dark mode if not explicitly set
    if dark_mode is None:
        dark_mode = _is_dark_mode()
    
    # Select theme based on dark mode preference if not specified
    if theme_name is None:
        if dark_mode is None:
            dark_mode = _is_dark_mode()
        theme_name = DEFAULT_DARK_THEME if dark_mode else DEFAULT_LIGHT_THEME
    
    # Normalize theme name for custom themes
    theme_lower = theme_name.lower()
    
    try:
        # Create style object with the original theme name first
        style = Style(theme=theme_name)
        
        # Apply custom themes if needed (using lowercase for consistency)
        if theme_lower == 'nord':
            style = _setup_nord_theme(style, 'nord')
        elif theme_lower == 'nordic':
            style = _setup_nord_theme(style, 'nordic')
        
        # Determine dark mode based on theme
        dark_mode = theme_lower in [t.lower() for t in DARK_THEMES]
        
        # Configure common widget styles
        _configure_widget_styles(style, dark_mode)
        
        # Set window background color
        bg = style.colors.bg if hasattr(style, 'colors') and hasattr(style.colors, 'bg') else '#f0f0f0'
        root.configure(bg=bg)
        
        # Apply custom styling for better cross-platform appearance
        _apply_cross_platform_fixes(style, dark_mode)
        
        return style
        
    except Exception as e:
        print(f"Error setting up theme '{theme_name}': {e}")
        print("Falling back to default theme...")
        # Fall back to default theme if there's an error
        if theme_name not in [DEFAULT_LIGHT_THEME, DEFAULT_DARK_THEME]:
            return setup_theme(root, dark_mode, DEFAULT_DARK_THEME if dark_mode else DEFAULT_LIGHT_THEME)
        raise

def _is_dark_mode():
    """Detect if the system is in dark mode."""
    system = platform.system().lower()
    
    if system == 'darwin':  # macOS
        try:
            import subprocess
            cmd = 'defaults read -g AppleInterfaceStyle'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            return result.returncode == 0 and result.stdout.strip().lower() == 'dark'
        except:
            return False
            
    elif system == 'windows':
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                              r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize') as key:
                value = winreg.QueryValueEx(key, 'AppsUseLightTheme')
                return value[0] == 0
        except:
            return False
            
    elif system == 'linux':
        # Try to detect dark mode from common desktop environments
        try:
            desktop = os.environ.get('XDG_CURRENT_DESKTOP', '').lower()
            
            # GNOME
            if 'gnome' in desktop:
                result = os.popen('gsettings get org.gnome.desktop.interface gtk-theme').read().lower()
                return 'dark' in result
                
            # KDE Plasma
            elif 'kde' in desktop:
                result = os.popen('kreadconfig5 --group "KDE" --key "ColorScheme"').read().lower()
                return 'dark' in result
                
        except:
            pass
            
    return False

def _configure_widget_styles(style, dark_mode):
    """Configure common widget styles."""
    # Common padding values
    padding = {'padding': 6}
    padding_small = {'padding': 3}
    
    # Platform-specific font settings
    if platform.system() == 'Windows':
        default_font = ('Segoe UI', 9)
    elif platform.system() == 'Darwin':  # macOS
        default_font = ('SF Pro Text', 12)
    else:  # Linux and others
        default_font = ('DejaVu Sans', 10)
    
    # Configure button styles
    style.configure('TButton', **padding)
    style.configure('Accent.TButton', **padding)
    
    # Configure notebook tabs
    style.configure('TNotebook', padding=0)
    tab_style = padding_small.copy()
    tab_style['font'] = default_font
    style.configure('TNotebook.Tab', **tab_style)
    
    # Configure entry and combobox
    style.configure('TEntry', **padding_small)
    style.configure('TCombobox', **padding_small)
    
    # Configure scrollbars with theme-appropriate colors
    scrollbar_style = padding_small.copy()
    scrollbar_style.update({
        'arrowsize': 14,
        'width': 12,
        'troughcolor': style.colors.border,
        'background': style.colors.bg,
        'bordercolor': style.colors.border,
        'darkcolor': style.colors.bg,
        'lightcolor': style.colors.bg,
        'troughrelief': 'flat',
        'relief': 'flat',
        'arrowcolor': style.colors.fg
    })
    
    style.configure('Vertical.TScrollbar', **scrollbar_style)
    style.configure('Horizontal.TScrollbar', **scrollbar_style)
    
    # Configure scrollbar elements
    style.map('Vertical.TScrollbar',
        background=[('active', style.colors.primary), ('!disabled', style.colors.border)],
        arrowcolor=[('!disabled', style.colors.fg), ('disabled', style.colors.border)],
    )
    style.map('Horizontal.TScrollbar',
        background=[('active', style.colors.primary), ('!disabled', style.colors.border)],
        arrowcolor=[('!disabled', style.colors.fg), ('disabled', style.colors.border)],
    )
    
    # Configure treeview
    treeview_style = padding_small.copy()
    treeview_style['rowheight'] = 25
    style.configure('Treeview', **treeview_style)
    style.configure('Treeview.Heading', **padding_small)
    
    # Configure labels
    label_style = {}
    if dark_mode:
        label_style['foreground'] = style.colors.fg if hasattr(style, 'colors') else 'white'
    else:
        label_style['foreground'] = style.colors.fg if hasattr(style, 'colors') else 'black'
    
    if label_style:  # Only configure if we have styles to apply
        style.configure('TLabel', **label_style)

def _apply_cross_platform_fixes(style, dark_mode):
    """Apply platform-specific fixes for better appearance."""
    system = platform.system().lower()
    
    # Common widget styles that work across platforms
    common_styles = {
        'TButton': {'font': None},
        'TLabel': {'font': None},
        'TEntry': {'font': None},
        'TCombobox': {'font': None},
        'TNotebook.Tab': {}
    }
    
    # Platform-specific adjustments
    if system == 'windows':
        font = ('Segoe UI', 9)
        common_styles['TNotebook.Tab']['padding'] = [8, 4]
        
    elif system == 'darwin':  # macOS
        font = ('SF Pro Text', 12)
        common_styles['TNotebook.Tab']['padding'] = [12, 6]
        
    elif system == 'linux':
        font = ('Ubuntu', 10) if 'ubuntu' in platform.version().lower() else ('DejaVu Sans', 10)
        common_styles['TNotebook.Tab']['padding'] = [8, 4]
    else:
        font = ('TkDefaultFont', 10)
    
    # Apply font to common widgets
    for widget, styles in common_styles.items():
        if styles.get('font') is None:
            styles['font'] = font
        try:
            style.configure(widget, **styles)
        except Exception as e:
            print(f"Warning: Could not configure {widget}: {e}")

def get_available_themes():
    """Get a list of available themes with their display names."""
    return {
        'Light': {
            'cosmo': 'Cosmo (Default Light)',
            'flatly': 'Flatly',
            'journal': 'Journal',
            'litera': 'Litera',
            'lumen': 'Lumen',
            'minty': 'Minty',
            'pulse': 'Pulse',
            'sandstone': 'Sandstone',
            'united': 'United',
            'yeti': 'Yeti',
            'morph': 'Morph',
            'simplex': 'Simplex',
            'cerculean': 'Cerculean'
        },
        'Dark': {
            'darkly': 'Darkly (Default Dark)',
            'solar': 'Solar',
            'superhero': 'Superhero',
            'cyborg': 'Cyborg',
            'vapor': 'Vapor',
            'sharish': 'Sharish',
            'hacker': 'Hacker'
        }
    }
