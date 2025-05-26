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
        'sharish', 'hacker'
    ]
    
    # Default theme selection
    DEFAULT_LIGHT_THEME = 'cosmo'
    DEFAULT_DARK_THEME = 'darkly'
    
    # Determine if we should use dark mode if not explicitly set
    if dark_mode is None:
        dark_mode = _is_dark_mode()
    
    # Select theme based on dark mode preference if not specified
    if theme_name is None:
        theme_name = DEFAULT_DARK_THEME if dark_mode else DEFAULT_LIGHT_THEME
    
    # Create and configure the style
    try:
        style = Style(theme=theme_name)
        
        # Update dark_mode based on the actual theme used
        dark_mode = theme_name.lower() in [t.lower() for t in DARK_THEMES]
        
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
    
    # Configure scrollbars
    scrollbar_style = padding_small.copy()
    scrollbar_style['arrowsize'] = 14
    style.configure('Vertical.TScrollbar', **scrollbar_style)
    style.configure('Horizontal.TScrollbar', **scrollbar_style)
    
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
