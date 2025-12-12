#!/usr/bin/env python3
"""
Modern launcher for ExPlot with ttkbootstrap theming.
"""
import sys
import os
import tkinter as tk
from tkinter import ttk
from pathlib import Path

# Import and disable ttkbootstrap localization to prevent msgcat errors
import ttkbootstrap.localization
ttkbootstrap.localization.initialize_localities = lambda *args, **kwargs: None

from ttkbootstrap_theme import setup_theme, get_available_themes

# Global flag to prevent multiple cleanup calls
_cleanup_called = False

def cleanup(root, app):
    """Clean up resources before exiting."""
    global _cleanup_called
    
    # Prevent multiple cleanup calls
    if _cleanup_called:
        return
    _cleanup_called = True
    
    try:
        # Close any open matplotlib figures
        import matplotlib.pyplot as plt
        plt.close('all')
        
        # If the app has a cleanup method, call it
        if hasattr(app, 'cleanup'):
            app.cleanup()
            
        # Destroy all widgets (only if root still exists)
        try:
            if root.winfo_exists():
                for widget in root.winfo_children():
                    try:
                        widget.destroy()
                    except Exception:
                        pass
        except Exception:
            # Root might already be destroyed
            pass
            
    except Exception as e:
        print(f"Error during cleanup: {e}")
    finally:
        # Destroy the root window if it still exists
        try:
            if root.winfo_exists():
                root.quit()
                root.destroy()
        except Exception:
            # Window already destroyed, which is fine
            pass

def main():
    """Launch the application with ttkbootstrap theming."""
    # Create the root window with ttkbootstrap
    root = tk.Tk()
    
    # Set application title
    root.title("ExPlot")
    
    # Set up proper window close handling
    def on_closing():
        cleanup(root, app)
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    # Set up the theme (auto-detect dark/light mode)
    style = setup_theme(root)
    
    # Import the main application after setting up the theme
    from explot import ExPlotApp
    
    # Create the application
    app = ExPlotApp(root)
    
    # Set window size and position
    window_width = 1200
    window_height = 800
    
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    
    root.geometry(f'{window_width}x{window_height}+{x}+{y}')
    root.minsize(1000, 700)
    
    # Add theme switcher to the menu
    _add_theme_switcher(root, style, app)
    
    # Start the main loop
    root.mainloop()

def _add_theme_switcher(root, style, app):
    """
    Add a theme switcher to the application menu.
    
    Args:
        root: The root window
        style: The ttkbootstrap.Style instance
        app: The main application instance
    """
    try:
        # Get the menu bar from the application
        menubar = root.nametowidget(root['menu'])
        
        # Create a View menu if it doesn't exist
        view_menu = None
        for i in range(menubar.index('end') + 1):
            try:
                if menubar.entrycget(i, 'label').lower() == 'view':
                    view_menu = menubar.nametowidget(f"{menubar}.menu{i+1}")
                    break
            except tk.TclError:
                continue
        
        # If no View menu found, create one
        if view_menu is None:
            view_menu = tk.Menu(menubar, tearoff=0)
            menubar.add_cascade(label="View", menu=view_menu)
        
        # Add theme submenu
        theme_submenu = tk.Menu(view_menu, tearoff=0)
        view_menu.add_cascade(label="Themes", menu=theme_submenu)
        
        # Define available themes (only include valid ttkbootstrap themes)
        # Note: Custom themes (nord, nordic) are handled specially
        light_themes = [
            "cosmo", "flatly", "journal", "litera", "lumen", 
            "minty", "pulse", "sandstone", "united", "yeti", 
            "morph", "simplex", "cerculean"
        ]
        
        dark_themes = [
            "darkly", "solar", "superhero", "cyborg", "vapor"
        ]
        
        custom_themes = [
            ("nord", True),     # (name, is_dark)
            ("nordic", True),
        ]
        
        # Add light themes
        light_menu = tk.Menu(theme_submenu, tearoff=0)
        theme_submenu.add_cascade(label="Light Themes", menu=light_menu)
        for theme in sorted(light_themes):
            light_menu.add_radiobutton(
                label=theme.title().replace('_', ' '),
                command=lambda t=theme: _change_theme(style, t, False, app)
            )
        
        # Add dark themes
        dark_menu = tk.Menu(theme_submenu, tearoff=0)
        theme_submenu.add_cascade(label="Dark Themes", menu=dark_menu)
        
        # Add built-in dark themes
        for theme in sorted(dark_themes):
            dark_menu.add_radiobutton(
                label=theme.title().replace('_', ' '),
                command=lambda t=theme: _change_theme(style, t, True, app)
            )
            
        # Add custom themes
        if custom_themes:
            dark_menu.add_separator()
            for theme_name, is_dark in custom_themes:
                if is_dark:
                    dark_menu.add_radiobutton(
                        label=theme_name.title().replace('_', ' '),
                        command=lambda t=theme_name, d=is_dark: _change_theme(style, t, d, app)
                    )
        
        # Add auto-detect option
        theme_submenu.add_separator()
        theme_submenu.add_command(
            label="Auto-detect System Theme",
            command=lambda: _auto_detect_theme(style, app)
        )
        
        # Load saved theme from application settings
        try:
            if hasattr(app, 'theme_name') and hasattr(app, 'dark_mode'):
                saved_theme = app.theme_name
                dark_mode = app.dark_mode
                # Check if theme is in any of our theme lists
                all_themes = light_themes + dark_themes + [t[0] for t in custom_themes]
                if saved_theme in all_themes:
                    # Use silent=True to prevent save message during initialization
                    _change_theme(style, saved_theme, dark_mode, app, update_menu=False, silent=True)
        except Exception as e:
            print(f"Warning: Could not load saved theme: {e}")
            import traceback
            traceback.print_exc()

    except Exception as e:
        print(f"Error adding theme switcher: {e}")
        import traceback
        traceback.print_exc()

def _change_theme(style, theme_name, dark_mode, app, update_menu=True, silent=False):
    """
    Change the current theme.
    
    Args:
        style: The ttkbootstrap.Style instance
        theme_name: Name of the theme to switch to (case-insensitive for custom themes)
        dark_mode: Whether the theme is a dark theme
        app: The main application instance
        update_menu: Whether to update the menu checkmarks
        silent: If True, suppress save messages and dialogs
    """
    try:
        # Define dark themes (all lowercase for comparison)
        dark_themes = [
            'solar', 'superhero', 'darkly', 'cyborg', 'vapor',
            'sharish', 'hacker', 'nord', 'nordic'
        ]
        
        # Normalize theme name for comparison (always use lowercase for custom themes)
        theme_lower = theme_name.lower()
        is_custom_theme = theme_lower in ['nord', 'nordic']
        
        # Special handling for custom themes
        if is_custom_theme:
            # For custom themes, we need to create them using our setup functions
            from ttkbootstrap_theme import _setup_nord_theme
            
            # Use the normalized theme name (lowercase)
            theme_name = theme_lower
            
            # Apply the appropriate theme
            if theme_name == 'nord':
                style = _setup_nord_theme(style, 'nord')
            elif theme_name == 'nordic':
                style = _setup_nord_theme(style, 'nordic')
                
            # Ensure the theme is set to use the correct name
            style.theme_use(theme_name)
            
            # Set dark_mode based on the theme (both are dark themes)
            dark_mode = True
        else:
            # For standard ttkbootstrap themes, switch normally
            available_themes = style.theme_names()
            if theme_name not in available_themes:
                raise ValueError(f"Theme '{theme_name}' is not available. Available themes: {', '.join(available_themes)}")
            style.theme_use(theme_name)
            dark_mode = theme_lower in [t.lower() for t in dark_themes]
        
        # Update theme-dependent styles
        _update_theme_dependent_styles(style, dark_mode)
        
        # Save the selected theme in application settings
        if hasattr(app, 'save_user_preferences') and hasattr(app, 'theme_name'):
            app.theme_name = theme_name
            app.dark_mode = dark_mode
            # Save theme settings to file
            if hasattr(app, '_save_theme_settings'):
                app._save_theme_settings(theme_name, dark_mode)
            # Only show save message if not in silent mode
            if silent:
                # Save without showing message
                app.save_user_preferences(silent=True)
            else:
                app.save_user_preferences()
        
        # Update the application colors if needed
        if hasattr(app, 'update_colors'):
            app.update_colors()
            
    except Exception as e:
        print(f"Error changing theme: {e}")
        import traceback
        traceback.print_exc()

def _update_theme_dependent_styles(style, dark_mode):
    """
    Update styles that depend on the current theme.
    
    Args:
        style: The ttkbootstrap.Style instance
        dark_mode: Whether the current theme is a dark theme
    """
    # Common padding values
    padding = {'padding': 6}
    padding_small = {'padding': 4}
    
    # Platform-specific font settings
    if sys.platform == 'Windows':
        default_font = ('Segoe UI', 9)
    elif sys.platform == 'Darwin':  # macOS
        default_font = ('SF Pro Text', 12)
    else:  # Linux and others
        default_font = ('DejaVu Sans', 10)
    
    # Get current theme colors
    try:
        bg = style.colors.bg if hasattr(style.colors, 'bg') else '#ffffff'
        fg = style.colors.fg if hasattr(style.colors, 'fg') else '#000000'
        select_bg = style.colors.selectbg if hasattr(style.colors, 'selectbg') else '#0078d7'
        select_fg = style.colors.selectfg if hasattr(style.colors, 'selectfg') else '#ffffff'
    except:
        bg = '#ffffff' if not dark_mode else '#2e2e2e'
        fg = '#000000' if not dark_mode else '#ffffff'
        select_bg = '#0078d7'
        select_fg = '#ffffff'
    
    # Configure button styles
    style.configure('TButton', **padding)
    style.configure('Accent.TButton', **padding)
    
    # Configure notebook tabs
    style.configure('TNotebook', padding=0)
    tab_style = padding_small.copy()
    tab_style['font'] = default_font
    style.configure('TNotebook.Tab', **tab_style)
                       
    # Configure entry and combobox
    entry_style = {
        'padding': 5,
        'font': ('Segoe UI', 10) if sys.platform == 'win32' else
               ('SF Pro Text', 12) if sys.platform == 'darwin' else
               ('Ubuntu', 10)
    }
    style.configure('TEntry', **entry_style)
    style.configure('TCombobox', **entry_style)
    
    # Configure hover and active states
    hover_bg = '#e0e0e0' if not dark_mode else '#404040'
    pressed_bg = '#d0d0d0' if not dark_mode else '#505050'
    
    style.map('TButton',
             background=[
                 ('active', '!disabled', hover_bg),
                 ('pressed', pressed_bg)
             ],
             relief=[
                 ('pressed', 'sunken'),
                 ('!pressed', 'raised')
             ])
             
    style.map('TEntry',
             fieldbackground=[('readonly', '!disabled', bg)],
             foreground=[('readonly', '!disabled', fg)])
             
    style.map('TCombobox',
             fieldbackground=[('readonly', '!disabled', bg)],
             selectbackground=[('readonly', '!disabled', select_bg)],
             selectforeground=[('readonly', '!disabled', select_fg)])
             
    # Configure scrollbars
    scrollbar_style = {
        'arrowsize': 14,
        'width': 12
    }
    style.configure('Vertical.TScrollbar', **scrollbar_style)
    style.configure('Horizontal.TScrollbar', **scrollbar_style)

def _auto_detect_theme(style, app):
    """Auto-detect the system theme and apply the appropriate theme.
    
    Args:
        style: The ttkbootstrap.Style instance
        app: The main application instance
    """
    try:
        # Check if we have a saved preference first
        if hasattr(app, 'theme_name') and app.theme_name:
            saved_theme = app.theme_name
            dark_mode = app.dark_mode
            _change_theme(style, saved_theme, dark_mode, app, silent=True)
            return
            
        # No saved preference, try to detect system theme
        try:
            import darkdetect
            if darkdetect.isDark():
                # Default to Nord theme for dark mode
                _change_theme(style, 'nord', True, app, silent=True)
            else:
                # Default to Cosmo for light mode
                _change_theme(style, 'cosmo', False, app, silent=True)
        except ImportError:
            # Fallback if darkdetect is not available
            _change_theme(style, 'cosmo', False, app, silent=True)
            
    except Exception as e:
        print(f"Warning: Could not auto-detect theme: {e}")
        # Fall back to light theme
        _change_theme(style, 'cosmo', False, app, silent=True)
        return  # Exit after handling the error

if __name__ == "__main__":
    main()
