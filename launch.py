#!/usr/bin/env python3
"""
Modern launcher for ExPlot with ttkbootstrap theming.
"""
import sys
import os
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from themes.ttkbootstrap_theme import setup_theme, get_available_themes

def main():
    """Launch the application with ttkbootstrap theming."""
    # Create the root window with ttkbootstrap
    root = tk.Tk()
    
    # Set application title
    root.title("ExPlot")
    
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
        light_themes = [
            "cosmo", "flatly", "journal", "litera", "lumen", 
            "minty", "pulse", "sandstone", "united", "yeti", 
            "morph", "simplex", "cerculean"
        ]
        
        dark_themes = [
            "darkly", "solar", "superhero", "cyborg", "vapor"
        ]
        
        # Add light themes
        light_menu = tk.Menu(theme_submenu, tearoff=0)
        theme_submenu.add_cascade(label="Light Themes", menu=light_menu)
        for theme in light_themes:
            light_menu.add_radiobutton(
                label=theme.title(),
                command=lambda t=theme: _change_theme(style, t, False, app)
            )
        
        # Add dark themes
        dark_menu = tk.Menu(theme_submenu, tearoff=0)
        theme_submenu.add_cascade(label="Dark Themes", menu=dark_menu)
        for theme in dark_themes:
            dark_menu.add_radiobutton(
                label=theme.title(),
                command=lambda t=theme: _change_theme(style, t, True, app)
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
                if saved_theme in light_themes + dark_themes:
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
        theme_name: Name of the theme to switch to
        dark_mode: Whether the theme is a dark theme
        app: The main application instance
        update_menu: Whether to update the menu checkmarks
        silent: If True, suppress save messages and dialogs
    """
    try:
        # Get available themes
        available_themes = style.theme_names()
        
        # Check if theme is available
        if theme_name not in available_themes:
            raise ValueError(f"Theme '{theme_name}' is not available. Available themes: {', '.join(available_themes)}")
            
        # Set the new theme
        style.theme_use(theme_name)
        
        # Update theme-dependent styles
        _update_theme_dependent_styles(style, dark_mode)
        
        # Save the selected theme in application settings
        if hasattr(app, 'save_user_preferences') and hasattr(app, 'theme_name'):
            app.theme_name = theme_name
            app.dark_mode = dark_mode
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
    # Common colors
    bg = style.colors.bg if hasattr(style, 'colors') else '#f0f0f0'
    fg = style.colors.fg if hasattr(style, 'colors') else '#000000'
    
    # Configure common widget styles
    style.configure('TButton', 
                   padding=6,
                   font=('Segoe UI', 10) if sys.platform == 'win32' else 
                       ('SF Pro Text', 12) if sys.platform == 'darwin' else 
                       ('Ubuntu', 10))
                       
    style.configure('TNotebook', padding=0)
    style.configure('TNotebook.Tab', 
                   padding=[12, 4],
                   font=('Segoe UI', 9) if sys.platform == 'win32' else
                       ('SF Pro Text', 11) if sys.platform == 'darwin' else
                       ('Ubuntu', 10))
                       
    style.configure('TEntry', 
                  padding=5,
                  font=('Segoe UI', 10) if sys.platform == 'win32' else
                      ('SF Pro Text', 12) if sys.platform == 'darwin' else
                      ('Ubuntu', 10))
                  
    style.configure('TCombobox',
                  padding=5,
                  font=('Segoe UI', 10) if sys.platform == 'win32' else
                      ('SF Pro Text', 12) if sys.platform == 'darwin' else
                      ('Ubuntu', 10))
    
    # Configure hover and active states
    style.map('TButton',
             background=[('active', '!disabled', '#e0e0e0'),
                       ('pressed', '#d0d0d0')],
             relief=[('pressed', 'sunken'), ('!pressed', 'raised')])
             
    style.map('TEntry',
             fieldbackground=[('readonly', '!disabled', bg)],
             foreground=[('readonly', '!disabled', fg)])
             
    style.map('TCombobox',
             fieldbackground=[('readonly', '!disabled', bg)],
             selectbackground=[('readonly', '!disabled', bg)],
             selectforeground=[('readonly', '!disabled', fg)])
             
    # Configure scrollbars
    style.configure('Vertical.TScrollbar',
                  arrowsize=14,
                  width=12)
                  
    style.configure('Horizontal.TScrollbar',
                  arrowsize=14,
                  width=12)

def _auto_detect_theme(style, app):
    """Auto-detect the system theme and apply the appropriate theme.
    
    Args:
        style: The ttkbootstrap.Style instance
        app: The main application instance
    """
    try:
        if sys.platform == 'darwin':
            # macOS
            import subprocess
            cmd = 'defaults read -g AppleInterfaceStyle'
            dark_mode = subprocess.run(
                cmd, shell=True, capture_output=True, text=True
            ).stdout.strip().lower() == 'dark'
            theme = 'darkly' if dark_mode else 'cosmo'
        elif sys.platform == 'win32':
            # Windows
            import ctypes
            dark_mode = not ctypes.windll.uxtheme.IsThemeActive()
            theme = 'darkly' if dark_mode else 'cosmo'
        else:
            # Linux - try to detect dark mode from gsettings
            try:
                import subprocess
                result = subprocess.run(
                    ['gsettings', 'get', 'org.gnome.desktop.interface', 'gtk-theme'],
                    capture_output=True, text=True
                )
                dark_mode = 'dark' in result.stdout.lower()
                theme = 'darkly' if dark_mode else 'cosmo'
            except:
                # Fallback to light theme if detection fails
                theme = 'cosmo'
                dark_mode = False
        
        _change_theme(style, theme, dark_mode, app)
        
    except Exception as e:
        print(f"Warning: Could not auto-detect theme: {e}")
        # Fall back to light theme
        _change_theme(style, 'cosmo', False, app)
        
        # Tab style
        tab_style = {
            'padding': [10, 4],
            'font': ('Segoe UI', 10) if platform.system() == 'Windows' else 
                   ('SF Pro Text', 12) if platform.system() == 'Darwin' else 
                   ('Ubuntu', 10) if 'ubuntu' in platform.version().lower() else ('DejaVu Sans', 10)
        }
        
        # Apply notebook and tab styles
        style.configure('TNotebook', **notebook_style)
        style.configure('TNotebook.Tab', **tab_style)
        
        # Configure treeview colors
        if hasattr(colors, 'bg') and hasattr(colors, 'fg'):
            style.configure('Treeview', 
                          background=colors.bg,
                          foreground=colors.fg,
                          fieldbackground=colors.bg,
                          borderwidth=1)
            
            style.map('Treeview',
                    background=[('selected', colors.selectbg if hasattr(colors, 'selectbg') else '#0078d7')],
                    foreground=[('selected', colors.selectfg if hasattr(colors, 'selectfg') else 'white')])
        
        # Configure entry and combobox
        style.configure('TEntry', fieldbackground=colors.inputbg if hasattr(colors, 'inputbg') else 'white')
        style.configure('TCombobox', fieldbackground=colors.inputbg if hasattr(colors, 'inputbg') else 'white')
        
    except Exception as e:
        print(f"Error updating theme-dependent styles: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
