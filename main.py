import time
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, Menu
import win32gui
import win32process
import win32con
import psutil
import pythoncom
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import shutil
import os
import json

DARK_THEME = {
    'bg': '#1a1a1a',
    'fg': '#e0e0e0',
    'entry_bg': '#2a2a2a',
    'entry_fg': '#ffffff',
    'console_bg': '#121212',
    'console_fg': '#d0d0d0',
    'button_bg': '#2a2a2a',
    'button_fg': '#ffffff',
    'slider_bg': '#2a2a2a',
    'slider_fg': '#ffffff',
    'select_bg': '#3a3a3a',
    'select_fg': '#ffffff',
    'dropdown_bg': '#2a2a2a',
    'dropdown_fg': '#ffffff',
    'dropdown_field_bg': '#2a2a2a',
    'dropdown_arrow': '#ffffff',
    'disabled_fg': '#707070'
}

LIGHT_THEME = {
    'bg': '#f0f0f0',
    'fg': '#000000',
    'entry_bg': '#ffffff',
    'entry_fg': '#000000',
    'console_bg': '#ffffff',
    'console_fg': '#000000',
    'button_bg': '#e0e0e0',
    'button_fg': '#000000',
    'slider_bg': '#e0e0e0',
    'slider_fg': '#000000',
    'select_bg': '#d0d0d0',
    'select_fg': '#000000',
    'dropdown_bg': '#ffffff',
    'dropdown_fg': '#000000',
    'dropdown_field_bg': '#ffffff',
    'dropdown_arrow': '#000000',
    'disabled_fg': '#a0a0a0'
}
def get_foreground_process_name():
    hwnd = win32gui.GetForegroundWindow()
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    try:
        return psutil.Process(pid).name()
    except psutil.NoSuchProcess:
        return None


def set_app_volume(app_name, volume_level):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        if session.Process and app_name.lower() in session.Process.name().lower():
            volume.SetMasterVolume(volume_level, None)
            return True
    return False


def create_image(width, height, color1, color2):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((0, 0, width, height), fill=color1)
    dc.ellipse((width // 4, height // 4, 3 * width // 4, 3 * height // 4), fill=color2)
    return image



class VolumeMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.iconbitmap("C:\\Users\\henym\\Downloads\\Screenshot 2025-07-21 011221.ico")
        self.root.title("Auto Volume Mixer")
        self.running = False
        self.thread = None
        self.last_state = None
        self.tray_icon = None
        self.minimized_to_tray = False
        self.config_file = os.path.join(os.getenv('APPDATA'), 'AutoVolumeMixer', 'settings.json')

        os.makedirs(os.path.dirname(self.config_file), exist_ok=True)

        self.root.protocol('WM_DELETE_WINDOW', self.root.destroy)

        self.volume_in_var = tk.DoubleVar(value=1.0)
        self.volume_out_var = tk.DoubleVar(value=0.1)
        self.app_name_var = tk.StringVar()

        # App dropdown
        ttk.Label(root, text="Select Application:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.app_dropdown = ttk.Combobox(root, textvariable=self.app_name_var, width=30)
        self.app_dropdown.grid(row=0, column=1, columnspan=2, sticky="w", padx=5, pady=5)

        # Volume sliders
        tk.Label(root, text="Volume when focused:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        tk.Scale(root, from_=0.0, to=1.0, orient="horizontal", resolution=0.01,
                 variable=self.volume_in_var, length=200).grid(row=1, column=1, sticky="w")
        tk.Label(root, textvariable=self.volume_in_var).grid(row=1, column=2, sticky="w", padx=5)

        tk.Label(root, text="Volume when unfocused:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        tk.Scale(root, from_=0.0, to=1.0, orient="horizontal", resolution=0.01,
                 variable=self.volume_out_var, length=200).grid(row=2, column=1, sticky="w")
        tk.Label(root, textvariable=self.volume_out_var).grid(row=2, column=2, sticky="w", padx=5)

        # Buttons
        self.toggle_button = ttk.Button(root, text="Start", command=self.toggle_monitoring)
        self.toggle_button.grid(row=3, column=0, columnspan=2, pady=10)

        self.refresh_button = ttk.Button(root, text="Manual Refresh", command=self.refresh_app_list)
        self.refresh_button.grid(row=0, column=2, pady=10)

        self.save_button = ttk.Button(root, text="Save Current Settings", command=self.save_current_settings)
        self.save_button.grid(row=4, column=0, pady=5)

        # Add dark mode toggle button
        self.dark_mode_var = tk.BooleanVar(value=False)
        self.dark_mode_button = ttk.Button(root, text="🌙 Dark Mode", command=self.toggle_dark_mode)
        self.dark_mode_button.grid(row=4, column=1, pady=5)

        # Add minimize to tray button
        self.tray_button = ttk.Button(root, text="▼ Minimize to Tray", command=self.minimize_to_tray)
        self.tray_button.grid(row=3, column=1, pady=10)

        self.startup_var = tk.BooleanVar(value=False)
        self.startup_checkbox = ttk.Checkbutton(root, text="Add to Startup", variable=self.startup_var, command=self.toggle_startup)
        self.startup_checkbox.grid(row=4, column=2, sticky="e", padx=5, pady=(0, 10))

        self.check_startup_status()

        # Console
        self.console = scrolledtext.ScrolledText(root, width=70, height=15, state="disabled", font=("Courier", 9))
        self.console.grid(row=5, column=0, columnspan=3, padx=5, pady=10)

        self.log("Welcome!")

        # Start auto-refresh
        self.refresh_app_list()
        self.auto_refresh()

    def toggle_dark_mode(self):
        """Toggle between dark and light mode"""
        self.dark_mode_var.set(not self.dark_mode_var.get())
        self.apply_theme()

    def apply_theme(self):
        """Apply the current theme (dark or light)"""
        theme = DARK_THEME if self.dark_mode_var.get() else LIGHT_THEME

        # Update button text
        self.dark_mode_button.config(text="🌞 Light Mode" if self.dark_mode_var.get() else "🌙 Dark Mode")

        # Apply theme to root window
        self.root.configure(bg=theme['bg'])

        # Apply to all standard widgets
        for child in self.root.winfo_children():
            if isinstance(child, tk.Label):
                child.configure(bg=theme['bg'], fg=theme['fg'])
            elif isinstance(child, tk.Scale):
                child.configure(
                    bg=theme['slider_bg'],
                    fg=theme['slider_fg'],
                    highlightbackground=theme['bg'],
                    troughcolor=theme['bg'],
                    activebackground=theme['select_bg']
                )

        # Special handling for console
        self.console.configure(
            bg=theme['console_bg'],
            fg=theme['console_fg'],
            insertbackground=theme['console_fg']
        )

        # Create and configure ttk style
        style = ttk.Style()
        style.theme_use('clam')  # Best theme for customization

        # Configure general styles
        style.configure('.',
                        background=theme['bg'],
                        foreground=theme['fg'],
                        fieldbackground=theme['entry_bg'])

        # Button styling
        style.configure('TButton',
                        background=theme['button_bg'],
                        foreground=theme['button_fg'],
                        bordercolor=theme['bg'],
                        darkcolor=theme['button_bg'],
                        lightcolor=theme['button_bg'])

        # Combobox styling (dropdown)
        style.configure('TCombobox',
                        fieldbackground=theme['dropdown_field_bg'],
                        background=theme['dropdown_bg'],
                        foreground=theme['dropdown_fg'],
                        selectbackground=theme['select_bg'],
                        selectforeground=theme['select_fg'],
                        arrowcolor=theme['dropdown_arrow'])

        style.map('TCombobox',
                  fieldbackground=[('readonly', theme['dropdown_field_bg'])],
                  selectbackground=[('readonly', theme['select_bg'])],
                  selectforeground=[('readonly', theme['select_fg'])])

        # Checkbutton styling
        style.configure('TCheckbutton',
                        background=theme['bg'],
                        foreground=theme['fg'],
                        indicatorbackground=theme['entry_bg'])

        # Entry styling
        style.configure('TEntry',
                        fieldbackground=theme['entry_bg'],
                        foreground=theme['entry_fg'],
                        insertcolor=theme['entry_fg'])

        # Disabled state styling
        style.map('TButton',
                  background=[('disabled', theme['bg'])],
                  foreground=[('disabled', theme['disabled_fg'])])

        style.map('TEntry',
                  fieldbackground=[('disabled', theme['bg'])],
                  foreground=[('disabled', theme['disabled_fg'])])

        style.map('TCombobox',
                  fieldbackground=[('disabled', theme['bg'])],
                  foreground=[('disabled', theme['disabled_fg'])])

        # Force update of all widgets
        self.app_dropdown.update_idletasks()
        for child in self.root.winfo_children():
            child.update_idletasks()

    def save_current_settings(self):
        app_name = self.app_name_var.get().strip()
        if not app_name:
            self.log("❌ No application selected to save")
            return

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    all_settings = json.load(f)
            else:
                all_settings = {}

            all_settings[app_name] = {
                'volume_in': self.volume_in_var.get(),
                'volume_out': self.volume_out_var.get()
            }

            # Save back to file
            with open(self.config_file, 'w') as f:
                json.dump(all_settings, f)

            self.log(f"✅ Saved settings at '{self.config_file}'")
        except Exception as e:
            self.log(f"❌ Failed to save settings: {e}")

    def load_app_settings(self, app_name=None):
        if app_name is None:
            app_name = self.app_name_var.get().strip()
            if not app_name:
                return False

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    all_settings = json.load(f)

                if app_name in all_settings:
                    settings = all_settings[app_name]
                    self.volume_in_var.set(settings.get('volume_in', 1.0))
                    self.volume_out_var.set(settings.get('volume_out', 0.1))
                    self.log(f"🔁 Loaded settings from '{self.config_file}'")
                    return True
        except Exception as e:
            self.log(f"❌ Failed to load settings: {e}")
        return False

    def log(self, message):
        def append():
            self.console.configure(state="normal")
            self.console.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
            self.console.see(tk.END)
            self.console.configure(state="disabled")

        self.root.after(0, append)

    def refresh_app_list(self):
        if not self.minimized_to_tray:
            def enum_windows(hwnd, apps):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        proc = psutil.Process(pid)
                        apps.add(proc.name())
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                return True

            apps = set()
            win32gui.EnumWindows(enum_windows, apps)
            app_list = sorted(apps)

            current_selection = self.app_name_var.get()
            self.app_dropdown['values'] = app_list

            # Add binding to load settings when app is selected
            def on_app_selected(event):
                self.load_app_settings()

            self.app_dropdown.bind('<<ComboboxSelected>>', on_app_selected)

            if current_selection in app_list:
                self.app_name_var.set(current_selection)
                self.load_app_settings()
            elif app_list:
                self.app_name_var.set(app_list[0])
                self.load_app_settings()
            else:
                self.app_name_var.set("")
            self.log("Refreshed App List")

    def auto_refresh(self):
        self.refresh_app_list()
        self.root.after(10000, self.auto_refresh)  # refresh every second

    def toggle_monitoring(self):
        if not self.running:
            self.start_monitoring()
        else:
            self.stop_monitoring()

    def start_monitoring(self):
        app_name = self.app_name_var.get().strip()
        if not app_name:
            self.log("❌ Please select an application.")
            return
        try:
            vol_in = float(self.volume_in_var.get())
            vol_out = float(self.volume_out_var.get())
            if not (0.0 <= vol_in <= 1.0 and 0.0 <= vol_out <= 1.0):
                raise ValueError
        except ValueError:
            self.log("❌ Volume values must be between 0.0 and 1.0")
            return

        self.app_name = app_name
        self.running = True
        self.toggle_button.config(text="Stop")
        self.last_state = None
        self.thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.thread.start()
        self.log(f"🟢 Started monitoring '{self.app_name}'.")

    def monitor_loop(self):
        pythoncom.CoInitialize()
        while self.running:
            current_proc = get_foreground_process_name()
            vol_in = self.volume_in_var.get()
            vol_out = self.volume_out_var.get()

            if current_proc and self.app_name.lower() in current_proc.lower():
                if self.last_state != "in":
                    if set_app_volume(self.app_name, vol_in):
                        self.log(f"[FOCUSED] Volume → {vol_in:.2f}")
                    self.last_state = "in"
                else:
                    set_app_volume(self.app_name, vol_in)
            else:
                if self.last_state != "out":
                    if set_app_volume(self.app_name, vol_out):
                        self.log(f"[UNFOCUSED] Volume → {vol_out:.2f}")
                    self.last_state = "out"
                else:
                    set_app_volume(self.app_name, vol_out)
            time.sleep(0.5)

    def stop_monitoring(self):
        self.toggle_button.config(text="Start")
        self.log("🛑 Stopped monitoring.")
        self.volume_out_var.set(1.0)
        self.volume_in_var.set(1.0)
        set_app_volume(self.app_name, 1.0)
        self.running = False
        self.log("Reset audio levels to 1.0")

    def minimize_to_tray(self):
        self.root.withdraw()  # Hide the main window
        self.minimized_to_tray = True

        # Create system tray icon
        image = create_image(64, 64, 'black', 'white')
        menu = pystray.Menu(
            item('AVM V1.2', self.restore_from_tray),
            item('Show', self.restore_from_tray),
            item('Exit', self.quit_application)
        )

        if self.tray_icon:
            self.tray_icon.stop()

        self.tray_icon = pystray.Icon("volume_monitor", image, "Auto Volume Mixer", menu)

        # Run the tray icon in a separate thread
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
        self.log("Minimized to system tray")

    def restore_from_tray(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

        self.minimized_to_tray = False
        self.root.deiconify()
        self.root.lift()
        self.log("Restored from system tray")
        self.refresh_app_list()

    def quit_application(self, icon=None, item=None):
        if self.tray_icon:
            self.tray_icon.stop()
        self.volume_out_var.set(1.0)
        self.volume_in_var.set(1.0)
        set_app_volume(self.app_name, 1.0)
        self.running = False
        self.root.quit()

    def get_startup_path(self):
        return os.path.join(os.getenv('APPDATA'), r"Microsoft\Windows\Start Menu\Programs\Startup", "autoVolumeMixer.exe")

    def check_startup_status(self):
        if os.path.exists(self.get_startup_path()):
            self.startup_var.set(True)
            self.log("🔁 Already in Startup folder")
        else:
            self.startup_var.set(False)

    def toggle_startup(self):
        exe_name = "autoVolumeMixer.exe"
        source_path = os.path.join(os.getcwd(), exe_name)
        target_path = self.get_startup_path()

        if self.startup_var.get():
            if not os.path.exists(source_path):
                self.log(f"❌ '{exe_name}' not found in current directory.")
                self.startup_var.set(False)
                return
            try:
                shutil.copyfile(source_path, target_path)
                self.log(f"✅ '{exe_name}' added to Startup.")
            except Exception as e:
                self.log(f"❌ Failed to copy to startup: {e}")
                self.startup_var.set(False)
        else:
            try:
                if os.path.exists(target_path):
                    os.remove(target_path)
                    self.log(f"🗑️Removed from Startup.")
            except Exception as e:
                self.log(f"❌ Failed to remove from startup: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VolumeMonitorApp(root)
    root.mainloop()
