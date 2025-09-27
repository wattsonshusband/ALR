import psutil, pywintypes, win32con, win32api, ctypes, winreg, sv_ttk, sys
from json import loads, dumps, JSONDecodeError
from time import sleep
from tkinter import Tk, ttk, IntVar, DISABLED, NORMAL, ACTIVE
from os import getenv, mkdir, path
from pathlib import Path
from ctypes import wintypes
from sys import getwindowsversion
from pywinstyles import change_header_color, apply_style
from pystray import Menu, MenuItem as item
from pystray import Icon
from threading import enumerate as enum_threads
from threading import Thread, Event
from PIL import Image

import win32com.client

APP_NAME = "ALR"
ENUM_CURRENT_SETTINGS = -1
GAME_PATH_KEYWORDS = ['steamapps', 'epic games', 'origin', 'battle.net', 'blizzard', 'ubisoft', 'rockstar games', 'gog galaxy', 'roblox']

def enable_startup():
    exe_path = sys.executable
    
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r'Software\Microsoft\Windows\CurrentVersion\Run',
        0, winreg.KEY_SET_VALUE
    ) as key: 
        winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, exe_path)
        
def disable_startup():
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r'Software\Microsoft\Windows\CurrentVersion\Run',
            0, winreg.KEY_SET_VALUE
        ) as key: 
            winreg.DeleteValue(key, APP_NAME)
    except FileNotFoundError:
        pass
        
def is_startup_enabled():
    try:
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run"
        ) as key:
            value, _ = winreg.QueryValueEx(key, APP_NAME)
            print(value)
            return True
    except FileNotFoundError:
        return False

class ResolutionChanger:
    def __init__(self):
        self.appdata_path = getenv('APPDATA')
        if not self.appdata_path:
            raise EnvironmentError('APPDATA environment variable not found.')

        self.is_valid_process_running = False
        self.has_changed_res = False
        self.last_dark_mode_check = False
        self.validated_process = ""
        self.all_open_processes = {}
        
        self.load_processes()

        self.folder_path = Path(self.appdata_path) / 'ALR'
        self.config_path = Path(self.appdata_path) / 'ALR' / 'config.json'
        self.game_settings_path = Path(self.appdata_path) / 'ALR' / 'game_settings.json'

        if not self.folder_path.exists():
            mkdir(self.folder_path)

        if not self.game_settings_path.exists():
            with open(self.game_settings_path, 'x+') as f:
                self.game_settings_data = {}
                f.write(dumps(self.game_settings_data, indent=1))
        else:
            self.game_settings_data = self.read_game_config()
            
        if not self.config_path.exists():
            with open(self.config_path, 'x+') as f:
                self.config_data = { "run_on_startup": False, "show_minimised": False }
                f.write(dumps(self.config_data, indent=1))
        else:
            self.config_data = self.read_config()
        
        self.start_minimised = self.config_data['show_minimised']
        self.run_on_startup = self.config_data['run_on_startup']
        
        if is_startup_enabled():
            self.check_startup_location()
        
        self.valid_process_names = self.game_settings_data.keys()

        self.find_valid_process_event = Event()
        self.find_valid_process_thread = Thread(target=self.find_valid_process, daemon=True, name="find_valid_process")
        self.find_valid_process_thread.start()
        
        self.icon_thread = None

        self.root = Tk()
        self.root.title('ALR')
        self.root.geometry('300x290')
        self.root.iconbitmap(resource_path('icon.ico'))
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.minimise)
        
        self.is_hidden = self.root.state() == "withdrawn"
        
        self.tray()

        self.y_var = IntVar(self.root, value=1080)
        self.x_var = IntVar(self.root, value=1920)
        self.rr_var = IntVar(self.root, value=60)

        self.vcmd = self.root.register(self.validate_input)
        
        ttk.Label(self.root, text="Process:").pack(padx=(5, 1), pady=1, anchor='w')
        self.process_selector = ttk.Combobox(self.root, state='readonly', width=30, values=list(self.all_open_processes.values()))
        self.process_selector.set('Select a process')
        self.process_selector.pack(pady=(1, 4), padx=1, anchor='center')

        ttk.Separator(self.root, orient='horizontal').pack()

        ttk.Label(self.root, text="Width:").pack(padx=(5, 1), pady=1, anchor='w')
        self.x_input = ttk.Entry(self.root, width=40, textvariable=self.x_var, validate="key", validatecommand=(self.vcmd, "%P"))
        self.x_input.pack(padx=(5, 5), pady=1, anchor='w')

        ttk.Label(self.root, text="Height:").pack(padx=(5, 1), pady=1, anchor='w')
        self.y_input = ttk.Entry(self.root, width=40, textvariable=self.y_var, validate="key", validatecommand=(self.vcmd, "%P"))
        self.y_input.pack(padx=(5, 5), pady=1, anchor='w')

        ttk.Label(self.root, text="Refresh Rate:").pack(padx=(5, 1), pady=1, anchor='w')
        self.rr_input = ttk.Entry(self.root, width=40, textvariable=self.rr_var, validate="key", validatecommand=(self.vcmd, "%P"))
        self.rr_input.pack(padx=(5, 5), pady=1, anchor='w')

        self.remove_btn = ttk.Button(self.root, text="Remove", width=7, state=DISABLED, command=lambda: self.remove_item_config())
        self.remove_btn.pack(side='left', anchor='center', padx=(77, 3))
        
        self.apply_btn = ttk.Button(self.root, text="Apply", width=5, command=lambda: self.write_game_config())
        self.apply_btn.pack(side='right', anchor='center', padx=(3, 77))

        self.process_selector.after(100, lambda: self.refresh_process_list())
        self.remove_btn.after(100, lambda: self.refresh_remove_btn())

        self.root.after(0, lambda: self.apply_style())
           
        if self.start_minimised == True:
            self.minimise()

        self.root.mainloop()
        
    def update_hidden_state(self):
        self.is_hidden = self.root.state() == "withdrawn"

    def validate_input(self, new_value: str):
        return new_value.isdigit() or new_value == ""
    
    def refresh_remove_btn(self):
        if self.validated_process == "" and ACTIVE in self.remove_btn.state():
            self.remove_btn.configure(state=DISABLED)
        
        self.remove_btn.after(1000, lambda: self.refresh_remove_btn())

    def refresh_process_list(self):
        self.all_open_processes.clear()
        self.load_processes()

        self.process_selector.configure(values=list(self.all_open_processes.values()))

        if self.validated_process != None and self.validated_process != "":
            self.process_selector.set(self.validated_process.replace('.exe', ''))

            process_resolution = self.game_settings_data[self.validated_process]
            w, h, rr = process_resolution['width'], process_resolution['height'], process_resolution['refresh_rate']

            self.x_var.set(w)
            self.y_var.set(h)
            self.rr_var.set(rr)

        self.process_selector.after(2500, lambda: self.refresh_process_list())

    def remove_item_config(self):
        self.game_settings_data.pop(self.validated_process)

        try:
            with open(self.game_settings_path, 'w') as file:
                file.write(dumps(self.game_settings_data, indent=1))
        except FileNotFoundError:
            raise r_error('Config file not found.')
        except JSONDecodeError:
            raise r_error('Config file is not valid JSON.')
        
        self.set_resolution(self.old_res['width'], self.old_res['height'], self.old_res['refresh_rate'])
        self.validated_process = ""

    def load_processes(self):
        """Returns a dictionary of all open processes."""
        for proc in psutil.process_iter(['name', 'exe']):
            try:
                if any(keyword.lower() in proc.exe().lower() for keyword in GAME_PATH_KEYWORDS):
                    name = proc.info['name']

                    processes = win32com.client.GetObject("winmgmts:").ExecQuery(
                        f"SELECT Name, Description FROM Win32_Process WHERE Name='{name}'"
                    )

                    try:
                        display_name = (processes[0].replace('.exe', ''))
                    except (AttributeError):
                        display_name = name.replace('.exe', '')

                    if name not in list(self.all_open_processes.keys()):
                        self.all_open_processes.update({ name: display_name })
            except (psutil.AccessDenied, psutil.NoSuchProcess):
                continue

    def header_colour(self):
        version = getwindowsversion()
        if version.major == 10 and version.build >= 22000:
            change_header_color(self.root, "#1c1c1c" if sv_ttk.get_theme(self.root) == "dark" else "#fafafa")
        elif version.major == 10:
            apply_style(self.root, "dark" if sv_ttk.get_theme(self.root) == "dark" else "normal")
        self.root.wm_attributes("-alpha", 0.99)
        self.root.wm_attributes("-alpha", 1)

    def apply_style(self):
        if self.root.state() == 'withdrawn':
            return
        
        self.is_dark_mode = self.is_windows_dark_mode()
        if self.is_dark_mode:
            sv_ttk.set_theme('dark', self.root)
        else:
            sv_ttk.set_theme('light', self.root)
            
        self.header_colour()

    def bring_to_front(self):
        self.root.deiconify()
        self.update_hidden_state()
        self.tray_menu()
        self.apply_style()
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(0, lambda: self.root.attributes('-topmost', False))

    def start_tray_icon(self):
        self.icon = Icon("ALR", self.image, "ALR", self.menu)
        self.icon.run()
        
    def toggle_start_minimised(self):
        self.start_minimised = not self.start_minimised
        self.config_data['show_minimised'] = self.start_minimised
        self.write_config()
        
    def toggle_startup(self):
        self.run_on_startup = not self.run_on_startup
        if self.run_on_startup == True and not is_startup_enabled():
            enable_startup()
        else: 
            disable_startup()
            
        self.config_data['run_on_startup'] = self.run_on_startup
            
        self.write_config()
        
    def check_startup_location(self):
        cur_exe_path = sys.executable
        try:
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run"
            ) as key:
                value, _ = winreg.QueryValueEx(key, APP_NAME)
                if cur_exe_path != value:
                    disable_startup()
                    sleep(1)
                    enable_startup()         
        except FileNotFoundError:
            pass

    def minimise(self):
        self.root.withdraw()
        self.update_hidden_state()
        self.tray_menu()
            
    def tray(self):
        self.tray_menu()
        if not self.icon_thread:  
            self.image = Image.open(resource_path('icon.ico'))
            self.icon_thread = Thread(target=self.start_tray_icon, daemon=True)
            self.icon_thread.start()
            
    def tray_menu(self):
        if hasattr(self, 'icon'):
            self.icon.menu = Menu(item('Start Minimised', self.toggle_start_minimised, checked=lambda _: self.start_minimised), item('Run on Startup', self.toggle_startup, checked=lambda _: self.run_on_startup), item('Show', self.open_window, enabled=lambda _: self.is_hidden), item('Exit', self.close)) # type: ignore
        else:
            self.menu = Menu(item('Start Minimised', self.toggle_start_minimised, checked=lambda _: self.start_minimised), item('Run on Startup', self.toggle_startup, checked=lambda _: self.run_on_startup), item('Show', self.open_window, enabled=lambda _: self.is_hidden), item('Exit', self.close)) # type: ignore
        
    def open_window(self):
        self.root.after(10, lambda: self.bring_to_front())

    def close(self):
        if hasattr(self, 'find_valid_process'):
            self.find_valid_process_event.set()
            sleep(0.1)
            self.find_valid_process_thread.join(timeout=1)
            
        if self.icon:
            self.icon.stop()
        
        sleep(1)
        
        self.root.destroy()

    def read_game_config(self):
        try:
            with open(self.game_settings_path, 'r') as file:
                return loads(file.read())
        except FileNotFoundError:
            raise r_error('Game config file not found.')
        except JSONDecodeError:
            raise r_error('Game config file is not valid JSON.')

    def write_game_config(self):
        try:
            with open(self.game_settings_path, 'w') as file:
                if not self.x_var.get() or not self.y_var.get() or not self.rr_var.get() or not self.process_selector.get():
                    raise r_error('Width, height and refresh rate cannot be empty.')

                for name, display_name in self.all_open_processes.items():
                    if self.process_selector.get() == display_name:
                        self.selected_process = name
                        break

                self.game_settings_data[self.selected_process] = {
                    'width': self.x_var.get(),
                    'height': self.y_var.get(),
                    'refresh_rate': self.rr_var.get()
                }

                file.write(dumps(self.game_settings_data, indent=1))
        except FileNotFoundError:
            raise r_error('Game config file not found.')
        except JSONDecodeError:
            raise r_error('Game config file is not valid JSON.')
        
    def read_config(self):
        try:
            with open(self.config_path, 'r') as file:
                return loads(file.read())
        except FileNotFoundError:
            raise r_error('Config file not found.')
        except JSONDecodeError:
            raise r_error('Config file is not valid JSON.')
        
    def write_config(self):
        try:
            with open(self.config_path, 'w') as file:
                file.write(dumps(self.config_data, indent=1))
        except FileNotFoundError:
            raise r_error('Config file not found.')
        except JSONDecodeError:
            raise r_error('Config file is not valid JSON.')

    def is_windows_dark_mode(self):
        try:
            reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            key = winreg.OpenKey(reg, r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize')
            value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
            return value == 0
        except FileNotFoundError:
            return False

    def get_resolution(self):
        devmode = DEVMODEW()
        devmode.dmSize = ctypes.sizeof(DEVMODEW)
        res = ctypes.windll.user32.EnumDisplaySettingsW(None, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode))
        if res:
            return { 'width': devmode.dmPelsWidth, 'height': devmode.dmPelsHeight, 'refresh_rate': devmode.dmDisplayFrequency}
        else:
            raise r_error('Failed to get current resolution.')

    def set_resolution(self, width, height, refresh_rate=60):
        if not hasattr(self, 'old_res'):
            self.old_res = self.get_resolution()

        devmode = DEVMODEW()
        devmode.dmSize = ctypes.sizeof(DEVMODEW)
        devmode.dmPelsWidth = width
        devmode.dmPelsHeight = height
        devmode.dmDisplayFrequency = refresh_rate
        devmode.dmFields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_DISPLAYFREQUENCY
        res = ctypes.windll.user32.ChangeDisplaySettingsW(ctypes.byref(devmode), 0)
        return res == win32con.DISP_CHANGE_SUCCESSFUL

    def change_res(self):
        self.old_res = self.get_resolution()

        if DISABLED in self.remove_btn.state():
            self.remove_btn.config(state=NORMAL)

        process_resolution = self.game_settings_data[self.validated_process]
        w, h, rr = process_resolution['width'], process_resolution['height'], process_resolution['refresh_rate']
        if self.old_res['width'] != w or self.old_res['height'] != h or self.old_res['refresh_rate'] != rr:
            if not self.set_resolution(w, h, rr):
                raise r_error('Failed to set resolution.')
            self.has_changed_res = True

    def find_valid_process(self):
        while not self.find_valid_process_event.is_set():
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] in list(self.valid_process_names):
                    self.is_valid_process_running = True
                    self.validated_process = proc.info['name']
                    break
            else: 
                self.is_valid_process_running = False
                self.validated_process = ""

            if self.is_valid_process_running and not self.has_changed_res:
                self.change_res()
            elif not self.is_valid_process_running and self.has_changed_res:
                if not self.set_resolution(self.old_res['width'], self.old_res['height'], self.old_res['refresh_rate']):
                    raise r_error('Failed to reset resolution.')
                
                self.has_changed_res = False
            elif self.is_valid_process_running and self.has_changed_res:
                current_resolution = self.get_resolution()
                process_resolution = self.game_settings_data[self.validated_process]
                w, h, rr = process_resolution['width'], process_resolution['height'], process_resolution['refresh_rate']
                if current_resolution['width'] != w and current_resolution['height'] != h and current_resolution['refresh_rate'] != rr:
                    self.change_res()
                
            for _ in range(25):
                if self.find_valid_process_event.is_set():
                    break
                
                self.find_valid_process_event.wait(0.1)

class r_error(Exception):
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message

class DEVMODEW(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        ("dmOrientation", wintypes.SHORT),
        ("dmPaperSize", wintypes.SHORT),
        ("dmPaperLength", wintypes.SHORT),
        ("dmPaperWidth", wintypes.SHORT),
        ("dmScale", wintypes.SHORT),
        ("dmCopies", wintypes.SHORT),
        ("dmDefaultSource", wintypes.SHORT),
        ("dmPrintQuality", wintypes.SHORT),
        ("dmColor", wintypes.SHORT),
        ("dmDuplex", wintypes.SHORT),
        ("dmYResolution", wintypes.SHORT),
        ("dmTTOption", wintypes.SHORT),
        ("dmCollate", wintypes.SHORT),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmDisplayFlags", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]

def resource_path(relative_path: str) -> str:
    if hasattr(sys, '_MEIPASS'):
        return path.join(sys._MEIPASS, relative_path) # type: ignore
    return path.join(path.dirname(__file__), f'assets\\{relative_path}')

if __name__ == '__main__':
    changer = ResolutionChanger()
